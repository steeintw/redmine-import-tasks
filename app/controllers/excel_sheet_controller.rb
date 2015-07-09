class ExcelSheetController < ApplicationController
  unloadable
  	# before_filter :find_project, :require_admin, :authorize,:only => :index

	def index

		@project = Project.find(params[:id])
		session[:project_id]=params[:id]

	end

	def save_configuration

	end

	def upload_sheet

		   uploaded_io = params[:file]

		   if uploaded_io.nil? || uploaded_io.tempfile.nil?

		   		flash[:notice] = 'Please Submit Excel File'
  				redirect_to :action => 'index', :id => session[:project_id]
  				return
		   end

		   unless File.exists?("#{Rails.root}/public/uploads")
		   		Dir::mkdir("#{Rails.root}/public/uploads")
		   end

		  FileUtils.cp  "#{uploaded_io.tempfile.to_path.to_s}", "#{Rails.root}/public/uploads/#{uploaded_io.original_filename}"

		 extname=File.extname("#{Rails.root}/public/uploads/#{uploaded_io.original_filename}")

		 case extname
		 #Microsoft Excel File
		 when ".xls"
		 	workbook = Roo::Excel.new  "#{Rails.root}/public/uploads/#{uploaded_io.original_filename}"
		 #Microsoft Excel Xml File
		 when ".xlsx"
		 	workbook =  Roo::Excelx.new  "#{Rails.root}/public/uploads/#{uploaded_io.original_filename}"
		 #ODF Spreadsheet/OpenOffice document
		 when ".ods"
		 	workbook =Roo::OpenOffice.new   "#{Rails.root}/public/uploads/#{uploaded_io.original_filename}"
		 else
		 	flash[:notice] = 'Please Submit Excel File'
  			redirect_to :action => 'index', :id => session[:project_id]
  			return
		 end

 		 workbook.default_sheet = workbook.sheets[0]

 		 headers = Hash.new
			workbook.row(1).each_with_index {|header,i|
				headers[header] = i
		 }

 		 project_name=workbook.cell(1,1)
 		 redmine_project = Project.find(session[:project_id])
 		 if !redmine_project
        	redmine_project = @redmine_project
      	 end
 		 excel_error_message="Excel File contains following error.<br>"
 		 excel_having_errors=false

 		 #get plugin configuration
 		 settings_conf=Setting.plugin_issue_importer_xls
 		 logger.info "#{settings_conf}"
 		 ((workbook.first_row + 1)..workbook.last_row).each do |row|

 		 	row_content=Array.new(workbook.row(row))
 		 	row_content.each {|content| logger.info "#{content} --- #{content.class.name}" }

 		 	if row_content[settings_conf['task_subject_column'].to_i].nil? || row_content[settings_conf['task_description_column'].to_i].nil?

 		 		excel_error_message.concat("Excel Row ##{row} does not contain task description.<br>")
 		 		excel_having_errors=true
 		 	end

 		 	#if row_content[settings_conf['start_date_column'].to_i].class.name != "Date" || row_content[settings_conf['due_date_column'].to_i].class.name != "Date"
 		 	#	excel_error_message.concat("Excel Row ##{row} does not contain valid Start Date/Due Date<br> ")
 		 	#	excel_having_errors=true
 		 	#end
 		 end



 		 unless excel_having_errors
       task_list = {}
 		 ((workbook.first_row + 1)..workbook.last_row).each do |row|

	 		 	#iterate through all rows
	 		 	row_content=Array.new(workbook.row(row))
	 		 	#Project Name/Task	Best Case	Worst Case	Average Case	Notes	Questions	Start Date	Due Date	Total(in weeks)	Asignee
	 		 	unless row_content[0]== l(:label_import_issue_task) || row_content[0] == l(:label_import_issue_design) || row_content[0] == l(:label_import_issue_development) || row_content[0] == l(:label_import_issue_documentation) || row_content[0] == l(:label_import_issue_testing)
          task_id = row_content[settings_conf['task_id_column'].to_i].to_i
          task_list[task_id] = {}
          #try find Tracker
          tracker = Tracker.find_by(name: row_content[settings_conf['task_tracker_column'].to_i])
          tracker = Tracker.find_by_name('Task') if tracker.nil? #default tracker
          #try find Assignee
          assignee = find_user_by_name(row_content[settings_conf['assignee_name_column'].to_i])
          #try find Approver
          approver = find_user_by_name(row_content[settings_conf['approver_column'].to_i])
          #try find Notices
          notices_str = row_content[settings_conf['notices_column'].to_i]
          noticeIds = []
          notices = notices_str.split(',')
          notices.each do |notice_name|
            notice_user = find_user_by_name(notice_name.strip)
            noticeIds << notice_user.id unless notice_user.nil?
          end
          #try find Status by name
          status = IssueStatus.find_by_name('New')
          status = IssueStatus.find_by_name(row_content[settings_conf['task_status_column'].to_i]) unless row_content[settings_conf['task_status_column'].to_i].nil?

          #try find Priority by name
          priority = IssuePriority.find_by_name('Normal')
          priority = IssuePriority.find_by_name(row_content[settings_conf['task_priority_column'].to_i]) unless row_content[settings_conf['task_priority_column'].to_i].nil?

          task_list[task_id][:parent_task_id] = row_content[settings_conf['parent_task_id_column'].to_i].to_i
          task_list[task_id][:related_tasks] = row_content[settings_conf['related_tasks_column'].to_i]

          #create new issue
	 		 	  issue = Issue.new
				  issue.author_id = User.current.id
				 	issue.project_id= redmine_project.id
          issue.tracker= tracker
				 	issue.subject=row_content[settings_conf['task_subject_column'].to_i]
          issue.assigned_to_id = assignee.id unless assignee.nil?
				 	issue.status_id= status.id
          issue.priority_id= priority.id
				 	issue.description=row_content[settings_conf['task_description_column'].to_i]
				 	issue.start_date=row_content[settings_conf['start_date_column'].to_i]
				 	issue.due_date=row_content[settings_conf['due_date_column'].to_i]
          #save custom fields
          f_id = Hash.new { |hash, key| hash[key] = nil }
          issue.available_custom_fields.each_with_index.map { |f,indx| f_id[f.name] = f.id }
          field_list = []
          unless approver.nil?
            if f_id.include? "Approver"
              field_id = f_id["Approver"].to_s
              field_list << Hash[field_id, approver.id]
            end
          end

          unless noticeIds.empty?
            if f_id.include? "Notice"
              field_id = f_id["Notice"].to_s
              field_list << Hash[field_id, noticeIds]
            end
          end
          issue.custom_field_values = field_list.reduce({}, :merge)
				 	#Save issue for project
			 		issue.save!
          task_list[task_id][:task_id] = issue.id
	 			end
	 		 end
        #second loop, update parent task id
        logger.info "task_list: #{task_list}"
        task_list.each do |key, task_rec|
          issue = Issue.find(task_rec[:task_id])
            unless issue.nil?
              # update parent task
              unless task_rec[:parent_task_id].eql? 0
                parent_task_rec = task_list[task_rec[:parent_task_id]]
                issue.parent_id = parent_task_rec[:task_id]
                issue.save!
              end
              #parse related tasks
              begin
              logger.info "related tasks: #{task_rec[:related_tasks]}"
              unless task_rec[:related_tasks].nil? or task_rec[:related_tasks].empty?
                task_rec[:related_tasks].split(';').each do |r_str|
                  par =  r_str.strip.split(':')
                  rel_str = par[0].strip
                  relation_type = rel_str == 'B' ? 'blocks' : 'precedes';
                  fake_ids_str = par[1].strip
                  fake_ids_str.split(',').each do |f_str|
                    fake_id = f_str.to_i
                    rel_to_rec = task_list[fake_id]
                    unless rel_to_rec.nil?
                      relation = IssueRelation.new
                      relation.issue_from = issue
                      relation.issue_to_id = rel_to_rec[:task_id]
                      relation.relation_type = relation_type
                      relation.delay = 1
                      relation.init_journals(User.current)
                      relation.save!
                    end
                  end
                end
              end
            rescue => e
              flash[:error]= "Tasks are created successfully, but failed to update Task Relations."
              logger.fatal "[Issue_Importer_Xls] Failed to update Task Relations. Exception: #{e}"
      	 	 	  redirect_to :action => 'index', :id => session[:project_id]
      	 	 	  return
            end
            end
        end
	 	 else

	 	 	flash[:error]=excel_error_message
	 	 	redirect_to :action => 'index', :id => session[:project_id]
	 	 	return

 		 end


		flash[:notice] = 'Issues successfully created'
  		redirect_to :action => 'index', :id => session[:project_id]

	end



	def generate_excel_sheet

		headers=Hash.new

    headers[params[:task_id_column]]="Id"
    headers[params[:task_tracker_column]]="Tracker"
		headers[params[:task_subject_column]]="Subject"
    headers[params[:assignee_name_column]]= "Assignee"
    headers[params[:approver_column]]= "Approver"
    headers[params[:notices_column]]= "Notice"
    headers[params[:start_date_column]]="Start Date(yyyy-MM-dd)"
		headers[params[:due_date_column]]= "Due Date(yyyy-MM-dd)"
    headers[params[:task_status_column]]="Status"
    headers[params[:task_priority_column]]="Priority"
		headers[params[:task_description_column]]="Description"
    headers[params[:parent_task_id_column]]="Parent Task Id"
    headers[params[:related_tasks_column]]="Related Tasks"




		column_headers=Array.new
		(0..12).each  do |i|

			if headers.has_key?(i.to_s)

				column_headers.push(headers.fetch(i.to_s))
			else
				column_headers.push("")
			end

		end
	    workbook = Spreadsheet::Workbook.new
	    sheet1 = workbook.create_worksheet :name => "Redmine Sample Sheet"

	    sheet1.row(0).replace column_headers
	    unless File.exists?("#{Rails.root}/public/uploads/exports")
		   	Dir::mkdir("#{Rails.root}/public/uploads/exports")
		end
			excel_sheet_file_path=["public", "uploads", "exports", "Redmine_Sample_Issue_Sheet.xls"].join("/")
		    export_file_path = [Rails.root,excel_sheet_file_path].join("/")
		    workbook.write export_file_path

		    render :text => ["public", "uploads", "exports", "Redmine_Sample_Issue_Sheet.xls"].join("/")
		return

	end

	def export_excel_sheet
			excel_sheet_file_path=["public", "uploads", "exports", "Redmine_Sample_Issue_Sheet.xls"].join("/")
			send_file excel_sheet_file_path, :content_type => "application/vnd.ms-excel", :disposition => 'attachment' ,:filename => "Redmine_Sample_Issue_Sheet.xls",:x_sendfile => true
	end

	def render_excel_sheet
	  excel_sheet_file_path=["public", "uploads", "exports", "Redmine_Sample_Issue_Sheet.xls"].join("/")
	  respond_to do |format|
      	 format.html
      	format.xls { send_data excel_sheet_file_path }
   	  end
	end

  private
  def find_user_by_name(name)
    return nil if name.nil? or name.empty?

    #first we try finding user by login
    assignee = User.find_by(login: name)
    #if not found, try finding by full name
    if assignee.nil?
      User.all.each do |user|
        if user.name.eql? name
          assignee = user
          break
        end
      end
    end
    assignee
  end
end
