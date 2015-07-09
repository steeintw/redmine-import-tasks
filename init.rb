Redmine::Plugin.register :issue_importer_xls do
  name 'Issue Importer Xls plugin'
  author 'Stee.Shen'
  description 'Import Excel Sheet to create Redmine issues, forked from Issue Importer Xls plugin by Aspire Software Solutions'
  url ''
  author_url ''
  version '0.0.1'

  if RUBY_VERSION >= "1.9"
    permission :excel_sheet, { :excel_sheet => [:index, :upload_sheet] }, :public => true
    # menu :project_menu, :polls, { :controller => 'polls', :action => 'index' }, :caption => 'Polls', :after => :activity, :param => :project_id

    # menu :application_menu, :issue_importer_xls, { :controller => 'excel_sheet', :action => 'index' },
    # 			:caption => 'Import Issues' ,:last => true
    menu :project_menu, :excel_sheet, { :controller => 'excel_sheet', :action => 'index' },
    			:caption => 'Import Issues' ,:last => true

    settings :default => {'task_id_column' => 0,'save_task_as' => 1,'task_tracker_column' => 1, 'task_subject_column' => 2, 'approver_column' => 4, 'notices_column' => 5 ,'start_date_column' => 6,
                          'due_date_column' => 7 ,'assignee_name_column' => 3 ,'task_status_column' => 8, 'task_priority_column' => 9, 'task_description_column' => 10,
                          'parent_task_id_column' => 11, 'related_tasks_column' => 12 }, :partial => 'settings/issue_importer_setting'

  end
end
