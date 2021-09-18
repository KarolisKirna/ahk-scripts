#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

;example file for my_variables.ahk file needed to set variables in notion_task.ahk script
;!!!!!! Do not delete any of the "" marks when pasting your values
responsible_relation_id := "your_responsible_relation_id_goes_here"
tags_relation_id := "your_tags_relation_id_goes_here"
notion_exe_dir = "your_notion_exe_dir_goes_here" ;probably no need to change as apps are installed at standard directories 
notion_api_version := "2021-05-13" ;get from notion api reference page https://developers.notion.com/reference/versioning
bearer_token := "your_bearer_token_goes_here"
work_task_database_id := "your_work_task_database_id_goes_here"
code_exe_path = "your_code_exe_path_goes_here" ;standard installation path for vs code