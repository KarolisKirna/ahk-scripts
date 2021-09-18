dir = %TEMP%\test
If !FileExist(dir) {
 FileCreateDir, %dir%
 If ErrorLevel
  MsgBox, 48, Error, An error occurred when creating the directory.`n`n%dir%
 Else MsgBox, 64, Success, Directory was created.`n`n%dir%
} Else MsgBox, 64, Exists, Directory already exists.`n`n%dir%
;Making a test commmit