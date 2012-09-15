SET winshell=WScript.CreateObject("WScript.Shell")
srv="http://sql-2/projectserver/"
'usr=""
'pwd=""
winshell.Run "WINPROJ.EXE /s " & srv' & " /u " & usr & " /p " & pwd
SET oP= CreateObject("MSProject.Application")
oP.Visible=true
oP.Macro("INVESTIMPORT")
oP.Quit
