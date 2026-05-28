' MarkdownPasteAddin Bridge Server - 后台启动脚本 (无窗口)
' 双击此文件在后台启动桥接服务

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
batPath = scriptDir & "\启动桥接服务.bat"

' Run hidden
WshShell.Run """" & batPath & """", 0, False
