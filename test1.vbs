' 파일을 복사하는 VB스크립트

Option Explicit
Const ForReading = 1
strFolder = "C:\Users1"
strDestination = "C:\Users1\sim"
Set fso = CreateObject("Scripting.FileSystemObject")
Set objTextFile = fso.OpenTextFile _ ("c:\userss\file.txt", ForReading)
Do Until objTextFile.AtEndOfStream strFile = objTextFile.ReadLine Wscript.Echo strFile Loop
objTextFile.Close


sourceFile = fso.GetAbsolutePathName(strFile)
destFolder = fso.GetAbsolutePathName(strDestination)



 
Set objShell = CreateObject("Shell.Application")
Set FilesInZip=objShell.NameSpace(sourceFile).Items()
objShell.NameSpace(strDestination).copyHere FilesInZip, 16
 
Set fso = Nothing
Set objShell = Nothing
Set FilesInZip = Nothing

