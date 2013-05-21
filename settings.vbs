window=Msgbox("First Run Message:" & VBNewline & VBNewline & "Before proceeding FTCVanity needs to configure itself based on system performance, this is a one-time only process." & VBNewline & VBNewline & "A short test will now commence, this can take up to 30 seconds to complete." & VBNewline & VBNewline & "Choose OK to CONTINUE..." , vbExclamation, "FTCVanity_v1.07 Configuration")
Set objShell = CreateObject("Wscript.Shell")
Return = objShell.Run("%comspec% /k vanitygen.exe" & " -i " & "1", 1 , false )
Set objShell = CreateObject("Wscript.Shell")

DIM strComputer,strProcess, x

x = 1
strComputer = "." 
strProcess = "vanitygen.exe"
Do Until x > 400
delay = x
x = x + 1
WScript.Sleep 1
If isProcessRunning(strComputer,strProcess) then
x = 401
End If
Loop
delay = delay * 50
' Terminate cmd.exe & vanitygen.exe
Set oShell = CreateObject("WScript.Shell")
Set oWmg = GetObject("winmgmts:")

strWndprs = "select * from Win32_Process where name='vanitygen.exe'"
Set objQResult = oWmg.Execquery(strWndprs)

For Each objProcess In objQResult
On Error Resume Next
intReturn = objProcess.Terminate(1)

Next
  Set oShell = CreateObject("WScript.Shell")
Set oWmg = GetObject("winmgmts:")

strWndprs = "select * from Win32_Process where name='cmd.exe'"
Set objQResult = oWmg.Execquery(strWndprs)

For Each objProcess In objQResult

intReturn = objProcess.Terminate(1)

Next

' Creat syssettings.conf and write with delay data
Set objFSO=CreateObject("Scripting.FileSystemObject")

outFile="syssettings.conf"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write delay 
objFile.Close
Window=MsgBox("All Done!" & VBNewline & VBNewline & "FTCvanity is now fully configured and ready to use."  & VBNewline & VBNewline & "You will now return to the main menu." , vbInformation, "FTCVanity_v1.07 Configuration")

' Function to check if a process is running

function isProcessRunning(byval strComputer,byval strProcessName)

	Dim objWMIService, strWMIQuery

	strWMIQuery = "Select * from Win32_Process where name like '" & strProcessName & "'"
	
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" _ 
			& strComputer & "\root\cimv2") 


	if objWMIService.ExecQuery(strWMIQuery).Count > 0 then
		isProcessRunning = true
	else
		isProcessRunning = false
	end if

end function
