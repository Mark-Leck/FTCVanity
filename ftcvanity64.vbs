Set objFSO = CreateObject("Scripting.FileSystemObject")
' First run setup one-time only script:
If objFSO.FileExists("syssettings.conf") Then
FileExists=True
Else
Set objShell = CreateObject("Wscript.Shell")
    objShell.Run "settings.vbs"
	WScript.Quit()
	End If


Dim objFileSystem, objInputFile
Set objFileSystem = CreateObject("Scripting.fileSystemObject")
Set objInputFile = objFileSystem.OpenTextFile("syssettings.conf", 1)
delay = objInputFile.ReadLine
objInputFile.Close

Dim objFileSystem2, objInputFile2
Set objFileSystem2 = CreateObject("Scripting.fileSystemObject")
Set objInputFile2 = objFileSystem2.OpenTextFile("version.conf", 1)
version = objInputFile2.ReadLine
objInputFile2.Close

If version=14 Then
name="Feathercoin"
example="6test" 
id="6"
Else
If version=48 Then
name="Litecoin"
example="Ltest" 
id="L"
Else
If version=0 Then
name="Bitcoin"
example="1test" 
id="1"
End If
End If
End If

Set objShell = CreateObject("Wscript.Shell")
strInput = InputBox( "Please enter the vanity pattern you would like to use to create your " & name & " address."  & VBNewline & VBNewline & _
"For faster address & key generation use no more than 5 or 6 characters, more characters can be used, but it may take a thousand years+ to generate :-)" & VBNewline & VBNewline & _
"Make sure you start your pattern with:" & id & " EG: " & example , "FTCVanity64 " & name )
If strInput = "" then
Wscript.Quit()
Else
		'Pre-start message....
		Window=Msgbox( chr(149) & "  Your Vanity " & name & " Address & Private Keys will now be generated." & VBNewline & VBNewline & CHR (149) & _
		"  Please be patient as this can take some time to complete." & VBNewline & VBNewline & CHR (149) & _
		"  To import the generated keys into the " & name & " QT client please refer to the README" & VBNewline & VBNewline & CHR (149) & _
		"If you have any questions or feedback please visit the Project page over @ the Feathercoin forum:  " & VBNewline & VBNewline & "http://forum.feathercoin.com/index.php?topic=489.0",64, "FTCVanity64 " & name )
Set objShell = CreateObject("Wscript.Shell")
Return = objShell.Run("%comspec% /k vanitygen64.exe" & " -X " & "14" & " -o tmp.txt" & " -i " & strInput  , 2 , false )



WScript.sleep delay
DIM strComputer,strProcess
strComputer = "." 
strProcess = "vanitygen64.exe"

if isProcessRunning(strComputer,strProcess) then
	MsgBox"Vanity Generator Started Successfully!" & VBNewline & VBNewline & "Choose OK to CONTINUE..." , vbOKOnly , "FTCVanity64"
else
	MsgBox"ERROR!" & VBNewline & VBNewline & CHR (149) & "  Unable to generate valid Feathercoin address from input pattern!" & VBNewline & CHR (149) & "  Pattern: " & strInput , 16 , "FTCVanity64 " & name
		Set oShell = CreateObject("WScript.Shell")
Set oWmg = GetObject("winmgmts:")

strWndprs = "select * from Win32_Process where name='vanitygen64.exe'"
Set objQResult = oWmg.Execquery(strWndprs)

For Each objProcess In objQResult

intReturn = objProcess.Terminate(1)

Next
	Set oShell = CreateObject("WScript.Shell")
Set oWmg = GetObject("winmgmts:")

strWndprs = "select * from Win32_Process where name='cmd.exe'"
Set objQResult = oWmg.Execquery(strWndprs)

For Each objProcess In objQResult

intReturn = objProcess.Terminate(1)

Next
'Added incase of a false positive and script actually ran and created tmp.txt
Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists("tmp.txt") Then
objFSO.DeleteFile "tmp.txt"
Msgbox "ERROR!! FALSE POSITIVE DETECTED!!" & VBNewline & VBNewline & CHR (149) & "  Please report this error message over on the project page @:" & VBNewline & VBNewline & CHR (149) & "  http://forum.feathercoin.com/index.php?topic=489.0" , 16 , "FTCVanity64 " & name
End If
	Wscript.Quit()
end if


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

Dim retval, fso, file
Set fso = CreateObject ("scripting.filesystemobject")

file = "tmp.txt"
retval = waitTilExists (file, true)

Function waitTilExists (ByVal file, withRepeat)
    ' Sleeps until the file exists
    ' The polling interval will increase gradually, but never rises above MAX_WAITTIME
    ' Times out after TIMEOUT msec. Will return false if caused by timeout.
    Dim waittime, totalwaittime, rep, doAgain
    Const INIT_WAITTIME = 20
    Const MAX_WAITTIME = 1000
    Const TIMEOUT = 100
    Const SLOPE = 1.1
    doAgain  = true
    Do While doAgain
        waittime = INIT_WAITTIME
        totalwaittime = 0
        Do While totalwaittime < TIMEOUT
            waittime = Int (waittime * SLOPE)
            If waittime>MAX_WAITTIME Then waittime=MAX_WAITTIME
            totalwaittime = totalwaittime + waittime
            WScript.sleep waittime
            If fso.fileExists (file) Then
                waitTilExists = true
				Exit Function
				End If
        Loop
        If withRepeat Then
            rep = MsgBox ("Still Generating, Please wait..." & VBNewline & VBNewline & "Choose OK to REFRESH or Choose CANCEL to QUIT..." , vbOkCancel , "FTCVanity64 " & name)
            doAgain = (rep = vbOk)
        Else
			
        End If
    Loop
	waitTilExists = false
   
Set oShell = CreateObject("WScript.Shell")
Set oWmg = GetObject("winmgmts:")

strWndprs = "select * from Win32_Process where name='vanitygen64.exe'"
Set objQResult = oWmg.Execquery(strWndprs)

For Each objProcess In objQResult

intReturn = objProcess.Terminate(1)

Next
	Set oShell = CreateObject("WScript.Shell")
Set oWmg = GetObject("winmgmts:")

strWndprs = "select * from Win32_Process where name='cmd.exe'"
Set objQResult = oWmg.Execquery(strWndprs)

For Each objProcess In objQResult

intReturn = objProcess.Terminate(1)

Next
	MsgBox"Vanity Generator Terminated!" , 16 , "FTCVanity64 " & name
	Wscript.Quit()
End Function

Window=MsgBox("All Done!" & VBNewline & VBNewline & "Choose OK to Continue..." ,64, "FTCVanity64 " & name)

	Set oShell = CreateObject("WScript.Shell")
Set oWmg = GetObject("winmgmts:")

strWndprs = "select * from Win32_Process where name='cmd.exe'"
Set objQResult = oWmg.Execquery(strWndprs)

For Each objProcess In objQResult

intReturn = objProcess.Terminate(1)

Next

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("tmp.txt")

strLine = objTextFile.ReadAll
Window=Msgbox( "Your New " & name & " Vanity Address & Private Key:" & VBNewline & VBNewline & strLine ,64, "FTCVanity64 " & name)
objTextFile.Close

Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists("key.txt") Then
Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run("%comspec% /K ftcvanity.bat"), 0, True
objFSO.DeleteFile "tmp.txt"
else
objFSO.CopyFile "tmp.txt" , "key.txt" 
objFSO.DeleteFile "tmp.txt"

End If

End If










