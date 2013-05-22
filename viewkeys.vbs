Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists("key.txt") Then
Set objTextFile = objFSO.OpenTextFile("key.txt")
strLine = objTextFile.ReadAll
	Window=Msgbox("Below is a list of your generated Vanity Addresses & Private Keys:" & VBNewline & VBNewline & strLine ,64, "FTCVanity v1.09")
    Return=Msgbox("Would you like to open your generated Vanity Address & Private Keys list in Notepad.exe?" , vbInformation+vbYesNo , "FTCVanity v1.09")
	If Return = vbYes Then
	Dim oShell
Set oShell = WScript.CreateObject ("WScript.Shell")
oShell.run "notepad.exe key.txt"
Set oShell = Nothing
else
WScript.Quit()
End If
Else
	Window=Msgbox("No stored Vanity Addresses or Private Keys found!" , 16 , "FTCVanity v1.09")
End If
