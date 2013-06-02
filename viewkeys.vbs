vsion="v1.11"
Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists("key.txt") Then
Set objTextFile = objFSO.OpenTextFile("key.txt")
strLine = objTextFile.ReadAll
	Window=Msgbox("Below is a list of your generated Vanity Addresses & Private Keys:" & VBNewline & VBNewline & strLine ,64, "FTCVanity " & vsion)
    Return=Msgbox("Would you like to open your generated Vanity Address & Private Keys list in Notepad.exe?" , vbQuestion+vbYesNo , "FTCVanity " & vsion)
	If Return = vbYes Then
	Dim oShell
Set oShell = WScript.CreateObject ("WScript.Shell")
oShell.run "notepad.exe key.txt"
Set oShell = Nothing
else
WScript.Quit()
End If
Else
	Window=Msgbox("No stored Vanity Addresses or Private Keys found!" , 16 , "FTCVanity " & vsion)
End If
