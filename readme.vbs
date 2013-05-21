Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists("README.txt") Then
Set objTextFile = objFSO.OpenTextFile("README.txt")
strLine = objTextFile.ReadAll
	Window=Msgbox(strLine ,64, "FTCVanity v1.07 README")
Else
	Window=Msgbox("Wot No README?" ,64, "FTCVanity v1.07")
End If
