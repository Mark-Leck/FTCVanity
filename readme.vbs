vsion="v1.11"
Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists("README.txt") Then
Set objTextFile = objFSO.OpenTextFile("README.txt")
strLine = objTextFile.ReadAll
	Window=Msgbox(strLine ,64, "FTCVanity " & vsion & " README")
Else
	Window=Msgbox("Wot No README?" ,64, "FTCVanity " & vsion)
End If
