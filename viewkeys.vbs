Set objFSO = CreateObject("Scripting.FileSystemObject")
  If objFSO.FileExists("key.txt") Then
Set objTextFile = objFSO.OpenTextFile("key.txt")
strLine = objTextFile.ReadAll
	Window=Msgbox("Below is a list of your generated Vanity Addresses & Private Keys:" & VBNewline & VBNewline & strLine ,64, "FTCVanity v1.02")
Else
	Window=Msgbox("No stored Vanity Addresses or Private Keys found!" , 16 , "FTCVanity v1.02")
End If
