vsion="v1.11"
Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists("key.txt") Then
	Ret = Msgbox("Are you sure you wish to permanently delete your stored Vanity Addresses & Private Keys?" ,  vbExclamation+vbOkCancel , "FTCVanity " & vsion)
	If Ret = vbOK then
	Set objDEL = CreateObject("Scripting.FileSystemObject")
            objDEL.DeleteFile ("key.txt") 
	Window=Msgbox("All stored Vanity Addresses & Private Keys have been permanently deleted!" , 64 , "FTCVanity " & vsion)	
	End If
Else

Window=Msgbox("No stored Vanity Addresses or Private Keys found!" , 16 , "FTCVanity " & vsion)



End If
