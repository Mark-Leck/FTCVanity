' FTCVanity by UKMark
' Basic GUI front-end for vanitygen 
' vanitygen.exe by samr7
strInput = VanityInput( "Please enter the vanity pattern you would like to use."  & VBNewline & VBNewline & _
"For faster address & key generation use no more than 5 or 6 characters, more characters can be used, but it may take a thousand years+ to generate :-)" & VBNewline & VBNewline & _
"Make sure you start your pattern with a number 6! EG:6test" )
Set objShell = wscript.CreateObject("wscript.Shell")
objShell.Run"%comspec% /k \FTCVanity\vanitygen.exe" & " -X 14 -i " & strInput, 1, true

Function VanityInput( myPattern )


    ' Check if the script runs in CSCRIPT.EXE
    If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
        ' If so, use StdIn and StdOut
        WScript.StdOut.Write myPattern & " "
        VanityInput = WScript.StdIn.ReadLine
  	
    Else
        ' If not, use InputBox
        VanityInput = InputBox( myPattern , "FTCVanity v1.00"  )
	' Test for Cancel or Empty Input.
	If VanityInput = "" Then  
    ' Quit
	WScript.quit()

	Else
 
		'Pre-start message....
		Window=Msgbox( chr(149) & "  Your Vanity Feathercoin Address & Private Keys will now be generated." & VBNewline & VBNewline & CHR (149) & _
		"  Please be patient as this can take some time to complete." & VBNewline & VBNewline & CHR (149) & _
		"  To import the generated keys into the Feathercoin client please refer to \FTCVanity\README or the Project page over @ the Feathercoin forum:  " & VBNewline & "http://forum.feathercoin.com/index.php?topic=489.0",64, "FTCVanity v1.00")
		
		
	End If
	
	End If
	
End Function

'objShell.run ^^




