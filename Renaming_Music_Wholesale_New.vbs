
'Parameters
gstrFilename  = "Renaming_Wholesale.vbs"
gstrExtension = "mp3"
gboolLive     = "false"
gboolFirstN   = "true"
gboolLastN    = "true"
FirstN        = 18
LastN         = 24

strCurDir  = "\\tsclient\Pudnik\zzztemp\u\GenerationX\ValleyOfTheDolls"

'globals
strReplaceString = " ()-,&'+_[]"
gintCount = 0
strBuildup = ""
varToUser = ""
gintLength = Len(gstrExtension)

'FSO
Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
'strCurDir = Left(Wscript.ScriptFullName,InStr(1,Wscript.ScriptFullName,gstrFileName)-1)
' This has to be singular for when doing only a single directory
Set objFolder = objFSO.GetFolder(strCurDir)

'renaming loop
'For Each objFolder In objFolder.SubFolders
  For Each objFile In objFolder.Files
    If LCase(Right(objFile.Name, gintLength)) = gstrExtension Then
      On Error Resume Next
        If LCase(gboolFirstN) = "true" Then
          objFile.Name = Right(objFile.Name, Len(objFile.Name) - FirstN)
        End If
        If LCase(gboolLastN) = "true" Then
          objFile.Name = Left(objFile.Name, Len(objFile.Name) - LastN - 4) & "." & gstrExtension
        End If
        objFile.Name = ReplaceString(objFile.Name, strReplaceString, "",0)
        If LCase(gboolLive) = "true" Then
          objFile.Name = Left(objFile.Name, Len(objFile.Name) - 4) & "Live" & Right(objFile.Name, 4)
        End If
      If Err.Number <> 0 Then
        If Err.Number = 58 Then
          'skip it
        Else
          varToUser = MsgBox("Loop Runtime Error = " & Err.Number, vbExclamation, "Wholesale Renaming Script")
          WScript.Quit
        End If
      Else
        gintCount = gintCount + 1
        On Error Goto 0
      End If
    End If
  Next
'Next

'cleanup
Set objFSO = Nothing    

'Prepare output and display to user
strBuildup = "Done." & " " & gintCount & " files processed."
If gintCount < 1 Then
  strBuildup = strBuildup & vbCrLf & "...Maybe cause it's set to find " & UCase(gstrExtension) & "'s ?"
End If

varToUser = MsgBox(strBuildup, vbinformation, "Wholesale Renaming Script")

'ReplaceString function
Function ReplaceString(strin, strmapin, strmapout, fCaseSensitive)

' This function comes from Getz and Gilberts' VBA Dev Handbook, it gets rid
' of all the crap in filenames like spaces, foreign characters, etc

' (input_string,chars_youwantto_replace,replacementchar,boolCaseSensitivity)

  If Len(strmapin) > 0 Then
    If fCaseSensitive Then
        intmode = vbBinaryCompare
    Else
        intmode = vbTextCompare
    End If
    If Len(strmapout) > 0 Then
        strmapout = left(strmapout & String(Len(strmapin), right(strmapout, 1)), _
        Len(strmapin))
    End If
    For inti = 1 To Len(strin)
        strchar = Mid(strin, inti, 1)
        intPos = InStr(1, strmapin, strchar, intmode)
        If intPos > 0 Then
            strOut = strOut & Mid(strmapout, intPos, 1)
        Else
            strOut = strOut & strchar
        End If
    Next
End If

ReplaceString = strOut

End Function 'ReplaceString
