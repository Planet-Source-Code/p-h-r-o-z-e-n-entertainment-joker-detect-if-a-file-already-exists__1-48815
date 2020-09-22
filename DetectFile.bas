Attribute VB_Name = "MyModule"
-------------------------------------------------------------------
'Author: |P|h|r|o|z|e|n| Entertainment - Joker
'Posted: 9/27/03
'
'How to detect if a file already exists.
'-------------------------------------------------------------------
' To detect an existing file, use the function below:

Function FileExists%(fname$)
On Local Error Resume Next

Dim ff%
        ff% = FreeFile
        Open fname$ For Input As ff%

        If Err Then
        FileExists% = False
        Else
        FileExists% = True
        End If

        Close ff%

End Function

' Add this code to the appropriate event:

Success% = FileExists%("C:\vb\vb.exe") 'A full path and filename

' FileExists% returns True if file exists
If Success% = True Then
    MsgBox "This file already exists.", 48, File Error
End If

