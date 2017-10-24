Attribute VB_Name = "mdl_SentenceCase"
Option Explicit

Sub Tested()
    Call ProperCaps("HERE IS A LONG, UGLY UPPERCASE SENTENCE. PLEASE AMEND ME IMMEDIATELY." & vbCrLf & "next line! now")
End Sub

Function ProperCaps(strIn As String) As String
    Dim objRegex As Object
    Dim objRegMC As Object
    Dim objRegM As Object
    Set objRegex = CreateObject("vbscript.regexp")
    strIn = LCase$(strIn)
    With objRegex
        .Global = True
        .ignoreCase = True
         .Pattern = "(^|[\.\?\!\r\t]\s?)([a-z])"
        If .Test(strIn) Then
            Set objRegMC = .Execute(strIn)
            For Each objRegM In objRegMC
                Mid$(strIn, objRegM.firstindex + 1, objRegM.Length) = UCase$(objRegM)
            Next
        End If
        MsgBox strIn
    End With
End Function
