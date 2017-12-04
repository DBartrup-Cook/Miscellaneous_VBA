Attribute VB_Name = "mdl_GetInitials"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : GetInitials
' Author    : Darren Bartrup-Cook
' Date      : 25/11/2016
' Purpose   : Returns the first letter of each word in a text string, or after a
'             non-alphabetical character.
' Example   : "The quick brown fox jumps over the lazy dog" returns TQBFJOTLD
'             "The-T quick'Q brown?B fox/F jumps#j over*o the lazy dog" returns TTQQBBFFJJOOTLD
'---------------------------------------------------------------------------------------
Public Function GetInitials(FullName As String) As String

    Dim RegEx As Object
    Dim Ret As Object
    Dim RetItem As Object
    
    On Error GoTo ERR_HANDLE
    
    Set RegEx = CreateObject("VBScript.RegExp")
    With RegEx
        .IgnoreCase = True
        .Global = True
        '.Pattern = "(^|\s)([A-Z])" 'This would just return DB.
        .Pattern = "(\b[a-zA-Z])[a-zA-Z]* ?"
        Set Ret = .Execute(FullName)
        For Each RetItem In Ret
            GetInitials = GetInitials & UCase(RetItem.Submatches(0))
        Next RetItem
    End With
    
EXIT_PROC:
        On Error GoTo 0
        Exit Function
    
ERR_HANDLE:
        'Add your own error handling here.
        'DisplayError Err.Number, Err.Description, "mdl_GetInitials.GetInitials()"
        Resume EXIT_PROC
    
End Function
