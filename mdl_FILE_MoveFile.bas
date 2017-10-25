Attribute VB_Name = "mdl_MoveFile"
Option Explicit

'----------------------------------------------------------------------
' MoveFile
'
'   Moves the file from FromFile to ToFile.
'   Returns True if it was successful.
'----------------------------------------------------------------------
Public Function MoveFile(FromFile As String, ToFile As String) As Boolean

    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    oFSO.MoveFile FromFile, ToFile
    MoveFile = (Err.Number = 0)
    Err.Clear
    
End Function
