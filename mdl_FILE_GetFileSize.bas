Attribute VB_Name = "mdl_GetFileSize"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : GetFileSize
' Author    : Darren Bartrup-Cook
' Date      : 04/04/2014
' Purpose   : Returns the size of the file in bytes.
'             Returns -1 if the file doesn't exist.
' To Use    : FSize = GetFileSize("C:\Bartrup-Cook\mdl_ACCESS_ImportXL.bas")
'---------------------------------------------------------------------------------------
Public Function GetFileSize(FilePath As String) As Double
    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    If oFSO.FileExists(FilePath) Then
        GetFileSize = oFSO.GetFile(FilePath).Size
    Else
        GetFileSize = -1
    End If
End Function
