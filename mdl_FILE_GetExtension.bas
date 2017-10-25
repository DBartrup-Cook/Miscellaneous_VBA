Attribute VB_Name = "mdl_GetExtension"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : GetExt
' DateTime  :
' Author    :
' Purpose   : Returns the extension of a given file name.
'---------------------------------------------------------------------------------------
Public Function GetExt(FileName As String) As String
    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    GetExt = oFSO.GetExtensionName(FileName)
    Set oFSO = Nothing
End Function
