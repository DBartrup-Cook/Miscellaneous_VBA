Attribute VB_Name = "mdl_FileExists"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : FileExists
' Author    : Darren Bartrup-Cook
' Date      : 04/04/2014
' Purpose   : Returns the TRUE if the file exists.
' To Use    : bFile = FileExists("S:\Bartrup-CookD\Customer Services Phone Reports")
'---------------------------------------------------------------------------------------
Public Function FileExists(ByVal FileName As String) As Boolean
    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    FileExists = oFSO.FileExists(FileName)
End Function
