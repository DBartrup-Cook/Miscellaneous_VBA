Attribute VB_Name = "mdl_GetSystemNames"
Option Explicit

Private Declare Function api_GetUserName Lib "advapi32.dll" _
        Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function api_GetComputerName Lib "Kernel32" _
        Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
                    
Public Enum SystemData
    ComputerUser = &O1
    ComputerName = &O2
End Enum

'---------------------------------------------------------------------------------------
' Procedure : GetSystemNames
' Purpose   : The Windows Networking name of the user or computer.
'---------------------------------------------------------------------------------------
Public Function GetSystemNames(UOrC As SystemData) As String
    On Error Resume Next
    
    Dim NBuffer As String
    Dim Buffsize As Long
    Dim Wok As Long
    
    Buffsize = 256
    NBuffer = Space$(Buffsize)
    
    Select Case UOrC
        Case ComputerUser
            Wok = api_GetUserName(NBuffer, Buffsize)
            GetSystemNames = Trim$(NBuffer)
        Case ComputerName
            Wok = api_GetComputerName(NBuffer, Buffsize)
            GetSystemNames = Trim$(NBuffer)
    End Select
    
    On Error GoTo 0
End Function

