Attribute VB_Name = "mdl_EnumerateFiles_Fast"
Option Explicit

Public Function EnumerateFiles(sDirectory As String, _
            Optional sFileSpec As String = "*", _
            Optional InclSubFolders As Boolean = True) As Variant
    
    EnumerateFiles = Filter(Split(CreateObject("WScript.Shell").Exec _
        ("CMD /C DIR """ & sDirectory & "*." & sFileSpec & """ " & _
        IIf(InclSubFolders, "/S ", "") & "/B /A:-D").StdOut.ReadAll, vbCrLf), ".")

End Function
    
