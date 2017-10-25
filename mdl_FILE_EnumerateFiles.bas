Attribute VB_Name = "mdl_EnumerateFiles"
Option Explicit

'//Places all file names with FileSpec extension into a collection.
'//To Use: EnumerateFiles "S:\Bartrup-CookD\Trackers", "*.xls", colFiles

Sub EnumerateFiles(ByVal sDirectory As String, _
    ByVal sFileSpec As String, _
    ByRef cCollection As Collection)

    Dim sTemp As String
    
    sTemp = Dir$(sDirectory & sFileSpec)
    Do While Len(sTemp) > 0
        cCollection.Add sDirectory & sTemp
        sTemp = Dir$
    Loop
End Sub

