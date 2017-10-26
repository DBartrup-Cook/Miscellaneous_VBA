Attribute VB_Name = "mdl_ExtractPostCode"
Option Explicit

Public Sub Test()

    Debug.Print ValidatePostCode("SomeAddress, Somewhere, Some Street, NG15 0DT")
    Debug.Print ValidatePostCode("SomeAddress, Somewhere Other place, CM2 0RG Some Street")

End Sub

Public Function ValidatePostCode(strData As String) As Variant
    
    Dim RE As Object, REMatches As Object
    
    Dim UKPostCode As String
    
    'Pattern could probably be improved.
    UKPostCode = "(?:(?:A[BL]|B[ABDHLNRST]?|C[ABFHMORTVW]|D[ADEGHLNTY]|E[CHNX]?|F[KY]|G[LUY]?|" _
                & "H[ADGPRSUX]|I[GMPV]|JE|K[ATWY]|L[ADELNSU]?|M[EKL]?|N[EGNPRW]?|O[LX]|P[AEHLOR]|R[GHM]|S[AEGKLMNOPRSTWY]?|" _
                & "T[ADFNQRSW]|UB|W[ACDFNRSV]?|YO|ZE)\d(?:\d|[A-Z])? \d[A-Z]{2})"
     
    Set RE = CreateObject("VBScript.RegExp")
    With RE
        .MultiLine = False
        .Global = False
        .IgnoreCase = True
        .Pattern = UKPostCode
    End With
     
    Set REMatches = RE.Execute(strData)
    If REMatches.Count = 0 Then
        ValidatePostCode = CVErr(xlErrValue)
    Else
        ValidatePostCode = REMatches(0)
    End If
     
End Function

