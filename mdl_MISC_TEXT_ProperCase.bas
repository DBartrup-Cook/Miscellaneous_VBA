Attribute VB_Name = "mdl_ProperCase"
Option Explicit

'http://www.tek-tips.com/faqs.cfm?fid=5749
'BobStubbs 13/03/05

Public Function ProperCase(strOneLine As String, intChangeType As Integer) As String

'---------------------------------------------------------------
'- This function will convert a string to Proper Case          -
'- The initial letter of each word is capitalised.             -
'- It will also handle special names such as O', Mc and        -
'- hyphenated names                                            -
'- if intChangeType = 1, all text is converted to proper case. -
'- e.g. 'FRED' is converted to 'Fred'                          -
'- if intChangeType = 0, upper case text is not converted.     -
'- e.g. 'fred' becomes 'Fred', but 'FRED' remains unchanged.   -
'---------------------------------------------------------------

Dim I As Integer
Dim bChangeFlag As Boolean
Dim strResult As String

'----------------------------------------------------------
'- No characters in string - nothing to do                -
'----------------------------------------------------------
If Len(strOneLine) = 0 Then
    ProperCase = ""
    Exit Function
End If

'----------------------------------------------------------
'- Always set first letter to upper case                  -
'----------------------------------------------------------
strResult = UCase$(Left$(strOneLine, 1))

'----------------------------------------------------------
'- Now look at the rest of the string                     -
'----------------------------------------------------------
For I = 2 To Len(strOneLine)
    
'----------------------------------------------------------
'- If the previous letter triggered a capital, change     -
'- this letter to upper case                              -
'----------------------------------------------------------
    If bChangeFlag = True Then
        strResult = strResult & UCase$(Mid$(strOneLine, I, 1))
        bChangeFlag = False
'----------------------------------------------------------
'- In other cases change letter to lower case if required -
'----------------------------------------------------------
    Else
        If intChangeType = 1 Then
            strResult = strResult & LCase$(Mid$(strOneLine, I, 1))
        Else
            strResult = strResult & Mid$(strOneLine, I, 1)
        End If
    End If
    
'----------------------------------------------------------
'- Set change flag if a space, apostrophe or hyphen found -
'----------------------------------------------------------
    Select Case Mid$(strOneLine, I, 1)
    Case " ", "'", "-"
        bChangeFlag = True
    Case Else
        bChangeFlag = False
    End Select
Next I

'----------------------------------------------------------
'- Special handling for Mc at start of a name             -
'----------------------------------------------------------
    If Left$(strResult, 2) = "Mc" Then
        Mid$(strResult, 3, 1) = UCase$(Mid$(strResult, 3, 1))
    End If
    
    I = InStr(strResult, " Mc")
    If I > 0 Then
        Mid$(strResult, I + 3, 1) = UCase$(Mid$(strResult, I + 3, 1))
    End If
   
ProperCase = strResult

End Function
