Attribute VB_Name = "mdl_NaturalDate"
Option Explicit

'--------------------------------------------------------------------------------------
' Procedure : NaturalDate
' Author    : Darren Bartrup-Cook
' Date      : 23/01/2014
' Purpose   : Returns a text string showing the formatted date with day ordinal suffix.
' Arguments : lDayName = 0 : No Day Name
'             lDayName = 1 : Full Day Name (Thursday)
'             lDayName = 2 : Shortened Day Name (Thu)
'             lMonthName = 1: Full Month Name (January)
'             lMonthName = 2: Shortened Month Name (Jan)
'             lYear = 1: Full Year (2014)
'             lYear = 2: Shortened Year ('14)
' To Use    : MsgBox NaturalDate(#1/2/2014#, 2, 2, 2) returns "Thu 2nd Jan '14"
'---------------------------------------------------------------------------------------
Public Function NaturalDate(dDate As Date, Optional lDayName As Long = 1, _
                                           Optional lMonthName As Long = 1, _
                                           Optional lYear As Long = 1) As String

    Dim sDayName As String, sMonthName As String, sYear As String, sOrdinalNumber As String
    
    Select Case lDayName
        Case 1
            sDayName = Format(dDate, "dddd")
        Case 2
            sDayName = Format(dDate, "ddd")
        Case Else
            sDayName = ""
    End Select
    
    Select Case lMonthName
        Case 2
            sMonthName = Format(dDate, "mmm")
        Case Else
            sMonthName = Format(dDate, "mmmm")
    End Select
    
    Select Case lYear
        Case 2
            sYear = Format(dDate, "'yy")
        Case Else
            sYear = Format(dDate, "yyyy")
    End Select
    
    sOrdinalNumber = Format(dDate, "d") & OrdinalSuffix(Day(dDate))
    
    NaturalDate = Trim(sDayName & " " & sOrdinalNumber & " " & sMonthName & " " & sYear)
    
End Function

'--------------------------------------------------------------------------------------
' Procedure : OrdinalSuffix
' Author    : Chip Pearson
' Date      : 06/11/2013
' Purpose   : Returns the suffix that can then be appended to a number to get an ordinal number.
'---------------------------------------------------------------------------------------
Function OrdinalSuffix(ByVal Num As Long) As String
    Dim N As Long
    Const cSfx = "stndrdthththththth" ' 2 char suffixes
    N = Num Mod 100
    If ((Abs(N) >= 10) And (Abs(N) <= 19)) _
            Or ((Abs(N) Mod 10) = 0) Then
        OrdinalSuffix = "th"
    Else
        OrdinalSuffix = Mid(cSfx, _
            ((Abs(N) Mod 10) * 2) - 1, 2)
    End If
End Function
