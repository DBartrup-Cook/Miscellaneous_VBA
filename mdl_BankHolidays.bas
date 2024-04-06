Attribute VB_Name = "mdl_BankHolidays"
Option Explicit

'This first procedure is specific to Excel.
Public Sub DisplayBankHolidays()

    Dim lYear As Long
    Dim BH As Collection
    Dim vBH As Variant
    Dim lRow As Long
    
    
    Set BH = New Collection
    
    lYear = Sheet1.Range("A1")
    
    Set BH = BankHolidays(lYear)
    lRow = 3
    
    For Each vBH In BH
        Sheet1.Cells(lRow, 1) = vBH
        lRow = lRow + 1
    Next vBH

End Sub

'This could needs improving - just haven't had time yet.
Public Function BankHolidays(lYear As Long) As Collection

    Dim colTemp As Collection
    Dim dDateInQuestion As Date
    Dim dTemp As Date
    Set colTemp = New Collection
    
    'New Years Day
    'If falls on a weekend then following Monday is BH.
    dDateInQuestion = DateSerial(lYear, 1, 1)
    If Weekday(dDateInQuestion, vbMonday) >= 6 Then
        dTemp = dDateInQuestion + 8 - Weekday(dDateInQuestion, vbMonday)
    Else
        dTemp = dDateInQuestion
    End If
    colTemp.Add dTemp, "NewYearsDay"
    
    'Easter
    'Easter is the Sunday so isn't added,
    'but Good Friday & Easter Monday are calculated from this date.
    dTemp = EasterDate(CInt(lYear))
    colTemp.Add dTemp - 2, "GoodFriday"
    colTemp.Add dTemp + 1, "EasterMonday"
    
    'Early May Bank Holiday.
    'First Monday in May.
    dDateInQuestion = DateSerial(lYear, 5, 1)
    If Weekday(dDateInQuestion, vbMonday) > 1 Then
        dTemp = dDateInQuestion + 8 - Weekday(dDateInQuestion, vbMonday)
    Else
        dTemp = dDateInQuestion
    End If
    colTemp.Add dTemp, "EarlyMayBankHoliday"
    
    'Spring Bank Holiday
    'Last Monday in May.
    dDateInQuestion = DateSerial(lYear, 6, 1)
    dTemp = dDateInQuestion - Weekday(dDateInQuestion, vbTuesday)
    colTemp.Add dTemp, "SpringBankHoliday"
    
    'Summer Bank Holiday
    dDateInQuestion = DateSerial(lYear, 9, 1)
    dTemp = dDateInQuestion - Weekday(dDateInQuestion, vbTuesday)
    colTemp.Add dTemp, "SummerBankHoliday"
    
    'Christmas Day
    'Records 25th as BH.
    'If 25th is Saturday, then following Monday is BH.
    'If 25th is Sunday, then following Tuesday is BH.
    dDateInQuestion = DateSerial(lYear, 12, 25)
    If Weekday(dDateInQuestion, vbMonday) >= 6 Then
        dTemp = dDateInQuestion + 8 - Weekday(dDateInQuestion, Weekday(dDateInQuestion, vbMonday) - 4)
        colTemp.Add dTemp, "ChristmasDay"
    Else
        colTemp.Add dDateInQuestion, "ChristmasDay"
    End If

    'Boxing Day
    'Records 26th as BH.
    'If 26th is Saturday, then following Monday is BH.
    'If 26th is Sunday, then following Tuesday is BH.
    dDateInQuestion = DateSerial(lYear, 12, 26)
    If Weekday(dDateInQuestion, vbMonday) >= 6 Then
        dTemp = dDateInQuestion + 8 - Weekday(dDateInQuestion, Weekday(dDateInQuestion, vbMonday) - 4)
        colTemp.Add dTemp, "BoxingDay"
    Else
        colTemp.Add dDateInQuestion, "BoxingDay"
    End If
    
    Set BankHolidays = colTemp

End Function

'---------------------------------------------------------------------------------------
' Procedure : EasterDate
' Author    : Chip Pearson
' Site      : http://www.cpearson.com/excel/Easter.aspx
' Purpose   : Calculates which date Easter Sunday is on.  Is good from 1900 to 2099.
'---------------------------------------------------------------------------------------
Public Function EasterDate(Yr As Integer) As Date
    Dim d As Integer
    d = (((255 - 11 * (Yr Mod 19)) - 21) Mod 30) + 21
    EasterDate = DateSerial(Yr, 3, 1) + d + (d > 48) + 6 - ((Yr + Yr \ 4 + _
            d + (d > 48) + 1) Mod 7)
End Function

