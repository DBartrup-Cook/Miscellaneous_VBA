Attribute VB_Name = "mdl_PreviousMondayDate"
Option Explicit

Public Function PreviousMonday(CurrentDate As Date) As Date
    PreviousMonday = CurrentDate - Weekday(CurrentDate - 2)
End Function
