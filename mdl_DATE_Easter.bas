Attribute VB_Name = "mdl_Easter"
Option Explicit

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
