Attribute VB_Name = "mdl_RandomInteger"
Option Explicit

Public Function RandomInteger(Optional Low As Long = 1, Optional High As Long = 100) As Long
    RandomInteger = Int((High - Low + 1) * Rnd() + Low)
End Function
