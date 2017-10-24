Attribute VB_Name = "mdl_GetFactors"
Option Explicit

Sub Test()

    Dim tmp As Collection
    Set tmp = New Collection

    Dim tmp1 As Collection
    Set tmp1 = New Collection

    Set tmp = AllFactors(1499) 'Change number - will return all factors.
    Set tmp1 = BestFactors(1499) 'Change number - will return best factors.

    Debug.Assert False

End Sub

'Returns the factors of a whole number.
Public Function AllFactors(NumToFactor As Single) As Collection

    Dim Count As Integer
    Dim Factor As Single
    Dim y As Single
    Dim tmpCollection As Collection

    Set tmpCollection = New Collection

    Count = 0
    For y = 1 To NumToFactor
        Factor = NumToFactor Mod y
        If Factor = 0 Then
            tmpCollection.Add y
        End If
    Next y

    Set AllFactors = tmpCollection

End Function

'Returns the highest factors of a number.
Public Function BestFactors(NumToFactor As Single) As Collection

    Dim tmpFactors As Collection
    Dim FactorNums As Collection
    Dim x As Single, y As Single, z As Single
    Dim FirstFactor As Single

    Set tmpFactors = New Collection
    Set FactorNums = New Collection

    'Get all factors for the number.
    Set FactorNums = AllFactors(NumToFactor)

    'If the collection has 1 item then the NumToFactor is 1.
    'If there's 2 items then it's a prime number (1 and NumToFactor)
    If FactorNums.Count = 1 Or FactorNums.Count = 2 Then
        tmpFactors.Add FactorNums(FactorNums.Count)
    Else
        For x = FactorNums.Count - 1 To 1 Step -1
            If FactorNums(x) ^ 2 = NumToFactor Then
                tmpFactors.Add FactorNums(x)
                tmpFactors.Add FactorNums(x)
                Exit For
            Else
                For y = x To 1 Step -1
                    FirstFactor = FactorNums(y)
                    For z = y - 1 To 1 Step -1
                        If FirstFactor * FactorNums(z) = NumToFactor Then
                            tmpFactors.Add FirstFactor
                            tmpFactors.Add FactorNums(z)
                            Exit For
                        End If
                    Next z
                    If tmpFactors.Count = 2 Then Exit For
                Next y
            End If
            If tmpFactors.Count = 2 Then Exit For
        Next x
    End If

    Set BestFactors = tmpFactors

End Function
