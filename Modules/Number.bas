Attribute VB_Name = "Number"
Function roundDown(amount As Double, digits As Integer) As Double
    roundDown = CLng((amount + (1 / (10 ^ (digits + 1)))) * (10 ^ digits)) / (10 ^ digits)
End Function

Function roundUp(amount As Double, digits As Integer) As Double
    roundUp = roundDown(amount + (5 / (10 ^ (digits + 1))), digits)
End Function

Function roundDownAuto(a As Double) As Double
    Dim i As Integer
    For i = 0 To 17
        If Abs(a * 10) > WorksheetFunction.Power(10, -(i - 1)) Then
            If a > 0 Then
                roundDownAuto = roundDown(a, i)
            Else
                roundDownAuto = roundUp(a, i)
            End If
        Exit Function
        End If
    Next
End Function

