Attribute VB_Name = "Module1"
Global AlphaNumeric1(0 To 19) As String
Global AlphaNumeric2(1 To 9) As String
Global AlphaNumeric3(1 To 9) As String
Function AbH(Number As String)

Dim IsNegative As String
Dim DotPosition As Integer
Dim IntegerSegment As String
Dim DecimalSegment As String
Dim DotTxt, DecimalTxt As String

If Val(Number) > 0 Then IsNegative = "" Else IsNegative = "„‰›Ì "
DotPosition = InStr(1, Number, ".")

If Not (DotPosition) = 0 Then
    IntegerSegment = Left(Abs(Number), DotPosition - 1)
    DecimalSegment = Left(Right(Number, Len(Number) - DotPosition), 5)
    
If Val(IntegerSegment) <> 0 Then DotTxt = " „„Ì“ " Else DotTxt = ""

Select Case Len(DecimalSegment)

    Case 1
        DecimalTxt = " œÂ„ "
    Case 2
        DecimalTxt = " ’œ„ "
    Case 3
        DecimalTxt = " Â“«—„ "
    Case 4
        DecimalTxt = " œÂ Â“«—„ "
    Case 5
        DecimalTxt = " ’œ Â“«—„ "
        
End Select

    
   AbH = IsNegative & Horof(IntegerSegment) & DotTxt & Horof(DecimalSegment) & DecimalTxt
   
    
Exit Function

End If
    
    
    
AbH = IsNegative & Horof(Abs(Number))


End Function

Sub alphaset()
   Dim i%
   AlphaNumeric1(0) = ""
   AlphaNumeric1(1) = "Ìﬂ"
   AlphaNumeric1(2) = "œÊ"
   AlphaNumeric1(3) = "”Â"
   AlphaNumeric1(4) = "çÂ«—"
   AlphaNumeric1(5) = "Å‰Ã"
   AlphaNumeric1(6) = "‘‘"
   AlphaNumeric1(7) = "Â› "
   AlphaNumeric1(8) = "Â‘ "
   AlphaNumeric1(9) = "‰Â"
   AlphaNumeric1(10) = "œÂ"
   AlphaNumeric1(11) = "Ì«“œÂ"
   AlphaNumeric1(12) = "œÊ«“œÂ"
   AlphaNumeric1(13) = "”Ì“œÂ"
   AlphaNumeric1(14) = "çÂ«—œÂ"
   AlphaNumeric1(15) = "Å«‰“œÂ"
   AlphaNumeric1(16) = "‘«‰“œÂ"
   AlphaNumeric1(17) = "Â›œÂ"
   AlphaNumeric1(18) = "ÂÌÃœÂ"
   AlphaNumeric1(19) = "‰Ê“œÂ"
   
   
   AlphaNumeric2(1) = "œÂ"
   AlphaNumeric2(2) = "»Ì” "
   AlphaNumeric2(3) = "”Ì"
   AlphaNumeric2(4) = "çÂ·"
   AlphaNumeric2(5) = "Å‰Ã«Â"
   AlphaNumeric2(6) = "‘’ "
   AlphaNumeric2(7) = "Â› «œ"
   AlphaNumeric2(8) = "Â‘ «œ"
   AlphaNumeric2(9) = "‰Êœ"
   
   AlphaNumeric3(1) = "Ìﬂ’œ"
   AlphaNumeric3(2) = "œÊÌ” "
   AlphaNumeric3(3) = "”Ì’œ"
   AlphaNumeric3(4) = "çÂ«—’œ"
   AlphaNumeric3(5) = "Å«‰’œ"
   AlphaNumeric3(6) = "‘‘’œ"
   AlphaNumeric3(7) = "Â› ’œ"
   AlphaNumeric3(8) = "Â‘ ’œ"
   AlphaNumeric3(9) = "‰Â’œ"
    
   
End Sub


Function Horof(Number As String) As String
   alphaset
    Dim No As Currency, N As String
    
    On Error GoTo Horoferror
    
    No = CCur(Number)
    N = CStr(No)
    
    Select Case Len(N)
        Case 1 To 3:
                If N < 20 Then
                    Horof = AlphaNumeric1(N)
                ElseIf N < 100 Then
                    If N Mod 10 = 0 Then
                        Horof = AlphaNumeric2(N \ 10)
                    Else
                        Horof = AlphaNumeric2(N \ 10) & " Ê " & Horof(N Mod 10)
                    End If
                ElseIf N < 1000 Then
                    If N Mod 100 = 0 Then
                        Horof = AlphaNumeric3(N \ 100)
                    Else
                        Horof = AlphaNumeric3(N \ 100) & " Ê " & Horof(N Mod 100)
                    End If
                        
                End If
        Case 4 To 6:
                If (Right(N, 3)) = 0 Then
                   Horof = Horof(Left(N, Len(N) - 3)) & " Â“«— "
                Else
                    Horof = Horof(Left(N, Len(N) - 3)) & " Â“«— Ê " & Horof(Right(N, 3))
                End If
        Case 7 To 9:
                If (Right(N, 6)) = 0 Then
                   Horof = Horof(Left(N, Len(N) - 6)) & " „Ì·ÌÊ‰ "
                Else
                    Horof = Horof(Left(N, Len(N) - 6)) & " „Ì·ÌÊ‰ Ê " & Horof(Right(N, 6))
                End If
        Case Else:
                If (Right(N, 9)) = 0 Then
                   Horof = Horof(Left(N, Len(N) - 9)) & " „Ì·Ì«—œ "
                Else
                    Horof = Horof(Left(N, Len(N) - 9)) & " „Ì·Ì«—œ Ê " & Horof(Right(N, 9))
                End If
            
    End Select
    
    Exit Function
Horoferror:
    Horof = "#Error"
End Function







