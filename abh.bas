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

If Val(Number) > 0 Then IsNegative = "" Else IsNegative = "ãäÝí "
DotPosition = InStr(1, Number, ".")

If Not (DotPosition) = 0 Then
    IntegerSegment = Left(Abs(Number), DotPosition - 1)
    DecimalSegment = Left(Right(Number, Len(Number) - DotPosition), 5)
    
If Val(IntegerSegment) <> 0 Then DotTxt = " ããíÒ " Else DotTxt = ""

Select Case Len(DecimalSegment)

    Case 1
        DecimalTxt = " Ïåã "
    Case 2
        DecimalTxt = " ÕÏã "
    Case 3
        DecimalTxt = " åÒÇÑã "
    Case 4
        DecimalTxt = " Ïå åÒÇÑã "
    Case 5
        DecimalTxt = " ÕÏ åÒÇÑã "
        
End Select

    
   AbH = IsNegative & Horof(IntegerSegment) & DotTxt & Horof(DecimalSegment) & DecimalTxt
   
    
Exit Function

End If
    
    
    
AbH = IsNegative & Horof(Abs(Number))


End Function

Sub alphaset()
   Dim i%
   AlphaNumeric1(0) = ""
   AlphaNumeric1(1) = "íß"
   AlphaNumeric1(2) = "Ïæ"
   AlphaNumeric1(3) = "Óå"
   AlphaNumeric1(4) = "åÇÑ"
   AlphaNumeric1(5) = "äÌ"
   AlphaNumeric1(6) = "ÔÔ"
   AlphaNumeric1(7) = "åÝÊ"
   AlphaNumeric1(8) = "åÔÊ"
   AlphaNumeric1(9) = "äå"
   AlphaNumeric1(10) = "Ïå"
   AlphaNumeric1(11) = "íÇÒÏå"
   AlphaNumeric1(12) = "ÏæÇÒÏå"
   AlphaNumeric1(13) = "ÓíÒÏå"
   AlphaNumeric1(14) = "åÇÑÏå"
   AlphaNumeric1(15) = "ÇäÒÏå"
   AlphaNumeric1(16) = "ÔÇäÒÏå"
   AlphaNumeric1(17) = "åÝÏå"
   AlphaNumeric1(18) = "åíÌÏå"
   AlphaNumeric1(19) = "äæÒÏå"
   
   
   AlphaNumeric2(1) = "Ïå"
   AlphaNumeric2(2) = "ÈíÓÊ"
   AlphaNumeric2(3) = "Óí"
   AlphaNumeric2(4) = "åá"
   AlphaNumeric2(5) = "äÌÇå"
   AlphaNumeric2(6) = "ÔÕÊ"
   AlphaNumeric2(7) = "åÝÊÇÏ"
   AlphaNumeric2(8) = "åÔÊÇÏ"
   AlphaNumeric2(9) = "äæÏ"
   
   AlphaNumeric3(1) = "íßÕÏ"
   AlphaNumeric3(2) = "ÏæíÓÊ"
   AlphaNumeric3(3) = "ÓíÕÏ"
   AlphaNumeric3(4) = "åÇÑÕÏ"
   AlphaNumeric3(5) = "ÇäÕÏ"
   AlphaNumeric3(6) = "ÔÔÕÏ"
   AlphaNumeric3(7) = "åÝÊÕÏ"
   AlphaNumeric3(8) = "åÔÊÕÏ"
   AlphaNumeric3(9) = "äåÕÏ"
    
   
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
                        Horof = AlphaNumeric2(N \ 10) & " æ " & Horof(N Mod 10)
                    End If
                ElseIf N < 1000 Then
                    If N Mod 100 = 0 Then
                        Horof = AlphaNumeric3(N \ 100)
                    Else
                        Horof = AlphaNumeric3(N \ 100) & " æ " & Horof(N Mod 100)
                    End If
                        
                End If
        Case 4 To 6:
                If (Right(N, 3)) = 0 Then
                   Horof = Horof(Left(N, Len(N) - 3)) & " åÒÇÑ "
                Else
                    Horof = Horof(Left(N, Len(N) - 3)) & " åÒÇÑ æ " & Horof(Right(N, 3))
                End If
        Case 7 To 9:
                If (Right(N, 6)) = 0 Then
                   Horof = Horof(Left(N, Len(N) - 6)) & " ãíáíæä "
                Else
                    Horof = Horof(Left(N, Len(N) - 6)) & " ãíáíæä æ " & Horof(Right(N, 6))
                End If
        Case Else:
                If (Right(N, 9)) = 0 Then
                   Horof = Horof(Left(N, Len(N) - 9)) & " ãíáíÇÑÏ "
                Else
                    Horof = Horof(Left(N, Len(N) - 9)) & " ãíáíÇÑÏ æ " & Horof(Right(N, 9))
                End If
            
    End Select
    
    Exit Function
Horoferror:
    Horof = "#Error"
End Function







