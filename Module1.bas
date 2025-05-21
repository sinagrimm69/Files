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

If Val(Number) > 0 Then IsNegative = "" Else IsNegative = "���� "
DotPosition = InStr(1, Number, ".")

If Not (DotPosition) = 0 Then
    IntegerSegment = Left(Abs(Number), DotPosition - 1)
    DecimalSegment = Left(Right(Number, Len(Number) - DotPosition), 5)
    
If Val(IntegerSegment) <> 0 Then DotTxt = " ���� " Else DotTxt = ""

Select Case Len(DecimalSegment)

    Case 1
        DecimalTxt = " ��� "
    Case 2
        DecimalTxt = " ��� "
    Case 3
        DecimalTxt = " ����� "
    Case 4
        DecimalTxt = " �� ����� "
    Case 5
        DecimalTxt = " �� ����� "
        
End Select

    
   AbH = IsNegative & Horof(IntegerSegment) & DotTxt & Horof(DecimalSegment) & DecimalTxt
   
    
Exit Function

End If
    
    
    
AbH = IsNegative & Horof(Abs(Number))


End Function

Sub alphaset()
   Dim i%
   AlphaNumeric1(0) = ""
   AlphaNumeric1(1) = "��"
   AlphaNumeric1(2) = "��"
   AlphaNumeric1(3) = "��"
   AlphaNumeric1(4) = "����"
   AlphaNumeric1(5) = "���"
   AlphaNumeric1(6) = "��"
   AlphaNumeric1(7) = "���"
   AlphaNumeric1(8) = "���"
   AlphaNumeric1(9) = "��"
   AlphaNumeric1(10) = "��"
   AlphaNumeric1(11) = "�����"
   AlphaNumeric1(12) = "������"
   AlphaNumeric1(13) = "�����"
   AlphaNumeric1(14) = "������"
   AlphaNumeric1(15) = "������"
   AlphaNumeric1(16) = "������"
   AlphaNumeric1(17) = "����"
   AlphaNumeric1(18) = "�����"
   AlphaNumeric1(19) = "�����"
   
   
   AlphaNumeric2(1) = "��"
   AlphaNumeric2(2) = "����"
   AlphaNumeric2(3) = "��"
   AlphaNumeric2(4) = "���"
   AlphaNumeric2(5) = "�����"
   AlphaNumeric2(6) = "���"
   AlphaNumeric2(7) = "�����"
   AlphaNumeric2(8) = "�����"
   AlphaNumeric2(9) = "���"
   
   AlphaNumeric3(1) = "����"
   AlphaNumeric3(2) = "�����"
   AlphaNumeric3(3) = "����"
   AlphaNumeric3(4) = "������"
   AlphaNumeric3(5) = "�����"
   AlphaNumeric3(6) = "����"
   AlphaNumeric3(7) = "�����"
   AlphaNumeric3(8) = "�����"
   AlphaNumeric3(9) = "����"
    
   
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
                        Horof = AlphaNumeric2(N \ 10) & " � " & Horof(N Mod 10)
                    End If
                ElseIf N < 1000 Then
                    If N Mod 100 = 0 Then
                        Horof = AlphaNumeric3(N \ 100)
                    Else
                        Horof = AlphaNumeric3(N \ 100) & " � " & Horof(N Mod 100)
                    End If
                        
                End If
        Case 4 To 6:
                If (Right(N, 3)) = 0 Then
                   Horof = Horof(Left(N, Len(N) - 3)) & " ���� "
                Else
                    Horof = Horof(Left(N, Len(N) - 3)) & " ���� � " & Horof(Right(N, 3))
                End If
        Case 7 To 9:
                If (Right(N, 6)) = 0 Then
                   Horof = Horof(Left(N, Len(N) - 6)) & " ������ "
                Else
                    Horof = Horof(Left(N, Len(N) - 6)) & " ������ � " & Horof(Right(N, 6))
                End If
        Case Else:
                If (Right(N, 9)) = 0 Then
                   Horof = Horof(Left(N, Len(N) - 9)) & " ������� "
                Else
                    Horof = Horof(Left(N, Len(N) - 9)) & " ������� � " & Horof(Right(N, 9))
                End If
            
    End Select
    
    Exit Function
Horoferror:
    Horof = "#Error"
End Function







