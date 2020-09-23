Attribute VB_Name = "Module2"
Public error As String, ans1 As Double, deg As Integer
Public Function compile() As Boolean
Dim i, j As Integer, r As String, mess As String
Form1.Label1.Caption = "Compiling....."
j = 1
deg = 0
Form3.Text1.Text = ""
Form3.Caption = "Compiling..."
mess = "Error = "
express(k + 1) = "$"
Do
Select Case express(j)
Case "("
    deg = deg + 1
 Case ")"
   j = j + 1
   Select Case express(j)
     Case "0" To "9"
      error = ""
      For s = 1 To j
      error = error & express(s)
      Next s
      error = mess & error & ": Operator Missing, ok"
      Form3.Show
      Form3.Text1.Text = Form3.Text1 & error & vbCr & vbLf
      i = 1
      Case Else
      j = j - 1
      End Select
     If deg = 0 Then
        error = ""
      For s = 0 To j
      error = error & express(s)
      Next s
     error = mess & error & ": ( " & " Open parathesis Missing, ok"
     Form3.Show
     Form3.Text1.Text = Form3.Text1 & error & vbCr & vbLf
     i = 1
   Else
    deg = deg - 1
  End If
 Case "/", "*", "^", "%", "Sin", "Cos", "Tan", "Sinh" _
  , "Cosh", "Tanh", "log", "In", "Alog", "_10^X" _
  , "X²", "X^3", "Sqr", "Cur", "XsrY", "_1/X", "!", "e^X", "ncr" _
  , "npr", ".", "Sinn", "Cosn", "Tann", "r", "g"
  j = j + 1
   Select Case exp1(j)
     Case "+", "-", "/", "*", "^", "%", "Sin", "Cos", "Tan", "Sinh" _
     , "Cosh", "Tanh", "log", "In", "Alog", "_10^X" _
     , "X²", "X^3", "Sqr", "Cur", "XsrY", "_1/X", "!", "e^X", "ncr" _
     , "npr", ".", "Sinn", "Cosn", "Tann", "r", "g"
     error = ""
     For s = 0 To j
     error = error & express(s)
     Next s
     error = mess & " " & error & ":  " & "invalid Operator, ok"
     Form3.Show
     Form3.Text1.Text = Form3.Text1 & error & vbCr & vbLf
     i = 1
     End Select
     j = j - 1
     If express(j) = "/" Then
     j = j + 1
     If express(j) = "0" Then
      error = ""
      For s = 0 To j
      error = error & express(s)
      Next s
      error = mess & " " & error & ": " & " Division by zero, ok"
      Form3.Show
      Form3.Text1.Text = Form3.Text1 & error & vbCr & vbLf
      i = 1
     End If
     j = j - 1
   End If
 Case "0" To "9", "."
 If express(j + 1) = "+" Or express(j + 1) = "-" Then
     If express(j + 2) = "$" Then
       error = ""
      For s = 0 To j + 1
      error = error & express(s)
      Next s
      error = mess & " " & error & "  " & "Operator is not Valid"
      Form3.Show
      Form3.Text1.Text = Form3.Text1 & error & vbCr & vbLf
      i = 1
      End If
  End If
 Case "e"
 exp1(j) = "2.718281828"
 Case "pi"
 exp1(j) = "3.14159265358979"
 Case "Ans"
 exp1(j) = ans1
 Case "A" To "Z"
 exp1(j) = Form1.sendvar(express(j))
  Case Else
 End Select
 j = j + 1
 Loop Until express(j) = "$"
   If deg <> 0 Then
      error = ""
      For s = 1 To j - 1
     error = error & express(s)
     Next s
     error = mess & error & ":  ) " & " Close parathesis Missing, ok"
     Form3.Show
    Form3.Text1.Text = Form3.Text1 & error & vbCr & vbLf
   i = 1
  End If
 top1 = -1
 If i = 1 Then
 compile = False
 Else
 compile = True
 End If
 End Function
Public Function isfileExist(fname As String) As Boolean
On Error GoTo label
Open fname For Input As #1
Close #1
isfileExist = True
Exit Function
label:
isfileExist = False
End Function

