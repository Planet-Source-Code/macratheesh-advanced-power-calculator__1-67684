Attribute VB_Name = "Module1"
Public flagstore As Boolean
Public stacks(100) As String, top1 As Integer
Public sign, temp As String, store, store1, d, z, e, exp11(10) As String
Public num(50) As String, stack(50) As String, exp1(100) As String
Public i As Integer, length, large1, array1(100) As String
Public b1, j As Integer, w1 As Integer, q, n, k As Integer, m, s, ans As Double
Public x1, express(100) As String, st
Public flag1 As Boolean, dd As String
Public a As Boolean, b As Boolean, C As Boolean
Public Function execute() As String
Dim i1 As Integer, j1 As Integer
Dim h1 As Integer, r As String
Dim d As String
exp1(k + 1) = "$"
i1 = 0
Do
 push (exp1(i1))
  If stacks(top1) = ")" Then
  j1 = -1
  Do
    d = pop()
    j1 = j1 + 1
    array1(j1) = d
  Loop Until stacks(top1) = "("
   array1(j1 + 1) = "$"
   q = j1 + 1
   r = pop()
   reverse
   On Error GoTo sk
   push (evaluvate())
sk:
 End If
i1 = i1 + 1
Loop Until stacks(top1) = "$"
execute = stacks(0)
End Function
Public Sub reverse()
Dim ad(100) As String
Dim i, j
i = q - 1
j = 0
Do
  ad(j) = array1(i)
  i = i - 1
  j = j + 1
 Loop Until array1(i) = ")"
i = 0
Do
  array1(i) = ad(i)
  i = i + 1
 Loop Until i = q
 q = i - 1
 End Sub

Public Sub clear1()
If b1 = 1 Then
Form1.Text1.Text = ""
exp1(0) = "("
b1 = 0
k = 0
top1 = -1
b = True
End If
End Sub
Public Sub display(txt As String)
x1 = 0
temp = Form1.Text1.Text
If flag1 = True Then
store = temp & " " & txt
flag1 = False
Else
store = temp & txt
End If
Form1.Text1.Text = store
k = k + 1
exp1(k) = txt
express(k) = txt
sign = k
 flagstore = False
 If check(txt) Then
   flag1 = True
 End If
  If exp1(k) = "(" Then
  k = k - 1
   Select Case exp1(k)
    Case "0" To "9", ")"
    k = k + 1
    exp1(k) = "*"
    express(k) = "*"
    k = k + 1
    exp1(k) = txt
    express(k) = txt
    Case Else
    k = k + 1
   End Select
  End If
 End Sub
Public Function execlog() As Integer
Dim large2 As Integer, w As String, j, a As Double, b As String, C As Double
Dim d As Double, stock(100) As String, s
j = 0
large2 = 0
Do While large2 <= q - 1
Select Case array1(large2)
Case "log", "In", "Sin", "Cos", "Tan", "Cosh", "Sinh", "Tanh", "Alog", "_10^X", "e^X", "_1/X" _
, "Sini", "Cosi", "Tani", "Sinn", "Cosn", "Tann", "Sqr", "Cur" _
, "Sinhi", "Coshi", "Tanhi", "Sec", "Csec", "Cot", "Seci", "Cseci" _
, "Coti", "Sech", "Csech", "Coth", "Sechi", "Csechi", "Cothi"

 b = array1(large2)
 C = array1(1 + large2)
 If b = "Sini" Or b = "Cosi" Then
    If C > 1 Then
     Form3.Show
     Form3.Text1.Text = Form3.Text1.Text & "Error = OverFlow" & vbCr & vbLf
     C = 0
    End If
 End If
 Select Case b
      Case "Tan", "Csec", "Sec", "Cot"
       If b = "Tan" And C = 90 Or C = 270 Then
        Form3.Show
        Form3.Text1.Text = Form3.Text1.Text & "Error = Infinte" & vbCr & vbLf
       End If
       If b = "Csec" And C = 0 Or C = 180 Then
        Form3.Show
        Form3.Text1.Text = Form3.Text1.Text & "Error = Infinte" & vbCr & vbLf
        C = 1
        End If
       If b = "Sec" And C = 90 Or C = 270 Then
        Form3.Show
        Form3.Text1.Text = Form3.Text1.Text & "Error = Infinte" & vbCr & vbLf
        End If
        If b = "Cot" And C = 0 Or C = 180 Then
        Form3.Show
        Form3.Text1.Text = Form3.Text1.Text & "Error = Infinte" & vbCr & vbLf
        C = 1
        End If
 End Select
 d = calcul(0, b, C)
 If large2 = 0 Then
 stock(j) = d
 ElseIf stock(j - 1) = "+" Or stock(j - 1) = "*" Then
  stock(j) = d
  Else
  stock(j) = "*"
  stock(j + 1) = d
  j = j + 1
  End If
large2 = large2 + 1
j = j + 1
Case "ncr", "npr", "XsrY"
a = stock(j - 1)
b = array1(large2)
C = array1(1 + large2)
d = calcul(a, b, C)
stock(j - 1) = d
large2 = large2 + 1
Case "!", "x²", "X^3", "r", "g"
a = stock(j - 1)
b = array1(large2)
d = calcul(a, b, 0)
stock(j - 1) = d
Case Else
 stock(j) = array1(large2)
 j = j + 1
 End Select
 large2 = large2 + 1
Loop
q = j
For s = 0 To q
array1(s) = stock(s)
Next s
execlog = s
End Function
Public Function ncr(n As Double, r As Double) As Double
Dim k As Double, C As Double, f, g, h
k = n - r
C = fact(n) / fact(r) * fact(k)
ncr = C
End Function
Public Function npr(n As Double, r As Double) As Double
Dim k As Double, C As Double, f, g, h
k = n - r
C = fact(n) / fact(k)
npr = C
End Function

Public Function fact(x1 As Double) As Double
Dim i As Integer, f As Integer, j As Integer
f = 1
For i = 1 To x1
f = f * i
Next i
fact = f
End Function

Public Function evaluvate() As String
Dim l As Integer, i As Integer, j As Integer, C As Double, a As Double, b As String
 Dim w As Integer, q1 As String
 q = execlog()
 large1 = 0
    stack(j) = array1(large1)
  Do While large1 <= q + 1
    Select Case array1(large1)
    Case "*", "/", "%", "^"
      a = stack(j)
      b = array1(large1)
      C = array1(1 + large1)
      d = calcul(a, b, C)
      stack(j) = d
     Case Else
     If array1(large1) = "+" Or array1(large1) = "-" Then
     j = j + 1
     stack(j) = array1(large1)
     j = j + 1
     stack(j) = array1(1 + large1)
     End If
     End Select
     large1 = large1 + 1
     Loop
     large1 = 0
     l = j
     j = 0
     Do While large1 <= l
     Select Case stack(large1)
     Case "+", "-"
      q1 = stack(j)
         If q1 = "-" Or q1 = "+" Then
         a = 0
         Else
         a = stack(j)
         End If
      b = stack(large1)
      q1 = stack(1 + large1)
          If q1 = "+" Or q1 = "-" Then
          C = 0
          Else
          C = stack(1 + large1)
          End If
      d = calcul(a, b, C)
      stack(j) = d
      End Select
     large1 = large1 + 1
     Loop
     evaluvate = stack(0)
     End Function
         
Public Function calcul(a1 As Double, b1 As String, c1 As Double) As Double
Dim z1 As Double, e As Double, pi As Double
e = 2.718281828
pi = 3.14159265358979
Select Case b1
Case "+"
z1 = a1 + c1
Case "-"
z1 = a1 - c1
Case "*"
z1 = a1 * c1
Case "/"
z1 = a1 / c1
Case "%"
z1 = a1 Mod c1
Case "^"
z1 = a1 ^ c1
Case "log"
z1 = Log(c1) * 0.43429448191
Case "In"
z1 = Log(c1)
Case "Alog"
z1 = Log(c1) * 0.43429448191
z1 = 10 ^ z1
Case "Sin"
z1 = Sin(c1 * (pi / 180))
Case "Cos"
z1 = Cos(c1 * (pi / 180))
Case "Tan"
z1 = Tan(c1 * (pi / 180))
Case "Sinh"
z1 = (e ^ c1 - (e ^ -c1)) / 2
Case "Cosh"
z1 = (e ^ c1 + (e ^ -c1)) / 2
Case "Tanh"
z1 = (e ^ c1 - (e ^ -c1)) / (e ^ c1 + (e ^ -c1))
Case "Sinn"
z1 = Sin((180 - c1 * pi / 180))
Case "Cosn"
z1 = Cos((360 - c1 * pi / 180))
Case "Tann"
z1 = Sin((180 - c1 * pi / 180)) / Cos((360 - c1 * pi / 180))
Case "Sini"
z1 = tanI(c1 / Sqr(-c1 * c1 + 1))
Case "Cosi"
z1 = tanI(-c1 / Sqr(-c1 * c1 + 1)) + 2 * tanI(1)
Case "Tani"
z1 = tanI(c1)
Case "Sinhi"
z1 = Log(c1 + Sqr(c1 * c1 + 1))
Case "Coshi"
z1 = Log(c1 + Sqr(c1 * c1 - 1))
Case "Tanhi"
z1 = Log((1 + c1) / (1 - c1)) / 2
Case "Sec"
z1 = 1 / Cos(c1 * (pi / 180))
Case "Csec"
z1 = 1 / Sin(c1 * (pi / 180))
Case "Cot"
z1 = 1 / Tan(c1 * (pi / 180))
Case "Seci"
z1 = tanI(c1 / Sqr(c1 * c1 - 1)) + Sgn((c1) - 1) * (2 * tanI(1))
Case "Cseci"
z1 = tanI(c1 / Sqr(c1 * c1 - 1)) + (Sgn(c1) - 1) * (2 * tanI(1))
Case "Coti"
z1 = tanI(c1) + 2 * tanI(1)
Case "Sech"
z1 = 2 / (Exp(c1) + Exp(-X))
Case "Csech"
z1 = 2 / (Exp(c1) - Exp(-X))
Case "Coth"
z1 = (Exp(c1) + Exp(-X)) / (Exp(c1) - Exp(-X))
Case "Sechi"
z1 = Log((Sqr(-c1 * c1 + 1) + 1) / c1)
Case "Csechi"
z1 = Log((Sgn(c1) * Sqr(c1 * c1 + 1) + 1) / c1)
Case "Cothi"
z1 = Log((c1 + 1) / (c1 - 1)) / 2
Case "x²"
z1 = a1 ^ 2
Case "X^3"
z1 = a1 ^ 3
Case "Sqr"
z1 = Sqr(c1)
Case "Cur"
z1 = c1 ^ (1 / 3)
Case "XsrY"
z1 = a1 ^ (1 / c1)
Case "_1/X"
z1 = 1 / c1
Case "_10^X"
z1 = 10 ^ c1
Case "e^X"
z1 = e ^ c1
Case "ncr"
z1 = ncr(a1, c1)
Case "npr"
z1 = npr(a1, c1)
Case "!"
z1 = fact(a1)
Case "g"
z1 = a1 * 0.9
Case "r"
z1 = a1 / pi * 180
Case Else
 MsgBox ("error1")
 End Select
 calcul = z1
 End Function
    
Public Function arrange() As Integer
 Dim q As Integer, q1 As Integer
 Dim arry(100) As String
 q = 0
 q1 = 0
 Do
 Select Case exp1(q)
 Case "0" To "9", "."
  arry(q1) = arry(q1) & exp1(q)
    Select Case exp1(q + 1)
    Case "0" To "9", "."
    Case Else
    q1 = q1 + 1
    End Select
  Case Else
  arry(q1) = exp1(q)
  q1 = q1 + 1
  End Select
  q = q + 1
  Loop Until exp1(q) = "$"
  For s1 = 0 To q1
  exp1(s1) = arry(s1)
  Next s1
  k = s1 - 1
 End Function

Public Sub push(item As String)
top1 = top1 + 1
stacks(top1) = item
End Sub
Public Function pop() As String
pop = stacks(top1)
top1 = top1 - 1
End Function
Public Function dec(b As Integer, n As Integer) As String
Dim z As Integer, i As Integer
Dim X As Integer
i = 1
X = 0
z = 0
Do While b <> 0
  X = b Mod 2
  z = z + X * i
  i = i * 10
  b = b / 2
  Loop
  dec = z
End Function
Private Function tanI(a As Double) As Double
tanI = Atn(a) * 180 / 3.14159265358979
End Function
Public Function check(b As String) As Boolean
Select Case b
Case "log", "In", "Sin", "Cos", "Tan", "Cosh", "Sinh", "Tanh", "Alog", "_10^X", "e^X", "_1/X" _
, "Sini", "Cosi", "Tani", "Sinn", "Cosn", "Tann", "Sqr", "Cur" _
, "Sinhi", "Coshi", "Tanhi", "Sec", "Csec", "Cot", "Seci", "Cseci" _
, "Coti", "Sech", "Csech", "Coth", "Sechi", "Csechi", "Cothi"
  check = True
   Case Else
    check = False
  End Select
End Function
Public Sub bin()
Dim a As Long, b, C, i As Integer, d As String
a = Form1.Text2.Text
Form1.Text2.Text = ""
C = a
Do
  a = C
  b = a Mod 2
  C = (a - b) / 2
  Form1.Text2.Text = Form1.Text2.Text & b
  Loop Until (a = 0) Or (a = 1)
  Form1.Text2.Text = rev(Form1.Text2.Text)
  End Sub
 Private Function rev(z As String) As String
Dim a As Integer
Dim b As String
a = Len(z)
Do While a >= 1
b = b & Mid$(z, a, 1)
a = a - 1
Loop
rev = b
 End Function
 
