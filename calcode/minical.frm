VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form7"
   ClientHeight    =   1140
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   2925
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton clear 
      Caption         =   "C"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton equal 
      Caption         =   "="
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu sc 
         Caption         =   "&Super Calculater"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clear_Click()
Text1.Text = " "
End Sub

Private Sub equal_Click()
call1
 k = k + 1
 exp1(k) = ")"
 exp1(k + 1) = "$"
 arrange
 ans = execute()
 Form4.List2.AddItem (ans)
 Text1.Text = ans
 ans1 = ans
 k = sign
 qw = Text1.Text
 top1 = -1
 k = 0
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub sc_Click()
Form1.Cls
End Sub
Public Sub call1()
Dim c As String * 1
Open "trt.dat" For Output As #1
Print #1, Text1.Text
Close #1

Open "trt.dat" For Input As #1
For s = 1 To Len(Text1.Text) + 2
k = k + 1
Seek #1, s
Input #1, c
exp1(k) = c
express(k) = c
sign = k
 If check(c) Then
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
    exp1(k) = c
    express(k) = c
    Case Else
    k = k + 1
   End Select
End If
Next s
Close #1
End Sub



