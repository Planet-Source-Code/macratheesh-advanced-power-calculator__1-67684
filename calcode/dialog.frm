VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Math Functions: "
   ClientHeight    =   5085
   ClientLeft      =   2055
   ClientTop       =   2145
   ClientWidth     =   5355
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame6 
      Caption         =   "Expression"
      Height          =   1905
      Left            =   210
      TabIndex        =   28
      Top             =   420
      Width           =   4950
      Begin VB.CommandButton C 
         Caption         =   "clear"
         Height          =   330
         Left            =   3885
         TabIndex        =   31
         Top             =   1050
         Width           =   900
      End
      Begin VB.TextBox Text1 
         Height          =   1590
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   30
         Top             =   210
         Width           =   3690
      End
      Begin VB.CommandButton savee 
         Caption         =   "SaveE"
         Height          =   330
         Left            =   3885
         TabIndex        =   29
         Top             =   1470
         Width           =   900
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Formula"
      Height          =   2180
      Left            =   210
      TabIndex        =   25
      Top             =   2310
      Width           =   4950
      Begin VB.CommandButton s1 
         Caption         =   "Save"
         Height          =   330
         Left            =   3255
         TabIndex        =   35
         Top             =   1680
         Width           =   750
      End
      Begin VB.CommandButton cls 
         Caption         =   "Cls"
         Height          =   330
         Left            =   4095
         TabIndex        =   34
         Top             =   1680
         Width           =   750
      End
      Begin VB.CommandButton remv 
         Caption         =   "Remv"
         Height          =   330
         Left            =   3255
         TabIndex        =   33
         Top             =   1260
         Width           =   750
      End
      Begin VB.CommandButton add 
         Caption         =   "Add"
         Height          =   330
         Left            =   4095
         TabIndex        =   32
         Top             =   1260
         Width           =   750
      End
      Begin VB.ComboBox fc 
         Height          =   315
         Left            =   3045
         TabIndex        =   27
         Text            =   "Combo4"
         Top             =   420
         Width           =   1800
      End
      Begin VB.ListBox List3 
         Height          =   1815
         Left            =   105
         TabIndex        =   26
         Top             =   210
         Width           =   2850
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Assign Ans"
      Height          =   1935
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   360
      Width           =   4935
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1215
         Begin VB.Label Label4 
            Caption         =   "Variable:"
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.ListBox List2 
         Height          =   840
         Left            =   1920
         TabIndex        =   21
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Assign"
         Height          =   350
         Left            =   3480
         TabIndex        =   20
         Top             =   1440
         Width           =   1120
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Remove"
         Height          =   350
         Left            =   2280
         TabIndex        =   19
         Top             =   1440
         Width           =   1120
      End
      Begin VB.CommandButton clearl 
         Caption         =   "Clear"
         Height          =   350
         Left            =   1080
         TabIndex        =   18
         Top             =   1440
         Width           =   1120
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   1
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   615
         Begin VB.Label Label1 
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   120
         Width           =   855
         Begin VB.Label Label5 
            Caption         =   "Anss ="
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   0
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Help"
      Height          =   350
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4680
      Width           =   1120
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add Func"
      Height          =   2175
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   2310
      Width           =   4935
      Begin VB.CommandButton Save 
         Caption         =   "Save"
         Height          =   350
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1680
         Width           =   1120
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.ListBox List1 
         Height          =   1230
         ItemData        =   "dialog.frx":0000
         Left            =   120
         List            =   "dialog.frx":0002
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         CausesValidation=   0   'False
         Height          =   315
         ItemData        =   "dialog.frx":0004
         Left            =   3000
         List            =   "dialog.frx":0006
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Math Functions:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Button Index:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.CommandButton apply 
      Caption         =   "Apply"
      Height          =   350
      Index           =   0
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1120
   End
   Begin VB.CommandButton apply 
      Caption         =   "OK"
      Height          =   350
      Index           =   1
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1120
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4575
      Index           =   2
      Left            =   105
      TabIndex        =   3
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8070
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "MFunctions"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Formula"
            Object.ToolTipText     =   "Read only formula"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------
' Developed by Ratheesh
'       macratheesh@yahoo.com
'       macratheesh@gmail.com
' Mobile NO: +91-9894555037
' webSite: www.pickSourcecode.com
'-------------------------------------

Dim flag6 As Boolean, sav As Boolean
Private Sub add_Click()
Dim a As String
a = InputBox("Enter the Formula")
List3.AddItem (a)
sav = True
End Sub
Private Sub apply_Click(Index As Integer)
Save.Enabled = True
Select Case Combo1.Text
Case 0
Form1.abs1(0).Caption = List1.Text
Case 1
Form1.abs1(1).Caption = List1.Text
Case 2
Form1.abs1(2).Caption = List1.Text
Case 3
Form1.abs1(3).Caption = List1.Text
Case 4
Form1.abs1(4).Caption = List1.Text
Case 5
Form1.abs1(5).Caption = List1.Text
Case 6
Form1.abs1(6).Caption = List1.Text
Case 7
Form1.abs1(7).Caption = List1.Text
Case 8
Form1.abs1(8).Caption = List1.Text
Case 9
Form1.abs1(9).Caption = List1.Text
Case 10
Form1.abs1(10).Caption = List1.Text
Case Else
Form1.abs1(0).Caption = "ABS"
End Select
If apply(Index).Caption = "OK" Then
Form4.Hide
End If
apply(0).Enabled = False
apply(1).Enabled = False
If flag6 Then
Form1.Label1 = List3.Text
apply(1).Enabled = False
End If
If sav Then
    s = MsgBox("Do you want to save the changes", vbYesNo + vbExclamation)
    If s = 6 Then
    dd = CurDir & "\" & "Form.dat"
    Open dd For Output As #1
    For s = 0 To List3.ListCount
    Print #1, List3.List(s)
    Next
    Close #1
    End If
sav = False
apply(1).Enabled = False
End If
End Sub

Private Sub C_Click()
Text1.Text = ""
End Sub

Private Sub clearl_Click()
List2.clear
Command1.Enabled = False
Command3.Enabled = False
End Sub

Private Sub cls_Click()
List3.clear
sav = True
End Sub
Private Sub Combo3_Click()
List1.clear
If Combo3.ListIndex = 0 Then
List1.AddItem ("Sin")
List1.AddItem ("Cos")
List1.AddItem ("Tan")
List1.AddItem ("Sinh")
List1.AddItem ("Cosh")
List1.AddItem ("Tanh")
List1.AddItem ("Sini")
List1.AddItem ("Cosi")
List1.AddItem ("Tani")
List1.AddItem ("Sinhi")
List1.AddItem ("Coshi")
List1.AddItem ("Tanhi")
List1.AddItem ("Sinn")
List1.AddItem ("Cosn")
List1.AddItem ("Tann")
List1.AddItem ("Sec")
List1.AddItem ("Csec")
List1.AddItem ("Cot")
List1.AddItem ("Sech")
List1.AddItem ("Csech")
List1.AddItem ("Coth")
List1.AddItem ("Seci")
List1.AddItem ("Cseci")
List1.AddItem ("Coti")
List1.AddItem ("Sechi")
List1.AddItem ("Csechi")
List1.AddItem ("Cothi")
ElseIf Combo3.ListIndex = 1 Then
List1.AddItem ("x²")
List1.AddItem ("X^3")
List1.AddItem ("Sqr")
List1.AddItem ("Cur")
List1.AddItem ("XsrY")
List1.AddItem ("X^Y")
List1.AddItem ("ncr")
List1.AddItem ("npr")
List1.AddItem ("X!")
ElseIf Combo3.ListIndex = 2 Then
List1.AddItem ("log")
List1.AddItem ("Alog")
List1.AddItem ("In")
List1.AddItem ("e^X")
List1.AddItem ("e")
ElseIf Combo3.ListIndex = 3 Then
List1.AddItem ("Date")
List1.AddItem ("Time")
List1.AddItem ("CDate")
List1.AddItem ("CTime")
List1.AddItem ("Mul")
List1.AddItem ("Avg")
List1.AddItem ("Roud")
List1.AddItem ("Ceil")
List1.AddItem ("Exp")
List1.AddItem ("Sgn")
List1.AddItem ("Fix")
List1.AddItem ("Fmt")
ElseIf Combo3.ListIndex = 4 Then
listdisplay
End If
End Sub

Private Sub Command1_Click()
Dim s
If Combo2.ListIndex = -1 Then
s = MsgBox("No variable in a list Select any variable" & vbCr & vbLf & "And Select any Answer", 16)
Else
Form1.sendvar1 Combo2.Text, List2.Text
End If
End Sub
Private Sub Command2_Click()
Form4.Hide
End Sub
Private Sub Command3_Click()
Dim a As Integer, b As Integer, s
b = 0
a = List2.ListIndex
 If a = -1 Then
 s = MsgBox("No data in a list", 16)
 Else
 List2.RemoveItem (a)
 End If
End Sub

Private Sub Form_Load()
Dim i As Integer, zx, sto As String, sto1 As String
Save.Enabled = False
apply(0).Enabled = False
apply(1).Enabled = False
For i = 0 To 10
Combo1.AddItem (i)
Next i
For s = 0 To 25
Combo2.AddItem Chr(65 + s)
Next
Combo3.Text = "All Functions"
Combo3.AddItem ("Trigonometry")
Combo3.AddItem ("Algebra")
Combo3.AddItem ("Logarithms")
Combo3.AddItem ("Others")
Combo3.AddItem ("All Functions")
listdisplay
Command1.Enabled = False
Command3.Enabled = False
Frame6.Visible = False
Frame7.Visible = False

dd = CurDir & "\" & "Exp.dat"
If isfileExist(dd) Then
Open dd For Input As #1
Do While Not EOF(1)
Line Input #1, sto1
 If sto1 = "" Then
  Else
  Text1.Text = Text1 & sto1 & vbCr & vbLf
 End If
Loop
Close #1
End If

dd = CurDir & "\" & "Exp1.dat"
If isfileExist(dd) Then
Open dd For Input As #1
Do While Not EOF(1)
Line Input #1, sto1
 If sto1 = "" Then
 Else
 List2.AddItem (sto1)
 End If
Loop
Close #1
End If
flagstore = False
dd = CurDir & "\" & "Form.dat"
If isfileExist(dd) Then
Open dd For Input As #1
Do While Not EOF(1)
Input #1, save1
List3.AddItem (save1)
Loop
Close #1
End If
fc.AddItem ("UserDefFormula")
fc.Text = "UserDefFormula"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 2
Form4.Hide
End Sub

Private Sub List1_Click()
apply(0).Enabled = True
apply(1).Enabled = True
End Sub

Private Sub List2_Click()
Command1.Enabled = True
Command3.Enabled = True
apply(1).Enabled = True
End Sub

Private Sub List3_Click()
flag6 = True
apply(1).Enabled = True
End Sub

Private Sub remv_Click()
Dim C As String
C = List3.ListIndex
On Error GoTo l
List3.RemoveItem (C)
If s = 100 Then
l:
s = MsgBox("No item in the list, Ok", 16)
End If
sav = True
End Sub

Private Sub s1_Click()
dd = CurDir & "\" & "Form.dat"
Open dd For Output As #1
For s = 0 To List3.ListCount
Print #1, List3.List(s)
Next
Close #1
sav = False
End Sub

Private Sub Save_Click()
Dim save1 As String
dd = CurDir & "\" & "callist.dat"
Open dd For Output As #1
For s = 0 To 10
save1 = Form1.abs1(s).Caption
Print #1, save1
Next
Close #1
Save.Enabled = False
End Sub
Private Sub listdisplay()
List1.AddItem ("Sin")
List1.AddItem ("Cos")
List1.AddItem ("Tan")
List1.AddItem ("Sinh")
List1.AddItem ("Cosh")
List1.AddItem ("Tanh")
List1.AddItem ("Sini")
List1.AddItem ("Cosi")
List1.AddItem ("Tani")
List1.AddItem ("Sinhi")
List1.AddItem ("Coshi")
List1.AddItem ("Tanhi")
List1.AddItem ("Sinn")
List1.AddItem ("Cosn")
List1.AddItem ("Tann")
List1.AddItem ("Sec")
List1.AddItem ("Csec")
List1.AddItem ("Cot")
List1.AddItem ("Sech")
List1.AddItem ("Csech")
List1.AddItem ("Coth")
List1.AddItem ("Seci")
List1.AddItem ("Cseci")
List1.AddItem ("Coti")
List1.AddItem ("Sechi")
List1.AddItem ("Csechi")
List1.AddItem ("Cothi")
List1.AddItem ("x²")
List1.AddItem ("X^3")
List1.AddItem ("Sqr")
List1.AddItem ("Cur")
List1.AddItem ("XsrY")
List1.AddItem ("X^Y")
List1.AddItem ("ncr")
List1.AddItem ("npr")
List1.AddItem ("X!")
List1.AddItem ("log")
List1.AddItem ("Alog")
List1.AddItem ("In")
List1.AddItem ("e^X")
List1.AddItem ("e")
List1.AddItem ("Date")
List1.AddItem ("Time")
List1.AddItem ("CDate")
List1.AddItem ("CTime")
List1.AddItem ("Mul")
List1.AddItem ("Avg")
List1.AddItem ("Roud")
List1.AddItem ("Ceil")
List1.AddItem ("Exp")
List1.AddItem ("Sgn")
List1.AddItem ("Fix")
List1.AddItem ("Fmt")
List1.AddItem ("Pur")
End Sub
Private Sub savee_Click()
dd = CurDir & "\" & "Exp.dat"
Open dd For Output As #1
Print #1, Text1.Text
Close #1
dd = CurDir & "\" & "Exp1.dat"
Open dd For Output As #1
For s = 0 To List2.ListCount
Print #1, List2.List(s)
Next
flagstore = True
Close #1
End Sub

Private Sub TabStrip1_Click(Index As Integer)

        If 1 = TabStrip1(Index).SelectedItem.Index - 1 Then
           Frame3(0).Visible = False
           Frame1(0).Visible = False
           Frame6.Visible = True
           Frame7.Visible = True
           Else
           Frame6.Visible = False
           Frame7.Visible = False
           Frame3(0).Visible = True
           Frame1(0).Visible = True
         End If
  
End Sub
