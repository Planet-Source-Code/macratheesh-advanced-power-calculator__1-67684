VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form6 
   Caption         =   "Help "
   ClientHeight    =   6750
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   8730
   Icon            =   "helpcontent.frx":0000
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   8730
   Begin VB.PictureBox Picture1 
      Height          =   6735
      Left            =   0
      ScaleHeight     =   6675
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton find 
         Caption         =   "Display"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1440
         TabIndex        =   4
         Top             =   6240
         Width           =   1120
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1680
         TabIndex        =   3
         Top             =   1440
         Width           =   1120
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   360
         Width           =   2415
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4140
         Left            =   120
         TabIndex        =   1
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Type in the keyword to find :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Help Topics"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Searched List:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2535
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6765
      Left            =   3000
      TabIndex        =   9
      Top             =   0
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   11933
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      FileName        =   "E:\Rose\DataCmp\welcome.rtf"
      TextRTF         =   $"helpcontent.frx":030A
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   3405
      Top             =   3135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "helpcontent.frx":105E4
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "helpcontent.frx":106F6
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "helpcontent.frx":10808
            Key             =   "back"
         EndProperty
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form6"
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

Private Sub Combo1_Click()
List1.clear
If Combo1.Text = "Contents" Then
List1.AddItem "Cut,Copy,Paste or Delete text"
List1.AddItem "View types"
List1.AddItem "Window arrangement"
List1.AddItem "Zip files"
List1.AddItem "Unzip files"
Label1.Visible = False
Text1.Visible = False
find.Visible = False
Command2.Visible = False

End If
If Combo1.Text = "Index" Then
List1.AddItem "compress text"
List1.AddItem "copying text"
List1.AddItem "cutting text"
List1.AddItem "delete text"
List1.AddItem "decompress text"
List1.AddItem "finding text"
List1.AddItem "inserting text"
List1.AddItem "moving text"
List1.AddItem "pasting text"
List1.AddItem "window size"
Label1.Visible = True
Text1.Visible = True
find.Visible = True
Command2.Visible = False
End If
If Combo1.Text = "Search" Then
End If

End Sub

Private Sub Command2_Click()
Text1.Text = ""
End Sub

Private Sub exit_Click()
Unload Me
End Sub
Private Sub find_Click()
List1_DblClick
End Sub

Private Sub Form_Load()
Combo1.Text = "All Topics"
Combo1.AddItem "Contents"
Combo1.AddItem "Index"
Combo1.AddItem "Search"

End Sub

Private Sub Form_Resize()
    Picture1.Visible = True
    RichTextBox1.Move 2900, 0, Me.ScaleWidth - 2880, Me.ScaleHeight
    RichTextBox1.RightMargin = RichTextBox1.Width
    Picture1.Move 0, 0, 2900, Me.ScaleHeight
End Sub

Private Sub List1_DblClick()
Dim s
If Combo1.Text = "Contents" Then
Select Case List1.ListIndex
Case 0:
dd = CurDir & "\" & "edit1.rtf"
If isfileExist(dd) Then
RichTextBox1.LoadFile (dd)
s = RichTextBox1.find(List1.Text, , , rtwholeword)
Else
s = MsgBox("Help File Found", 16)
End If
Case 1:
dd = CurDir & "\" & "view.rtf"
If isfileExist(dd) Then
RichTextBox1.LoadFile (dd)
s = RichTextBox1.find(List1.Text, , , rtwholeword)
Else
s = MsgBox("Help File Found", 16)
End If
Case 2:
dd = CurDir & "\" & "window.rtf"
If isfileExist(dd) Then
RichTextBox1.LoadFile (dd)
s = RichTextBox1.find(List1.Text, , , rtwholeword)
Else
s = MsgBox("Help File Found", 16)
End If
Case 3:
dd = CurDir & "\" & "compress.rtf"
If isfileExist(dd) Then
RichTextBox1.LoadFile (dd)
s = RichTextBox1.find(List1.Text, , , rtwholeword)
Else
s = MsgBox("Help File Found", 16)
End If
Case 4:
dd = CurDir & "\" & "decompress.rtf"
If isfileExist(dd) Then
RichTextBox1.LoadFile (dd)
s = RichTextBox1.find(List1.Text, , , rtwholeword)
Else
s = MsgBox("Help File Found", 16)
End If
End Select
End If
If Combo1.Text = "Index" Then
Text1.Text = List1.Text
Select Case List1.ListIndex
Case 0:
dd = CurDir & "\" & "compress.rtf"
If isfileExist(dd) Then
RichTextBox1.LoadFile (dd)
s = RichTextBox1.find(List1.Text, , , rtwholeword)
Else
s = MsgBox("Help File Found", 16)
End If
Case 1:
dd = CurDir & "\" & "copy.rtf"
If isfileExist(dd) Then
RichTextBox1.LoadFile (dd)
s = RichTextBox1.find(List1.Text, , , rtwholeword)
Else
s = MsgBox("Help File Found", 16)
End If
Case 2:
dd = CurDir & "\" & "cut.rtf"
If isfileExist(dd) Then
RichTextBox1.LoadFile (dd)
s = RichTextBox1.find(List1.Text, , , rtwholeword)
Else
s = MsgBox("Help File Found", 16)
End If
Case 3:
dd = CurDir & "\" & "delete.rtf"
If isfileExist(dd) Then
RichTextBox1.LoadFile (dd)
s = RichTextBox1.find(List1.Text, , , rtwholeword)
Else
s = MsgBox("Help File Found", 16)
End If
Case 4:
dd = CurDir & "\" & "decompress.rtf"
If isfileExist(dd) Then
RichTextBox1.LoadFile (dd)
s = RichTextBox1.find(List1.Text, , , rtwholeword)
Else
s = MsgBox("Help File Found", 16)
End If
Case 5:
dd = CurDir & "\" & "find.rtf"
If isfileExist(dd) Then
RichTextBox1.LoadFile (dd)
s = RichTextBox1.find(List1.Text, , , rtwholeword)
Else
s = MsgBox("Help File Found", 16)
End If
Case 6:
dd = CurDir & "\" & "insert.rtf"
If isfileExist(dd) Then
RichTextBox1.LoadFile (dd)
s = RichTextBox1.find(List1.Text, , , rtwholeword)
Else
s = MsgBox("Help File Found", 16)
End If
Case 7:
dd = CurDir & "\" & "move.rtf"
If isfileExist(dd) Then
RichTextBox1.LoadFile (dd)
s = RichTextBox1.find(List1.Text, , , rtwholeword)
Else
s = MsgBox("Help File Found", 16)
End If
Case 8:
dd = CurDir & "\" & "paste.rtf"
If isfileExist(dd) Then
RichTextBox1.LoadFile (dd)
s = RichTextBox1.find(List1.Text, , , rtwholeword)
Else
s = MsgBox("Help File Found", 16)
End If
Case 9:
dd = CurDir & "\" & "window1.rtf"
If isfileExist(dd) Then
RichTextBox1.LoadFile (dd)
s = RichTextBox1.find(List1.Text, , , rtwholeword)
Else
s = MsgBox("Help File Found", 16)
End If
End Select
End If
End Sub

