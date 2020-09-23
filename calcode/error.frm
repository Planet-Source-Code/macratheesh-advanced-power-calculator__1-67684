VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compile Error....."
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4680
   DrawMode        =   3  'Not Merge Pen
   DrawStyle       =   3  'Dash-Dot
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Bookman Old Style"
      Size            =   11.25
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000006&
   HasDC           =   0   'False
   HelpContextID   =   1
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command2 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3240
      TabIndex        =   2
      Top             =   3120
      Width           =   1120
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1120
   End
End
Attribute VB_Name = "Form3"
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

Private Sub Command1_Click()
Form3.Hide
End Sub

Private Sub Command2_Click()
Form6.Show
End Sub
