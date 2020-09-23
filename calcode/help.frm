VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4050
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   3360
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   360
      Left            =   1560
      TabIndex        =   0
      Top             =   2640
      Width           =   1120
   End
   Begin VB.Label Label9 
      Caption         =   "Website: www.pickSourcecode.com"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label7 
      Caption         =   "Email ID: macratheesh@gmail.com"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Mathematic Functions done By"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label8 
      Caption         =   "Jorden Lab Presents:"
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   3120
   End
   Begin VB.Label Label6 
      Caption         =   "Jen Sunil"
      Height          =   225
      Left            =   1260
      TabIndex        =   5
      Top             =   1515
      Width           =   1905
   End
   Begin VB.Label Label5 
      Caption         =   "Dones, BE"
      Height          =   225
      Left            =   1260
      TabIndex        =   4
      Top             =   1320
      Width           =   1905
   End
   Begin VB.Label Label4 
      Caption         =   "Ramesh, BE"
      Height          =   225
      Left            =   1260
      TabIndex        =   3
      Top             =   1095
      Width           =   1590
   End
   Begin VB.Label Label2 
      Caption         =   "Ratheesh"
      Height          =   225
      Left            =   1260
      TabIndex        =   2
      Top             =   480
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Programed By "
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   1275
   End
End
Attribute VB_Name = "Form2"
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
Unload Me
End Sub

Private Sub Command2_Click()
Form6.Show 1
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
