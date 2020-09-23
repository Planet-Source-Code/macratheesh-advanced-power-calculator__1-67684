VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "RefList"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form5"
   ScaleHeight     =   4095
   ScaleWidth      =   4950
   Begin VB.ListBox List3 
      Height          =   3765
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   3765
      ItemData        =   "relist.frx":0000
      Left            =   1680
      List            =   "relist.frx":0002
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.ListBox list1 
      Height          =   3765
      ItemData        =   "relist.frx":0004
      Left            =   0
      List            =   "relist.frx":0006
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Purpose:"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Ans:"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Expressen:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Form_Load()

End Sub
Private Sub Form_Resize()
list1.Move 0, 200, Me.ScaleWidth / 3, Form5.Height - 525
List2.Move list1.Width + 20, 200, Form5.ScaleWidth / 3, Form5.Height - 525
Form5.List3.Move list1.Width + List2.Width + 50, 200, Form5.Width, Form5.Height - 525
Label1.Left = list1.Left
Label2.Left = List2.Left
Label3.Left = List3.Left
End Sub

Private Sub List2_DblClick()
Form4.List2.AddItem (Form5.List2.List(List2.ListIndex))
End Sub
