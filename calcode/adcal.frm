VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced Power Calculator"
   ClientHeight    =   4995
   ClientLeft      =   690
   ClientTop       =   2190
   ClientWidth     =   7365
   ForeColor       =   &H8000000C&
   HasDC           =   0   'False
   Icon            =   "adcal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   4995
   ScaleWidth      =   7365
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame set 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   -50
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1455
         Left            =   0
         TabIndex        =   90
         Top             =   0
         Width           =   3135
         Begin VB.CommandButton hep 
            Caption         =   "Help"
            Height          =   415
            Left            =   2280
            TabIndex        =   94
            Top             =   960
            Width           =   520
         End
         Begin VB.TextBox Text3 
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   1  'Horizontal
            TabIndex        =   93
            ToolTipText     =   "Enter the Expression through Keyboard"
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton equal1 
            Caption         =   "="
            Default         =   -1  'True
            Height          =   415
            Left            =   1080
            TabIndex        =   92
            Top             =   960
            Width           =   520
         End
         Begin VB.CommandButton cls1 
            Caption         =   "C"
            Height          =   415
            Left            =   1680
            TabIndex        =   91
            Top             =   960
            Width           =   520
         End
      End
      Begin VB.Frame Frame9 
         Height          =   495
         Left            =   3720
         TabIndex        =   80
         Top             =   840
         Width           =   3615
         Begin VB.OptionButton hyp 
            Caption         =   "Hyp Inves"
            Height          =   255
            Index           =   3
            Left            =   1680
            TabIndex        =   84
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton hyp 
            Caption         =   "Natul"
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   83
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton hyp 
            Caption         =   "Inves"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   82
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton hyp 
            Caption         =   "Hyp"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   81
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame Frame8 
         Height          =   480
         Left            =   3720
         TabIndex        =   75
         Top             =   840
         Width           =   3615
         Begin VB.OptionButton Option1 
            Caption         =   "Hexdec"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   3
            Left            =   2640
            TabIndex        =   79
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Octal"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   78
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Binary"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   77
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Dec"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   76
            Top             =   120
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   72
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   360
         Width           =   3615
      End
      Begin VB.Frame Frame1 
         Height          =   2265
         Left            =   120
         TabIndex        =   54
         Top             =   1320
         Width           =   2625
         Begin VB.CommandButton but 
            Caption         =   "8"
            Height          =   415
            Index           =   1
            Left            =   720
            TabIndex        =   70
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "9"
            Height          =   415
            Index           =   2
            Left            =   1320
            TabIndex        =   69
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton clear 
            Caption         =   "C"
            Height          =   415
            Left            =   1920
            TabIndex        =   68
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "4"
            Height          =   415
            Index           =   3
            Left            =   120
            TabIndex        =   67
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "5"
            Height          =   415
            Index           =   4
            Left            =   720
            TabIndex        =   66
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "6"
            Height          =   415
            Index           =   5
            Left            =   1320
            TabIndex        =   65
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton Command8 
            Caption         =   "M+"
            Height          =   415
            Left            =   1920
            TabIndex        =   64
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "1"
            Height          =   415
            Index           =   6
            Left            =   120
            TabIndex        =   63
            Top             =   1200
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "2"
            Height          =   415
            Index           =   7
            Left            =   720
            TabIndex        =   62
            Top             =   1200
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "3"
            Height          =   415
            Index           =   8
            Left            =   1320
            TabIndex        =   61
            Top             =   1200
            Width           =   520
         End
         Begin VB.CommandButton Command12 
            Caption         =   "M-"
            Height          =   415
            Left            =   1920
            TabIndex        =   60
            Top             =   1200
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "0"
            Height          =   415
            Index           =   9
            Left            =   120
            TabIndex        =   59
            Top             =   1680
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   10
            Left            =   720
            TabIndex        =   58
            Top             =   1680
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "Ans"
            Height          =   415
            Index           =   11
            Left            =   1320
            TabIndex        =   57
            Top             =   1680
            Width           =   520
         End
         Begin VB.CommandButton equal 
            Caption         =   "="
            Height          =   415
            Left            =   1920
            TabIndex        =   56
            Top             =   1680
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "7"
            Height          =   415
            Index           =   0
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   520
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2265
         Left            =   2760
         TabIndex        =   49
         Top             =   1320
         Width           =   800
         Begin VB.CommandButton shift 
            Caption         =   "Shift"
            Height          =   415
            Left            =   120
            TabIndex        =   53
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton Command38 
            Caption         =   "MEM"
            Height          =   415
            Left            =   120
            TabIndex        =   52
            Top             =   1200
            Width           =   520
         End
         Begin VB.CommandButton Ac 
            Caption         =   "AC"
            Height          =   415
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton del 
            Caption         =   "DEL"
            Height          =   415
            Left            =   120
            TabIndex        =   50
            Top             =   1680
            Width           =   520
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Stan Func"
         Height          =   2265
         Left            =   3600
         TabIndex        =   24
         Top             =   1320
         Width           =   3750
         Begin VB.CommandButton but 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   12
            Left            =   120
            TabIndex        =   48
            ToolTipText     =   "Multiply"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   13
            Left            =   720
            TabIndex        =   47
            ToolTipText     =   "Divide"
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "log"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   20
            Left            =   1320
            TabIndex        =   46
            ToolTipText     =   "Log"
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "In"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   21
            Left            =   1920
            TabIndex        =   45
            ToolTipText     =   "In"
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "Alog"
            Height          =   415
            Index           =   22
            Left            =   2520
            TabIndex        =   44
            ToolTipText     =   "Anti log"
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   14
            Left            =   120
            TabIndex        =   43
            ToolTipText     =   "Addtion"
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "Sin"
            Height          =   415
            Index           =   24
            Left            =   1320
            TabIndex        =   42
            ToolTipText     =   "Sin & Sinh"
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "Cos"
            Height          =   415
            Index           =   25
            Left            =   1920
            TabIndex        =   41
            ToolTipText     =   "Cos & Cosh "
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "Tan"
            Height          =   415
            Index           =   26
            Left            =   2520
            TabIndex        =   40
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   16
            Left            =   120
            TabIndex        =   39
            ToolTipText     =   "Mod"
            Top             =   1200
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "^"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   17
            Left            =   720
            TabIndex        =   38
            ToolTipText     =   "Power"
            Top             =   1200
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "x²"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   28
            Left            =   1320
            TabIndex        =   37
            Top             =   1200
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "X^3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   29
            Left            =   1920
            TabIndex        =   36
            Top             =   1200
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "Sqr"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   30
            Left            =   2520
            TabIndex        =   35
            Top             =   1200
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "("
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   18
            Left            =   120
            TabIndex        =   34
            ToolTipText     =   "Open parathesiss"
            Top             =   1680
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   ")"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   19
            Left            =   720
            TabIndex        =   33
            ToolTipText     =   "Close parathesis"
            Top             =   1680
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "Cur"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   31
            Left            =   3120
            TabIndex        =   32
            Top             =   1200
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "XsrY"
            Height          =   415
            Index           =   32
            Left            =   3120
            TabIndex        =   31
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "_1/X"
            Height          =   415
            Index           =   33
            Left            =   2520
            TabIndex        =   30
            Top             =   1680
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "_10^X"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   23
            Left            =   3120
            TabIndex        =   29
            ToolTipText     =   "10 power X  & G"
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "ncr"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   27
            Left            =   1920
            TabIndex        =   28
            Top             =   1680
            Width           =   520
         End
         Begin VB.CommandButton fact1 
            Caption         =   "X!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Left            =   1320
            TabIndex        =   27
            Top             =   1680
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "e^X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   34
            Left            =   3120
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   520
         End
         Begin VB.CommandButton but 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   12
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   15
            Left            =   720
            MaskColor       =   &H0080FF80&
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Subration"
            Top             =   720
            Width           =   520
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Variables"
         Height          =   1335
         Left            =   120
         TabIndex        =   13
         Top             =   3600
         Width           =   3255
         Begin VB.Frame Frame5 
            Caption         =   "Variables"
            Height          =   1335
            Left            =   0
            TabIndex        =   85
            Top             =   0
            Width           =   3255
            Begin VB.CommandButton back 
               Caption         =   "BCK"
               Height          =   415
               Left            =   120
               TabIndex        =   89
               Top             =   720
               Width           =   520
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   1440
               TabIndex        =   88
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton rcl 
               Caption         =   "RCL"
               Height          =   415
               Index           =   1
               Left            =   840
               TabIndex        =   87
               Top             =   240
               Width           =   520
            End
            Begin VB.CommandButton sto 
               Caption         =   "STO"
               Height          =   415
               Index           =   1
               Left            =   120
               TabIndex        =   86
               Top             =   240
               Width           =   520
            End
         End
         Begin VB.CommandButton sto 
            Caption         =   "STO"
            Height          =   415
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton rcl 
            Caption         =   "RCL"
            Height          =   415
            Index           =   0
            Left            =   720
            TabIndex        =   22
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton vari 
            Caption         =   "A"
            Height          =   415
            Index           =   0
            Left            =   1320
            TabIndex        =   21
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton vari 
            Caption         =   "B"
            Height          =   415
            Index           =   1
            Left            =   1920
            TabIndex        =   20
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton vari 
            Caption         =   "C"
            Height          =   415
            Index           =   2
            Left            =   2520
            TabIndex        =   19
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton vari 
            Caption         =   "D"
            Height          =   415
            Index           =   3
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton vari 
            Caption         =   "E"
            Height          =   415
            Index           =   4
            Left            =   720
            TabIndex        =   17
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton vari 
            Caption         =   "F"
            Height          =   415
            Index           =   5
            Left            =   1320
            TabIndex        =   16
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton vari 
            Caption         =   "G"
            Height          =   415
            Index           =   6
            Left            =   1920
            TabIndex        =   15
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton more 
            Caption         =   "AV"
            Height          =   415
            Left            =   2520
            TabIndex        =   14
            Top             =   720
            Width           =   520
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Ex Func"
         Height          =   1335
         Left            =   3600
         TabIndex        =   1
         Top             =   3600
         Width           =   3750
         Begin VB.CommandButton abs1 
            Caption         =   "Fix"
            Height          =   415
            Index           =   10
            Left            =   2520
            TabIndex        =   95
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton AF 
            Caption         =   "AF"
            Height          =   415
            Left            =   3120
            TabIndex        =   12
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton abs1 
            Caption         =   "Avg"
            Height          =   415
            Index           =   8
            Left            =   1320
            TabIndex        =   11
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton abs1 
            Caption         =   "Mul"
            Height          =   415
            Index           =   7
            Left            =   720
            TabIndex        =   10
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton abs1 
            Caption         =   "Sgn"
            Height          =   415
            Index           =   9
            Left            =   1920
            TabIndex        =   9
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton abs1 
            Caption         =   "TIME"
            Height          =   415
            Index           =   6
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   520
         End
         Begin VB.CommandButton abs1 
            Caption         =   "DAT"
            Height          =   415
            Index           =   5
            Left            =   3120
            TabIndex        =   7
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton abs1 
            Caption         =   "Fmt"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   415
            Index           =   4
            Left            =   2520
            TabIndex        =   6
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton abs1 
            Caption         =   "RND"
            Height          =   415
            Index           =   3
            Left            =   1920
            TabIndex        =   5
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton abs1 
            Caption         =   "FLR"
            Height          =   415
            Index           =   2
            Left            =   1320
            TabIndex        =   4
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton abs1 
            Caption         =   "CEIL"
            Height          =   415
            Index           =   1
            Left            =   720
            TabIndex        =   3
            Top             =   240
            Width           =   520
         End
         Begin VB.CommandButton abs1 
            Caption         =   "ABS"
            Height          =   415
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   520
         End
      End
      Begin MSComctlLib.ImageList imlToolbarIcons 
         Left            =   3255
         Top             =   2835
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "adcal.frx":030A
               Key             =   "Back"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "adcal.frx":041C
               Key             =   "Forward"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "adcal.frx":052E
               Key             =   "Copy"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "adcal.frx":0640
               Key             =   "Cut"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "adcal.frx":0752
               Key             =   "Paste"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "adcal.frx":0864
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "adcal.frx":0976
               Key             =   "Help"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000008&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   0
         Width           =   3495
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000008&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3720
         TabIndex        =   73
         Top             =   0
         Width           =   3615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         BorderWidth     =   2
         X1              =   120
         X2              =   120
         Y1              =   480
         Y2              =   840
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         X1              =   120
         X2              =   120
         Y1              =   480
         Y2              =   840
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      WindowList      =   -1  'True
      Begin VB.Menu mini 
         Caption         =   "Standard"
      End
      Begin VB.Menu scientific 
         Caption         =   "Scientific"
      End
      Begin VB.Menu ref 
         Caption         =   "refList"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu HelpT 
         Caption         =   "&Help Topics"
      End
      Begin VB.Menu h 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
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

Public flag5 As Boolean, flag10 As Boolean
Public qw As Double, var As Boolean, var1 As Boolean
Public zx As Boolean, ans2 As String
Dim storevar(26, 1) As String
Private Sub about_Click()
Form2.Timer1.Enabled = False
Form2.Show 1
End Sub
Private Sub abs1_Click(Index As Integer)
Select Case abs1(Index).Caption
Case "ABS"
Text2.Text = Abs(Text2.Text)
Case "CEIL"
Case "FLR"
Case "RND"
Text2.Text = Round(Text2.Text, 2)
Case "Fmt"
Text2.Text = Format(Text2.Text, "#,##,###,###.#0")
Case "DAT"
Label2.Caption = "Date is  " & Date
x1 = 0
Case "TIME"
Label2.Caption = "Time is  " & Time
x1 = 1
Case "Exp"
On Error GoTo label
Text2.Text = Exp(Text2.Text)
label:
Case "Fix"
On Error GoTo labe
Text2.Text = Fix(Text2.Text)
labe:
Case "Ceil"
On Error GoTo lab
Text2.Text = Int(Text2.Text)
lab:
Case "MORE"
Form4.Show 1
Case "Pur"
rs1.MoveLast
rs1.Edit
rs1!purp = InputBox("Enter the Purpose:", "Purpose")
rs1.Update
Case Else
display abs1(Index).Caption
End Select
End Sub
Private Sub Ac_Click()
Dim flag3
flag3 = MsgBox("Or You Sure You Want to Clear Memory", vbYesNoCancel, "Advanced Power Calculator")
If flag3 = 6 Then
Text1.Text = ""
Text2.Text = "0.0"
Label1.Caption = ""
Label2.Caption = ""
Form4.Text1.Text = ""
exp1(0) = "("
b1 = 2
k = 0
zx = True
top1 = -1
b = True
Form3.Cls
For s = 0 To 25
storevar(s, 0) = 0
Next
End If
End Sub
Private Sub AF_Click()
Form4.Show 1
End Sub

Private Sub back_Click()
Frame5.Visible = False
End Sub

Private Sub cls1_Click()
Text3.Text = ""
End Sub

Private Sub Combo1_Click()
Dim i As Integer
If var = True Then
 Label2.Caption = storevar(Combo1.ListIndex, 1) & " = " & Text2.Text
 storevar(Combo1.ListIndex, 0) = Text2.Text
 var = False
ElseIf var1 = True Then
    Text2.Text = storevar(Combo1.ListIndex, 0)
    Label2.Caption = storevar(Combo1.ListIndex, 1) & " = " & Text2.Text
    var1 = False
    zx = False
Else
display Combo1.Text
End If
 dd = CurDir & "\" & "Var.dat"
 Open dd For Output As #1
 For s = 0 To 26
 Print #1, storevar(s, 0)
 Next s
 Close #1
End Sub
Private Sub equal1_Click()
 call1
 k = k + 1
 exp1(k) = ")"
 exp1(k + 1) = "$"
 If flag5 Then
 arrange
 On Error GoTo skip
 Text3.Text = store
 ans = execute()
 Text3.Alignment = 1
 Text3.Text = ans
 cls1.SetFocus
 Form4.List2.AddItem (ans)
 flag5 = True
 flag10 = True
 Else
skip:
 MsgBox "Character Error or Invalid Operater error", , "Quick Calculator_Error"
 End If
 top1 = -1
 k = 0
End Sub
Private Sub exit_Click()
Dim s
If flagstore Then
 End
 Unload Form1
 Unload Form2
 Unload Form3
 Unload Form4
 Unload Form6
Else
s = MsgBox("Do you want to save the changes", vbYesNoCancel + 48, "Advanced Power Calculator")
 If s = 6 Then
    save2
    End
    Unload Form1
    Unload Form2
    Unload Form3
    Unload Form4
    Unload Form6
  ElseIf s = 7 Then
  End
    Unload Form1
    Unload Form2
    Unload Form3
    Unload Form4
    Unload Form6
  ElseIf s = 2 Then
  Cancel = 2
 End If
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim s
If flagstore Then
 End
 Unload Form1
 Unload Form2
 Unload Form3
 Unload Form4
 Unload Form6
Else
s = MsgBox("Do you want to save the changes", vbYesNoCancel + 48, "Advanced Power Calculator")
 If s = 6 Then
    save2
    End
    Unload Form1
    Unload Form2
    Unload Form3
    Unload Form4
    Unload Form6
  ElseIf s = 7 Then
  End
    Unload Form1
    Unload Form2
    Unload Form3
    Unload Form4
    Unload Form6
  ElseIf s = 2 Then
  Cancel = 2
 End If
End If
End Sub

Private Sub HelpT_Click()
Form6.Visible = True
End Sub
Private Sub hep_Click()
MsgBox "Help Not Found", 16, "Quick Calculator_Error"
End Sub

Private Sub hyp_Click(Index As Integer)
If hyp(0).Value = True Then
but(24).Caption = "Sinh"
but(25).Caption = "Cosh"
but(26).Caption = "Tanh"
ElseIf hyp(1).Value = True Then
but(24).Caption = "Sini"
but(25).Caption = "Cosi"
but(26).Caption = "Tani"
ElseIf hyp(2).Value = True Then
but(24).Caption = "Sinn"
but(25).Caption = "Cosn"
but(26).Caption = "Tann"
ElseIf hyp(3).Value = True Then
but(24).Caption = "Sinhi"
but(25).Caption = "Coshi"
but(26).Caption = "Tanhi"
End If
End Sub

Private Sub mini_Click()
Text1.Visible = False
Text2.Visible = False
Label1.Visible = False
Label2.Visible = False
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame8.Visible = False
Frame9.Visible = False
Frame7.Visible = True
Form1.Height = 2050
Form1.Width = 3130
End Sub

Private Sub more_Click()
Frame5.Visible = True
End Sub

Private Sub rcl_Click(Index As Integer)
clear2
var1 = True
var = False
End Sub

Private Sub ref_Click()
MsgBox "Not Found", 16, "Quick Calculator_Error"
End Sub

Private Sub scientific_Click()
Text1.Visible = True
Text2.Visible = True
Label1.Visible = True
Label2.Visible = True
Frame1.Visible = True
Frame2.Visible = True
Frame3.Visible = True
Frame4.Visible = True
Frame6.Visible = True
Frame8.Visible = True
Frame7.Visible = False
Form1.Height = 5655
Form1.Width = 7455
End Sub

Private Sub shift_Click()
If a = True Then
hyp(0).Value = True
Frame9.Visible = True
but(22).Caption = "EXP"
but(24).Caption = "Sinh"
but(25).Caption = "Cosh"
but(26).Caption = "Tanh"
but(27).Caption = "npr"
but(17).Caption = "+/-"
but(23).Caption = "g"
but(34).Caption = "e"
but(33).Caption = "pi"
but(32).Caption = "r"
a = False
Else
Frame9.Visible = False
but(24).Caption = "Sin"
but(22).Caption = "Alog"
but(25).Caption = "Cos"
but(26).Caption = "Tan"
but(27).Caption = "ncr"
but(17).Caption = "^"
but(23).Caption = "_10^X"
but(34).Caption = "e^X"
but(33).Caption = "_1/X"
but(32).Caption = "XsrY"
Option1(0).Value = True
a = True
End If
End Sub

Private Sub sto_Click(Index As Integer)
var = True
var1 = False
End Sub

Private Sub Timer1_Timer()
If x1 = 1 Then
Label2.Caption = "Time is" & Time
End If
End Sub
Private Sub Form_Load()
Dim save1 As String, l
Form2.Show 1
flagstore = True
Frame7.Visible = False
Combo1.Clear
For s = 0 To 25
Combo1.AddItem Chr(65 + s)
storevar(s, 1) = Chr(65 + s)
Next
Text2.Text = "0.0"
b = True
a = True
num(0) = -1
k = 0
exp1(0) = "("
p1 = 0
top1 = -1
ans1 = 0
Frame5.Visible = False
Frame9.Visible = False
Option1(0).Value = True
dd = CurDir & "\" & "callist.dat"
If isfileExist(dd) Then
Open dd For Input As #1
For s = 0 To 10
Input #1, save1
abs1(s).Caption = save1
Next
Close #1
Else
Open dd For Output As #1
For s = 0 To 10
save1 = Form1.abs1(s).Caption
Print #1, save1
Next
Close #1
End If
flag5 = True

dd = CurDir & "\" & "Var.dat"
If isfileExist(dd) Then
Open dd For Input As #1
For s = 0 To 26
Input #1, save1
storevar(s, 0) = save1
Next s
Close #1
Else
Open dd For Output As #1
For s = 0 To 26
Print #1, 0
Next s
Close #1
End If
End Sub
Private Sub but_Click(Index As Integer)
clear1
b = True
zx = False
display but(Index).Caption
End Sub
Private Sub equal_Click()
Dim X As Boolean
 If k = 0 Then
  k = k + 1
  exp1(1) = 0#
  express(1) = 0#
 End If
Option1(0).Value = True
X = compile()
If X = True Then
Form3.Hide
If b = True Then
 k = k + 1
 exp1(k) = ")"
 exp1(k + 1) = "$"
arrange
If Len(Text1.Text) = 0 Then
Else
ans = execute()
Label1.Caption = "Success."
Form4.List2.AddItem (ans)
Form4.Text1.Text = Form4.Text1 & Text1.Text & " = " & ans & vbCr & vbLf
End If
Label2.Caption = "Ans = " & ans
Text2.Text = ans
st = ans
ans1 = ans
k = sign
b = False
End If
b1 = 1
Else
Label1.Caption = "Compile Error.."
Label2.Caption = "Ans = 0.0"
End If
qw = Text2.Text
top1 = -1
End Sub
Private Sub clear_Click()
Text1.Text = ""
Text2.Text = "0.0"
exp1(0) = "("
b1 = 2
k = 0
zx = True
top1 = -1
b = True
Form3.Cls
End Sub
Private Sub del_Click()
Form3.Cls
store = " "
If sign > 0 Then
sign = sign - 1
k = sign
For s = 1 To sign
exp1(s) = express(s)
If flag1 = True Then
store = store & " " & express(s)
flag1 = False
Else
store = store & express(s)
End If
 If check(express(s)) Then
   flag1 = True
 End If
Next s
Text1.Text = store
store = " "
top1 = -1
b1 = 0
b = True
zx = False
Else
b = False
End If
End Sub
Private Sub fact1_Click()
display "!"
End Sub
Private Sub Option1_Click(Index As Integer)
Dim dec1 As Integer, f
If Option1(1).Value = True Then
bin
ElseIf Option1(0).Value = True Then
Text2.Text = st
ElseIf Option1(2).Value = True Then
Text2.Text = Oct(dec1)
ElseIf Option1(3).Value = True Then
Text2.Text = Hex(dec1)
f = dec1
End If
End Sub

Private Sub Text3_Click()
Text3.Alignment = 0
flag10 = False
End Sub

Private Sub Text3_GotFocus()
If flag10 = True Then
Text3.Alignment = 1
flag10 = False
Else
Text3.Alignment = 0
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.Alignment = 0
End If
End Sub
Private Sub vari_Click(Index As Integer)
Dim i As Integer
If var = True Then
 Label2.Caption = storevar(Index, 1) & " = " & Text2.Text
 storevar(Index, 0) = Text2.Text
 var = False
ElseIf var1 = True Then
   Text2.Text = storevar(Index, 0)
   Label2.Caption = storevar(Index, 1) & " = " & storevar((Index), 0)
   var1 = False
Else
display vari(Index).Caption
End If
 dd = CurDir & "\" & "Var.dat"
 Open dd For Output As #1
 For s = 0 To 26
 Print #1, storevar(s, 0)
 Next s
 Close #1
End Sub
Public Function sendvar(a5 As String) As String
 Dim i As Integer
 i = 0
 Do
   sendvar = storevar(i, 0)
   i = i + 1
  Loop Until storevar(i, 1) <> a5
 End Function

Public Sub clear2()
Text1.Text = ""
Text2.Text = "0.0"
exp1(0) = "("
b1 = 2
k = 0
zx = True
top1 = -1
b = True
Form3.Cls
End Sub
Private Sub trans()
For s = 0 To Len(Form1.Text1.Text) + 2
exp1(s) = express(s)
Next s
k = s
End Sub
Public Sub sendvar1(a5 As String, b As Double)
 Dim i As Integer
 i = -1
 Do
   i = i + 1
   If storevar(i, 1) = a5 Then
   storevar(i, 0) = b
   End If
  Loop Until storevar(i, 1) = a5
 dd = CurDir & "\" & "Var.dat"
 Open dd For Output As #1
 For s = 0 To 26
 Print #1, storevar(s, 0)
 Next s
 Close #1
 End Sub
Private Sub call1()
Dim C As String * 1
dd = CurDir & "\" & "Minic.dat"
Open dd For Output As #1
Print #1, Text3.Text
Close #1

Open dd For Input As #1
For s = 1 To Len(Text3.Text) + 1
k = k + 1
Seek #1, s
Input #1, C
exp1(k) = C
express(k) = C
 If check(C) Then
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
    exp1(k) = C
    express(k) = C
    Case Else
    k = k + 1
   End Select
End If
Next s
Close #1
End Sub
Private Sub save2()
    dd = CurDir & "\" & "Exp.dat"
    Open dd For Output As #1
    Print #1, Form4.Text1.Text
    Close #1
    dd = CurDir & "\" & "Exp1.dat"
    Open dd For Output As #1
    For s = 0 To Form4.List2.ListCount
    Print #1, Form4.List2.List(s)
    Next
    flagstore = True
    Close #1
    dd = CurDir & "\" & "callist.dat"
    Open dd For Output As #1
    For s = 0 To 10
    save1 = Form1.abs1(s).Caption
    Print #1, save1
    Next
    Close #1
End Sub
