VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "`"
   ClientHeight    =   11355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   11355
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   495
      Left            =   2160
      TabIndex        =   20
      Top             =   9720
      Width           =   1575
   End
   Begin VB.Frame STUDENTDATA 
      Caption         =   "STUDENT DATA"
      Height          =   1815
      Left            =   480
      TabIndex        =   16
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton OEN 
         Caption         =   "OEN"
         Height          =   615
         Left            =   360
         TabIndex        =   19
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox Search 
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Text            =   "Search..."
         Top             =   360
         Width           =   4455
      End
      Begin VB.CommandButton GRADE 
         Caption         =   "GRADE"
         Height          =   615
         Left            =   2760
         TabIndex        =   17
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   7
      Left            =   3120
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   8640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   6
      Left            =   3120
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   7680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   5
      Left            =   3120
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   6720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   4
      Left            =   3120
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   3
      Left            =   3120
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   2
      Left            =   3120
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   1
      Left            =   3120
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Index           =   7
      Left            =   1200
      TabIndex        =   8
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Index           =   6
      Left            =   1200
      TabIndex        =   7
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Index           =   5
      Left            =   1200
      TabIndex        =   6
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Index           =   4
      Left            =   1200
      TabIndex        =   5
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Index           =   3
      Left            =   1200
      TabIndex        =   4
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   0
      Left            =   3120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form1.Visible = True
For X = 0 To 7
Command1(X).Caption = "Course" & Str(X + 1)
Text1(X).Text = ""
Next
Search.Text = Form2.Search.Text
End Sub

