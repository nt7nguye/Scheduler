VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form studentList 
   Caption         =   "Cool Boi"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10905
   LinkTopic       =   "Form2"
   ScaleHeight     =   8205
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SELECTING STUDENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   26
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox courseText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   9360
      TabIndex        =   25
      Text            =   "Course"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox courseText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   9360
      TabIndex        =   24
      Text            =   "Course"
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox courseText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   9360
      TabIndex        =   23
      Text            =   "Course"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox courseText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   9360
      TabIndex        =   22
      Text            =   "Course"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox courseText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   9360
      TabIndex        =   21
      Text            =   "Course"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox courseText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   9360
      TabIndex        =   20
      Text            =   "Course"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox courseText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   9360
      TabIndex        =   19
      Text            =   "Course"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox courseText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   9360
      TabIndex        =   18
      Text            =   "Course"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox teacherText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   6720
      TabIndex        =   17
      Text            =   "Teacher"
      Top             =   6360
      Width           =   2415
   End
   Begin VB.TextBox teacherText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   6720
      TabIndex        =   16
      Text            =   "Teacher"
      Top             =   5640
      Width           =   2415
   End
   Begin VB.TextBox teacherText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   6720
      TabIndex        =   15
      Text            =   "Teacher"
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox teacherText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   6720
      TabIndex        =   14
      Text            =   "Teacher"
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox teacherText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   6720
      TabIndex        =   13
      Text            =   "Teacher"
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox teacherText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   6720
      TabIndex        =   12
      Text            =   "Teacher"
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox teacherText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   6720
      TabIndex        =   11
      Text            =   "Teacher"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox teacherText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   6720
      TabIndex        =   10
      Text            =   "Teacher"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox perText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   5760
      TabIndex        =   9
      Text            =   "H"
      Top             =   6360
      Width           =   615
   End
   Begin VB.TextBox perText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   5760
      TabIndex        =   8
      Text            =   "G"
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox perText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   5760
      TabIndex        =   7
      Text            =   "F"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox perText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   5760
      TabIndex        =   6
      Text            =   "E"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox perText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   5760
      TabIndex        =   5
      Text            =   "D"
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox perText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   5760
      TabIndex        =   4
      Text            =   "C"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox perText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   5760
      TabIndex        =   3
      Text            =   "B"
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox perText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5760
      TabIndex        =   2
      Text            =   "A"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox studentText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   1
      Text            =   "Student"
      Top             =   360
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid studentListGrid 
      Height          =   6375
      Left            =   720
      TabIndex        =   0
      Top             =   1200
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   11245
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "studentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
studentList.Visible = True
Open "C:\Users\Tan Nguyen\Desktop\schedule\datas\student.txt" For Input As #1
cRow = -1
Do Until EOF(1)
    Input #1, temp$
    cRow = cRow + 1
    studentListGrid.Rows = studentListGrid.Rows + 1
    studentListGrid.TextMatrix(cRow, 0) = temp$
Loop
Close #1
studentListGrid.ColWidth(0) = 4000
studentListGrid.ColWidth(1) = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
CHOICELIST.Enabled = True
End Sub

Private Sub studentListGrid_Click()
Row = studentListGrid.Row
studentListGrid.Enabled = False
Open "C:\Users\Tan Nguyen\Desktop\schedule\datas\studentSched.txt" For Input As #2
If Row <> 0 Then
    For i = 0 To Row - 1
        For j = 0 To 8
            Input #2, garb$
        Next
    Next
End If
Input #2, student$
studentText.Text = student$
For k = 0 To 7
    Input #2, temp$
    teacherText(k) = Left(temp$, Len(temp$) - 6)
    courseText(k) = Mid(temp$, Len(temp$) - 5, 5)
Next
studentListGrid.Enabled = True
Close #2
End Sub

