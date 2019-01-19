VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form classList 
   Caption         =   "Cool Boi"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9135
   LinkTopic       =   "Form2"
   ScaleHeight     =   7815
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid classGrid 
      Height          =   7455
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   13150
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
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
   Begin VB.TextBox courseText 
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
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox periodText 
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
      Left            =   120
      TabIndex        =   1
      Text            =   "Period: "
      Top             =   960
      Width           =   3015
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "classList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
classList.Visible = True
Open "C:\Users\Tan Nguyen\Desktop\finalSched\datas\tempClassDat.txt" For Input As #1
Input #1, teacher$
Input #1, course$
Input #1, period$
Close #1
teacherText.Text = teacher$
courseText.Text = course$
periodText.Text = periodText.Text + period$
Open "C:\Users\Tan Nguyen\Desktop\finalSched\datas\classListDat.txt" For Input As #2
Do Until temp$ = teacher$
    Input #2, temp$
Loop

Do Until temp$ = course$
    Input #2, temp$
Loop
cRow = 0
cCol = 0
While temp$ <> "."
    Input #2, temp$
    If temp$ <> "." Then
        classGrid.TextMatrix(cRow, cCol) = temp$
        If cCol = 1 Then
            cCol = 0
            cRow = cRow + 1
            classGrid.Rows = classGrid.Rows + 1
            Else
            cCol = 1
        End If
    End If
Wend
Close #2
classGrid.ColWidth(0) = 2500
classGrid.ColWidth(1) = 2500

End Sub


