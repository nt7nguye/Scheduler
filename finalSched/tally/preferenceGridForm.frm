VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form preferenceGridForm 
   Caption         =   "Form2"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10260
   LinkTopic       =   "Form2"
   ScaleHeight     =   7920
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ConfirmBut 
      BackColor       =   &H0000FF00&
      Caption         =   "Confirm Choice'"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
      Width           =   3135
   End
   Begin VB.TextBox instructionsText 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Text            =   "Place an 'X' in the period where you would prefere a spare"
      Top             =   360
      Width           =   4695
   End
   Begin MSFlexGridLib.MSFlexGrid preferenceGrid 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9763
      _Version        =   393216
      Rows            =   100
      Cols            =   9
   End
End
Attribute VB_Name = "preferenceGridForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim teachers$(100)
Dim teacherNumber As Integer



Private Sub ConfirmBut_Click()
'go through the whole thing and print it all into a text file like so:
'teacher 0 0 0 0 0 0 0 0 (0 is spare you get by 'X'
Open "C:\Users\youse\OneDrive\Desktop\datas\prefDat.txt" For Output As #2
For i = 1 To preferenceGrid.Rows - 1
    Print #2, preferenceGrid.TextMatrix(i, 0)
    For j = 1 To 8
    If preferenceGrid.TextMatrix(i, j) = "X" Then
        Print #2, 0
    Else
        Print #2, 1
    End If
    
    Next
Next
Close #2
Load Form1
Unload preferenceGridForm
End Sub
Private Sub formattable()
With preferenceGrid
    .ColWidth(0) = 1500

End With
End Sub
Private Sub Form_Load()
Call formattable
preferenceGrid.Rows = 1
preferenceGrid.TextMatrix(0, 0) = "Teachers"

For i = 1 To 8
preferenceGrid.ColAlignment(i) = flexAlignCenterCenter
preferenceGrid.TextMatrix(0, i) = Chr$(64 + i)
Next
Open "C:\Users\youse\OneDrive\Desktop\datas\teacher.txt" For Input As #1
teacherNumber = 1
Do Until EOF(1)
    preferenceGrid.Rows = preferenceGrid.Rows + 1
    Input #1, teachers$(teacherNumber)
    For i = 0 To 3
        
    Next
    preferenceGrid.TextMatrix(teacherNumber, 0) = teachers$(teacherNumber)
    teacherNumber = teacherNumber + 1
Loop
Close #1

End Sub

Private Sub preferenceGrid_Click()
Row = preferenceGrid.Row
Col = preferenceGrid.Col
If preferenceGrid.TextMatrix(Row, Col) = "" Then
    preferenceGrid.TextMatrix(Row, Col) = "X"
ElseIf preferenceGrid.TextMatrix(Row, Col) = "X" Then
    preferenceGrid.TextMatrix(Row, Col) = ""
End If

End Sub

