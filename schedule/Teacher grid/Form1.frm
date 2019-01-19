VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox add 
      Caption         =   "Add course"
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin VB.CheckBox clear 
      Caption         =   "Clear function"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.ComboBox combotan 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Text            =   "Course Code"
      Top             =   0
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   12015
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   20415
      _ExtentX        =   36010
      _ExtentY        =   21193
      _Version        =   393216
      Rows            =   55
      Cols            =   20
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   20
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x(3000)
Dim y(3000)
Dim course$(75)
Dim teacher$(53)
Private Sub Form_Load()

Open "G:\ICS4\schedule\datas\course.txt" For Input As #1
For a = 1 To 75
Input #1, course$(a)
combotan.AddItem (course$(a))
Next a
Close #1

For subject = 1 To 8
grid.TextMatrix(0, subject) = Chr$(64 + subject)
grid.TextMatrix(0, 9 + subject) = Chr$(64 + subject)
Next

Open "G:\ICS4\schedule\datas\teacher.txt" For Input As #2
For a = 1 To 52
Input #2, teacher$(a)
grid.TextMatrix(a, 0) = teacher$(a)
Next
Close #2

End Sub

Private Sub grid_Click()
Row = grid.Row
Col = grid.Col
If combotan.Text <> "Course Code" Then
If add.Value = 1 Then
grid.TextMatrix(Row, Col) = combotan.Text
End If
End If
If clear.Value = 1 Then
grid.TextMatrix(Row, Col) = ""
End If
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'grid.TextMatrix(x, y) = combotan.Text
End Sub
