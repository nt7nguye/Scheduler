VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   13080
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox comboroom 
      Height          =   315
      Left            =   10800
      TabIndex        =   7
      Text            =   "Room#"
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox combohallway 
      Height          =   315
      Left            =   7080
      TabIndex        =   6
      Text            =   "Rooms by subject"
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox combolocation 
      Height          =   315
      Left            =   9000
      TabIndex        =   5
      Text            =   "Room #/Location"
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox comboyousef 
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Text            =   "Course Code"
      Top             =   0
      Width           =   1335
   End
   Begin VB.CheckBox add 
      Caption         =   "Add course"
      Height          =   195
      Left            =   5640
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.CheckBox clear 
      Caption         =   "Clear function"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.ComboBox combotan 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Text            =   "Departments"
      Top             =   0
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   12015
      Left            =   480
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
Dim ap$(3000)
Dim course$(75)
Dim department$(75)
Dim section$(35)
Dim teacher$(53)

Private Sub add_Click()
If clear.Value = 1 Then
clear.Value = 0
End If
End Sub

Private Sub clear_Click()
If add.Value = 1 Then
add.Value = 0
End If
End Sub

Private Sub combohallway_Click()
 combolocation.clear
 
        Open "H:\Teacher grid\rooms\" + combohallway.Text + ".txt" For Input As #10
        If combohallway.Text = "GeneralPurpose" Then comboroom.Visible = True
        If combohallway.Text <> "GeneralPurpose" Then comboroom.Visible = False
            For q = 1 To 100
                Input #10, ap$(q)
                If ap$(q) = "" Then
                Exit For
            End If
                combolocation.AddItem (ap$(q))
                combolocation.Text = ap$(1)
            Next
        Close #10

End Sub


Private Sub combolocation_Click()
If combohallway.Text = "GeneralPurpose" Then
comboroom.clear

        Open "H:\Teacher grid\rooms\" + combolocation.Text + ".txt" For Input As #9
             For q = 1 To 100
            Input #9, ap$(q)
            If ap$(q) = "" Then
            Exit For
            End If
                comboroom.AddItem (ap$(q))
                comboroom.Text = ap$(1)
            Next
        Close #9
End If

End Sub

Private Sub combotan_Click()

 comboyousef.clear
        Open "H:\Teacher grid\departments\" + combotan.Text + ".txt" For Input As #9
             For q = 1 To 100
            Input #9, ap$(q)
            If ap$(q) = "" Then
            Exit For
            End If
                comboyousef.AddItem (ap$(q))
                comboyousef.Text = ap$(1)
            Next
        Close #9
End Sub
Private Sub Form_Load()

For subject = 1 To 8
grid.TextMatrix(0, subject) = Chr$(64 + subject)
grid.TextMatrix(0, 9 + subject) = Chr$(64 + subject)
Next

Open "H:\Teacher grid\teacher.txt" For Input As #1
For a = 1 To 52
Input #1, teacher$(a)
grid.TextMatrix(a, 0) = teacher$(a)
Next
Close #1

Open "H:\Teacher grid\departments\departments.txt" For Input As #1
For a = 1 To 13
Input #1, department$(a)
combotan.AddItem (department$(a))
Next
Close #1

Open "H:\Teacher grid\rooms\sections.txt" For Input As #12
For a = 1 To 7
Input #12, section$(a)
combohallway.AddItem (section$(a))
Next
Close #12



End Sub
Private Sub grid_Click()
Row = grid.Row
Col = grid.Col

If add.Value = 1 Then
    If Col < 9 And Col > 1 Then
        If comboyousef.Text <> "Course Code" Then
        grid.TextMatrix(Row, Col) = comboyousef.Text
        End If
    End If
    
    If Col < 18 And Col > 9 Then
        If combohallway.Text = "GeneralPurpose" Then
            If comboroom.Text <> "Room#" Then
            grid.TextMatrix(Row, Col) = comboroom.Text
            End If
        End If
        If combohallway.Text <> "GeneralPurpose" Then
            If combolocation.Text <> "Room #/Location" Then
            grid.TextMatrix(Row, Col) = combolocation.Text
            End If
        End If
    End If
End If
If clear.Value = 1 Then
grid.TextMatrix(Row, Col) = ""
End If
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'grid.TextMatrix(x, y) = combotan.Text
End Sub
