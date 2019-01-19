VERSION 5.00
Begin VB.Form TeacherEdit 
   Caption         =   "Teacher"
   ClientHeight    =   11175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15630
   LinkTopic       =   "Form1"
   ScaleHeight     =   11175
   ScaleWidth      =   15630
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox TextTeacher 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox TextSearch 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "Search..."
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "TeacherEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub TextSearch_Click()
If TextSearch.Text = "Search..." Then
    TextSearch.Text = ""
End If
End Sub

Private Sub TextSearch_KeyPress(KeyAscii As Integer)

If TextSearch.Text = "Search..." Then
    TextSearch.Text = ""
End If

If KeyAscii = 8 And Len(TextSearch.Text) = 1 Then
    TextSearch.Text = "Search..."
End If
End Sub


