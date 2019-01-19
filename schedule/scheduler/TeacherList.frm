VERSION 5.00
Begin VB.Form TeacherList 
   Caption         =   "Teacher list"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ButConfirm 
      Caption         =   "Confirm"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton ButCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   6240
      Width           =   1575
   End
   Begin VB.VScrollBar ScrollTeacher 
      Height          =   5535
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
   Begin VB.ListBox ListTeacher 
      Height          =   5520
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "TeacherList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub
