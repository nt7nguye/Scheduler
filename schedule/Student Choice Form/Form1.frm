VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form PLAN 
   Caption         =   "Form2"
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16005
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   16005
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton UPDATE 
      Appearance      =   0  'Flat
      Caption         =   "UPDATE RECORDS"
      Height          =   735
      Left            =   8760
      TabIndex        =   19
      Top             =   1320
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   7320
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Tan Nguyen\Desktop\Records.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Tan Nguyen\Desktop\Records.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from history"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   19440
      TabIndex        =   18
      Text            =   "40"
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox reqtext 
      Height          =   735
      Left            =   19440
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox taktext 
      Height          =   735
      Left            =   19440
      TabIndex        =   16
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton req 
      Caption         =   "CREDITS REQUIRED"
      Height          =   735
      Left            =   17520
      TabIndex        =   14
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton tak 
      Caption         =   "CREDITS TAKING"
      Height          =   735
      Left            =   17520
      TabIndex        =   13
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   " "
      Height          =   4455
      Left            =   17280
      TabIndex        =   11
      Top             =   3240
      Width           =   4335
      Begin VB.TextBox fintext 
         Height          =   735
         Left            =   2160
         TabIndex        =   15
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Fin 
         Caption         =   "CREDITS FINISHED"
         Height          =   735
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   3960
         Y1              =   3360
         Y2              =   3360
      End
   End
   Begin VB.CommandButton GRADE 
      Caption         =   "GRADE: ..."
      Height          =   615
      Index           =   1
      Left            =   7920
      TabIndex        =   10
      Top             =   2520
      Width           =   3255
   End
   Begin VB.CommandButton NAME 
      Caption         =   "NAME: ..."
      Height          =   615
      Index           =   0
      Left            =   1200
      TabIndex        =   9
      Top             =   2520
      Width           =   3255
   End
   Begin VB.CommandButton CHOICE 
      Appearance      =   0  'Flat
      Caption         =   "STUDENT CHOICE"
      Height          =   735
      Left            =   6360
      TabIndex        =   8
      Top             =   1320
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   7215
      Left            =   720
      TabIndex        =   7
      Top             =   3720
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   12726
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4105
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4105
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton ALL 
      Caption         =   "ALL"
      Height          =   615
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame STUDENTDATA 
      Caption         =   "STUDENT DATA"
      Height          =   2535
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   16095
      Begin VB.CommandButton Searchbutton 
         Caption         =   "Search"
         Height          =   495
         Left            =   5520
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Search 
         Height          =   735
         Left            =   360
         TabIndex        =   4
         Text            =   "Search..."
         Top             =   360
         Width           =   5055
      End
      Begin VB.CommandButton OEN 
         Caption         =   "OEN: ..."
         Height          =   615
         Left            =   3720
         TabIndex        =   3
         Top             =   1560
         Width           =   3255
      End
   End
   Begin VB.CommandButton COURSES 
      Caption         =   "COURSES"
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton PLAN 
      Caption         =   "PLAN"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PLAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Dim con As New ADODB.Connection
Private Sub SELECT_Click(Index As Integer)
Load Form1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Courses_Click()
Unload PLAN
Load COURSES
End Sub

Private Sub Form_Load()
Search.Locked = True
Call formattable
End Sub

Private Sub Search_Click()
Search.Locked = False
Search.Text = ""
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex _
As Integer)
Dim sortField As String
Dim sortString As String

sortField = DataGrid1.Columns(ColIndex).Caption
If InStr(Adodc1.Recordset.Sort, "Asc") Then
    sortString = sortField & " Desc"
Else
    sortString = sortField & " Asc"
End If
Adodc1.Recordset.Sort = sortString
End Sub

Private Sub formattable()
DataGrid1.Columns("OEN").Visible = False
'change this
With DataGrid1
    .Columns("Name").Width = 2000
    .Columns("Date").Width = 2000
    .Columns("Grade").Width = 2000
    .Columns("Course").Width = 2000
    .Columns("Average").Width = 2000
    .Columns("Credit").Width = 2000
End With

End Sub

Private Sub Searchbutton_Click()
Adodc1.RecordSource = "Select * from history where Name ='" + Search.Text + "'"
Adodc1.Refresh
Call formattable
Set OEN.DataSource = rec
OEN.DataField = "OEN"
End Sub

Private Sub Search_KeyPress(a As Integer)
If a = 13 Then Call Searchbutton_Click
End Sub

Private Sub All_Click()
Adodc1.RecordSource = "Select * from history"
Adodc1.Refresh
Call formattable
End Sub


