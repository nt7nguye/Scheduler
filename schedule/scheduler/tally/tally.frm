VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter(100)
Dim class(100)
Dim course$(100)
Dim teacher$(100)
Dim courseTeach(100, 100) As Boolean
Dim j As Integer

Private Sub Form_Load()
Open "G:\ICS4\schedule\datas\course.txt" For Input As #1
j = 0
Do Until EOF(1)
Input #1, course$(j)
j = j + 1
Loop
Close #1

Open "G:\ICS4\schedule\datas\studentDat.txt" For Input As #2

Do Until EOF(2)
Input #2, inpers$
Print inpers$
For i = 0 To j
If Right(inpers$, 5) = course$(i) Then
    counter(i) = counter(i) + 1
End If
Next
Loop
Close #2


'remove this in final version
Open "G:\ICS4\schedule\datas\tallyDat.txt" For Output As #3

For i = 0 To j
    If counter(i) = "" Then
        counter(i) = 0
    End If
    If counter(i) > 7 And counter(i) < 17 Then
    class(i) = 1
    Else
    class(i) = Int(counter(i) / 17)
    End If
'remove this in final version
    Print #3, course$(i)
    Print #3, counter(i)
    Print #3, class(i)
Next

'-----------------------------------------Teacher Tally-----------------------------------------

Open "G:\ICS4\schedule\datas\Teachercourses.txt" For Input As #4
Open "G:\ICS4\schedule\datas\tallyTeach.txt" For Output As #5
teachNum = 0
Do Until EOF(4)
    Input #4, tempCourse$
    Input #4, tempTeach$
    If tempTeach$ <> teacher$(teachNum) Then
       teachNum = teachNum + 1
        teacher$(teachNum) = tempTeach$
    End If
    For i = 0 To j
        If tempCourse$ = course$(i) Then
            courseTeach(teachNum, i) = True
       End If
    Next
Loop
Close #4


For i = 1 To j
    Print #5, teacher$(i)
    courseCount = 0
    For k = 0 To j
        If courseTeach(i, k) = True Then
            Print #5, course$(k)
            courseCount = courseCount + 1
        End If
    Next
    If teacher$(i) <> "" Then
    For l = courseCount To 3
    Print #5, "."
    Next
    End If
Next
Close #5


Close #3
End Sub


