VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CHOICELIST 
   Caption         =   "Cool Boi"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18555
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleWidth      =   18555
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "BACK TO RECORDS"
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
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox logText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   14280
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   6855
   End
   Begin VB.Frame logFrame 
      Caption         =   "SCHEDULE LOGS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   14160
      TabIndex        =   3
      Top             =   1080
      Width           =   7095
   End
   Begin VB.CommandButton studentListBut 
      Caption         =   "STUDENT LIST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   16440
      TabIndex        =   1
      Top             =   9840
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   11055
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   19500
      _Version        =   393216
      Rows            =   1
      Cols            =   9
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
Attribute VB_Name = "CHOICELIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter(100) 'used when counting up students for class#
Dim class(100) 'how many classes each course has, shares the index of the course it reps
Dim assignedCourse(100) 'used for balancing teacher schedule, counts courses teacher has been assigned, shares teacher index
Dim teachIndexList(10) 'stores a list of teacher indexes used for
Dim teachPrefList(100, 7) As Integer 'storage for teacher's preference for spares
Dim student$(1000) 'stores all of the student names
Dim course$(100) 'stores all of the course names
Dim teacher$(100) 'stores all of the teacher names
Dim studInClass(100, 7) As Integer 'teacher, period (stores the # of students in the teachers class at that period)
Dim studentSched$(1000, 7) 'student, period (stores the specific class code of the student at that period)
Dim studentCourse(1000, 100) As Boolean 'student, course (stores the courses that a student has)
Dim courseTeach(100, 100) As Boolean ' teacher, course (stores the courses that a teacher teaches)
Dim teachClassCount(100, 100) As Integer 'teacher, course (stores how many classes for a given course a teacher has)
Dim courseCount As Integer 'integer representation of the amount of courses
Dim teachCount As Integer 'integer representation of the amount of teachers
Dim studentCount As Integer 'integer representation of the amount of students

Private Sub Command1_Click()
RECORDS.Show
End Sub

Private Sub Form_Load()
CHOICELIST.Visible = True
'------------------------------------Find # of classes for each course based on student data-----------------------------
'"G:\ICS4\schedule\datas\course.txt"
Open "C:\Users\Tan Nguyen\Desktop\finalSched\datas\course.txt" For Input As #1
courseCount = 0
Do Until EOF(1)
    Input #1, course$(courseCount)
    courseCount = courseCount + 1
Loop
Close #1
courseCount = courseCount - 1
'C:\Users\youse\OneDrive\Desktop\
'G:\ICS4\schedule\datas\studentDat.txt
Open "C:\Users\Tan Nguyen\Desktop\finalSched\datas\studentDat.txt" For Input As #2
Do Until EOF(2)
    Input #2, inpers$
    For i = 0 To courseCount
        If Right(inpers$, 5) = course$(i) Then
            counter(i) = counter(i) + 1
        End If
    Next
Loop
Close #2
'finding out how many classes for each course
For i = 0 To courseCount
    If counter(i) = "" Then
        counter(i) = 0
    End If
    If counter(i) > 7 And counter(i) < 17 Then
        class(i) = 1
    Else
        class(i) = Int(counter(i) / 17)
    End If
'class and counter correspond to courses
Next
'-----------------------------------------Student Tally-----------------------------------------
Open "C:\Users\Tan Nguyen\Desktop\finalSched\datas\studentDat.txt" For Input As #2
studentCount = 0
Do Until EOF(2)
    Input #2, tempLast$
    Input #2, tempFirst$
    tempCourse$ = Right(tempFirst$, 5)
    tempStudent$ = tempLast$ + " " + Trim(Left(tempFirst$, Len(tempFirst$) - 6))
    
    For i = 0 To 1000
        If tempStudent$ = student$(i) Then
            For j = 0 To courseCount
                If tempCourse$ = course$(j) Then
                    studentCourse(i, j) = True
                    Exit For
                End If
            Next
            Exit For
        ElseIf student$(i) = "" Then
            student$(i) = tempStudent$
            If studentCount < i Then
                studentCount = i
            End If
            For j = 0 To courseCount
                If tempCourse$ = course$(j) Then
                    studentCourse(i, j) = True
                    Exit For
                End If
            Next
            Exit For
        End If
    Next
Loop
Close #2

Open "C:\Users\Tan Nguyen\Desktop\finalSched\datas\studentTally.txt" For Output As #4

For i = 0 To studentCount
    organizeCount = 1
    Print #4, student$(i)
    For j = 0 To courseCount
        If studentCourse(i, j) = True And organizeCount < 9 Then
            Print #4, course$(j)
            organizeCount = organizeCount + 1
        End If
    Next
    If organizeCount <> 9 Then
        For k = organizeCount To 9
        Print #4, "."
        Next
    End If
Next
Close #4

Open "C:\Users\Tan Nguyen\Desktop\finalSched\datas\student.txt" For Output As #9
For studentListIt = 0 To studentCount
    Print #9, student$(studentListIt)
Next
Close #9
'-----------------------------------------Teacher Tally-----------------------------------------
'"G:\ICS4\schedule\datas\Teachercourses.txt"
'"G:\ICS4\schedule\datas\tallyTeach.txt"
Open "C:\Users\Tan Nguyen\Desktop\finalSched\datas\Teachercourses.txt" For Input As #3
Open "C:\Users\Tan Nguyen\Desktop\finalSched\datas\tallyTeach.txt" For Output As #5

teachCount = 0

Input #3, tempCourse$
Input #3, teacher$(teachCount)

For i = 0 To courseCount
    If tempCourse$ = course$(i) Then
        courseTeach(teachCount, i) = True
        Exit For
    End If
Next

'teacher starts at 0
Do Until EOF(3)
    Input #3, tempCourse$
    Input #3, tempTeach$
    If tempTeach$ <> teacher$(teachCount) Then
        teachCount = teachCount + 1
        teacher$(teachCount) = tempTeach$
    End If
    For i = 0 To courseCount
        If tempCourse$ = course$(i) Then
            courseTeach(teachCount, i) = True
            Exit For
       End If
    Next
Loop
Close #3

'don't need tallyTeach anymore, new method doesnt access files keeping for data storage allocation stuff
For i = 0 To teachCount
    Print #5, teacher$(i)
    courseCounter = 0
    For k = 0 To courseCount
        If courseTeach(i, k) = True Then
            Print #5, course$(k)
            courseCounter = courseCounter + 1
        End If
    Next
    If teacher$(i) <> "" Then
        For j = courseCounter To 3
           Print #5, "."
        Next
    End If
Next
Close #5
'----------------------------------------Getting Teacher preferences----------------------------------------------
Open "C:\Users\Tan Nguyen\Desktop\finalSched\datas\prefDat.txt" For Input As #8
For i = 0 To teachCount
    Input #8, garb
    For j = 0 To 7
        Input #8, teachPrefList(i, j)
    Next
Next
Close #8

'----------------------------------------Assigning Teachers to Classes--------------------------------------------

Dim taughtCourse(100) As Boolean
For m = 0 To courseCount

    'find the course with the least amount of teachers, check it hasn't already been done
    For j = 0 To courseCount
        If taughtCourse(j) = False Then
            bestCourse = j
        End If
    Next
    
    For j = 0 To courseCount
        If getTeachCourseNum(course$(bestCourse)) > getTeachCourseNum(course$(j)) And taughtCourse(j) = False Then
            bestCourse = j
        End If
    Next
    taughtCourse(bestCourse) = True
    
    'find all of the teachers that teach the course
    TeachIt = 0
    For i = 0 To teachCount
        Dim TempdoesTeach$()
        TempdoesTeach$ = getTeachCourse(i)
        For teacherhaveCourseIt = 0 To 3
            If TempdoesTeach$(teacherhaveCourseIt) = course$(bestCourse) Then
                teachIndexList(TeachIt) = i
                TeachIt = TeachIt + 1
                Exit For
            End If
        Next
    Next
    TeachIt = TeachIt - 1
    
    tempClass = class(bestCourse)
    While tempClass > 0
        'check if you can assign a class without breaking preference
        canPref = False
        For i = 0 To TeachIt
            'find if anyone has classes + preferences < 8
            If getTotalClass(teachIndexList(i)) + getPrefNum(teachIndexList(i)) < 8 Then
                canPref = True
                bestTeachIndex = teachIndexList(i)
                Exit For
            End If
        Next
        
        'if we can assign classes without violating preference, then give the class to the person who has the
        'least classes + preference
        If canPref = True Then
            For i = 0 To TeachIt
                If getTotalClass(bestTeachIndex) + getPrefNum(bestTeachIndex) > getTotalClass(teachIndexList(i)) + getPrefNum(teachIndexList(i)) Then
                    bestTeachIndex = teachIndexList(i)
                ElseIf getTotalClass(bestTeachIndex) + getPrefNum(bestTeachIndex) = getTotalClass(teachIndexList(i)) + getPrefNum(teachIndexList(i)) And getTeachAmount(teachIndexList(i)) - assignedCourse(teachIndexList(i)) < getTeachAmount(bestTeachIndex) - assignedCourse(teachIndexList(i)) Then
                    bestTeachIndex = teachIndexList(i)
                End If
            Next
        Else
            'we must violate preference, or there will be a fatal error in the assigning of classes,
            'go thru each person, and find the one with the most preference, give them a course
            'then report it in the log text, and remove one of their preferences
            maxPref = 0
            For i = 0 To TeachIt
                If getPrefNum(teachIndexList(i)) > maxPref Then
                    bestTeachIndex = teachIndexList(i)
                    maxPref = getPrefNum(bestTeachIndex)
                End If
            Next
            'go thru their preference and take away the first one u see and then set them to
            For k = 0 To 7
                If teachPrefList(bestTeachIndex, k) = 0 Then
                    teachPrefList(bestTeachIndex, k) = 1
                    Exit For
                End If
            Next
            logText.Text = logText.Text + "Unfortunately the spare preferences of " + teacher$(bestTeachIndex) + " could not be met." & vbCrLf
        End If
        'you have found the best teacher for a class, give them a class
        teachClassCount(bestTeachIndex, bestCourse) = teachClassCount(bestTeachIndex, bestCourse) + 1
        tempClass = tempClass - 1
    Wend
    'you have assigned a course, give every teacher who had that course an assigned course
    For i = 0 To TeachIt
        assignedCourse(i) = assignedCourse(i) + 1
    Next
Next

'-----------------------------------------Putting Teachers into a teacher grid------------------------------------------

grid.Rows = 1
grid.TextMatrix(0, 0) = "Teachers"
'setting course columns
For i = 1 To 8
    grid.ColAlignment(i) = flexAlignCenterCenter
    grid.TextMatrix(0, i) = Chr$(64 + i)
Next

'putting the teachers in row ap
'teacher$ starts at 0

For i = 0 To teachCount
    grid.Rows = grid.Rows + 1
    grid.TextMatrix(i + 1, 0) = teacher$(i)
    'find their preference and give spares

    For j = 1 To 8
        If teachPrefList(i, j - 1) = 0 Then
            grid.TextMatrix(i + 1, j) = "SPARE"
        End If
    Next
Next
Call formattable

'now fill the grid with their courses considering the cost, and avoiding preferences

For i = 0 To grid.Rows - 2
    For j = 1 To 8
        If grid.TextMatrix(i + 1, j) <> "SPARE" Then
            For k = 0 To 100
                If teachClassCount(i, k) > 0 Then
                    Col = getBestCol(i + 1)
                    grid.TextMatrix(i + 1, Col) = course$(k)
                    teachClassCount(i, k) = teachClassCount(i, k) - 1
                    
                    Exit For
                End If
            Next
        End If
    Next
Next

'fill the blank spots with spares

For i = 0 To grid.Rows - 2
    For j = 1 To 8
        If grid.TextMatrix(i + 1, j) = "" Then
            grid.TextMatrix(i + 1, j) = "SPARE"
        End If
    Next
Next

'--------------------------------------------Filling Classes with Kiddies---------------------------------
'go thru list of students, look at all of their courses, find which one has least classes, go accross and assign the student to the class with the least students
'repeat and make sure it doesnt conflict with any of their other classes.
'assign them to a teacher's spare
'when you click on the grid, a form should open up with all of the students.
'classes are called: teacher code period (ex smithICS4UA)
'------------------------------------------Looking at them classes and seeing what's good-------------------------
'look at courses, see how many classes are available for that course find the min and assign that first
'repeat until all classes are assigned and just spares are left
'assign to the spares with the least amount of peeps
Dim tempStudentCourses$()
Dim tempStudentGivenCourse(7) As Boolean
Dim tempStudentGivenPer(7) As Boolean

'set up spare
course$(courseCount + 1) = "SPARE"
For i = 0 To teachCount
    For j = 1 To 8
        If grid.TextMatrix(i + 1, j) = "SPARE" Then
            class(courseCount + 1) = class(courseCount + 1) + 1
        End If
    Next
Next

For i = 0 To studentCount
    For k = 0 To 7
        tempStudentGivenCourse(k) = False
        tempStudentGivenPer(k) = False
    Next
    
    tempStudentCourses$ = getStudCourse(i)
    
    For j = 0 To 7
        
        'find course with lowest amount of classes
        bestCourseIndex = 0
        For k = 1 To 7
            'check if the thing is the lowest and the course isnt already assigned
            If class(getCourseIndex(tempStudentCourses$(bestCourseIndex))) > class(getCourseIndex(tempStudentCourses$(k))) And tempStudentGivenCourse(k) = False Then
                bestCourseIndex = k
            End If
        Next
        
        
        'find the class of that course with the least students that doesnt conflict with
        'studInClass(100,7) minimized find the course save best row and column
        
        minClassCount = 100
        bestRealCourse = getCourseIndex(tempStudentCourses$(bestCourseIndex))
        For Row = 0 To grid.Rows - 2
            For Col = 0 To 7
                If grid.TextMatrix(Row + 1, Col + 1) = course$(bestRealCourse) And studInClass(Row, Col) < minClassCount And tempStudentGivenPer(Col) = False Then
                    'check if it is less than the min, if it is set the bestRow and bestCol to the Row and Col
                    bestRow = Row
                    bestCol = Col
                    minClassCount = studInClass(Row, Col)
                End If
            Next
        Next
        'you found the one with the least amount of ppl in the class and that isnt taken up
        'now assign them to that class in their studentSched$(s, p) add a student to studInClass(t, p)
        'and update so the period and course they assigned are not able to be assigned again
        studentSched$(i, bestCol) = teacher$(bestRow) + course$(bestRealCourse) + Chr$(65 + bestCol)
        studInClass(bestRow, bestCol) = studInClass(bestRow, bestCol) + 1
        tempStudentGivenCourse(bestCourseIndex) = True
        tempStudentGivenPer(bestCol) = True
    Next
Next

'now you have studentSched showing who is in what class
'-----------------------------------------------make the class Lists--------------------------------------
Open "C:\Users\Tan Nguyen\Desktop\finalSched\datas\classListDat.txt" For Output As #6
For i = 0 To teachCount
    Print #6, teacher$(i)
    For k = 0 To 7
        Print #6, grid.TextMatrix(i + 1, k + 1)
        For j = 0 To studentCount
            If studentSched$(j, k) = teacher$(i) + grid.TextMatrix(i + 1, k + 1) + Chr$(65 + k) Then
                Print #6, student$(j)
            End If
        Next
        Print #6, "."
    Next
Next
Close #6
'---------------------------------------------make student scheduel---------------------------------------
'____________________________________________________________________________________________________
Open "C:\Users\Tan Nguyen\Desktop\finalSched\datas\studentSched.txt" For Output As #7
For i = 0 To studentCount
    Print #7, student$(i)
    For k = 0 To 7
        If studentSched$(i, k) <> "" Then
            Print #7, studentSched$(i, k)
        Else
            For j = 0 To grid.Rows - 2
                If grid.TextMatrix(j + 1, k + 1) = "SPARE" Then
                    Print #7, grid.TextMatrix(j + 1, 0) + "SPARE" + grid.TextMatrix(0, k + 1)
                    Exit For
                End If
            Next
        End If
    Next
Next
Close #7
'______________________________________________________________________________________________________
logText.Text = logText.Text + "Schedule Successfully Created"
End Sub
'----------------------------------Opening individual class schedules---------------------------------------------
Private Sub grid_click()
Row = grid.Row
Col = grid.Col
If Row > 0 And Col > 0 And Col < 9 Then
    'open the student form, make a temp file where the teacher, course code, and period are given
    Open "C:\Users\Tan Nguyen\Desktop\finalSched\datas\tempClassDat.txt" For Output As #12
    Print #12, teacher$(Row - 1) + "," + grid.TextMatrix(Row, Col) + "," + grid.TextMatrix(0, Col)
    Close #12
    classList.Show
End If
End Sub
'---------------------------------Method for finding the number of classes that are taught for a course------------
Function getCourseIndex(ByVal courseName As String) As Integer
    'find what course it is
    For getCourseIndexIt = 0 To courseCount + 1
        If courseName = course$(getCourseIndexIt) Then
            getCourseIndex = getCourseIndexIt
            Exit For
        End If
    Next
End Function
'--------------------------------------------Method for Getting preferred # of spares of a teacher---------------------
Function getPrefNum(ByVal teacherIndex As Integer) As Integer
getPrefNumCount = 0
For getPrefNumIt = 0 To 7
    If teachPrefList(teacherIndex, getPrefNumIt) = 0 Then
        getPrefNumCount = getPrefNumCount + 1
    End If
Next

If getPrefNumCount = 0 Then
    getPrefNumCount = 2
End If

getPrefNum = getPrefNumCount
End Function
'------------------------------------------Method that gets the number of courses a teacher teaches------------------
Function getTeachAmount(ByVal teacherIndex As Integer) As Integer
    teachAmount = 0
    Dim tempTeachAmount$()
    tempTeachAmount$ = getTeachCourse(teacherIndex)
    For teachAmountIt = 0 To 3
        If tempTeachAmount$(teachAmountIt) <> "." Then
            teachAmount = teachAmount + 1
        Else
            Exit For
        End If
    Next
    getTeachAmount = teachAmount
End Function
'---------------------------------Method for getting total # of classes a teacher has--------------------------------
Function getTotalClass(ByVal teacherIndex As Integer) As Integer
    totalClassCount = 0
    For getTotalClassIt = 0 To courseCount
        totalClassCount = totalClassCount + teachClassCount(teacherIndex, getTotalClassIt)
    Next
    getTotalClass = totalClassCount
End Function
'--------Method for finding # of teachers teaching a course----------------------------------------------------------
Function getTeachCourseNum(ByVal courseTaught As String) As Integer
    tempTeachCourseNum = 0
    Open "C:\Users\Tan Nguyen\Desktop\finalSched\datas\tallyTeach.txt" For Input As #7
    Do Until EOF(7)
        Input #7, garb$
        If garb$ = courseTaught Then
            tempTeachCourseNum = tempTeachCourseNum + 1
        End If
    Loop
    Close #7
    getTeachCourseNum = tempTeachCourseNum
End Function
'---------------------Method for finding teacher Courses-------------------------------------------------------------
Function getTeachCourse(ByVal teacherIndex As Integer) As String()
    Dim tempTeachCourse$(3)
    getTeachCourseIndex = 0
    For getTeachCourseIt1 = 0 To 3
        tempTeachCourse$(getTeachCourseIt1) = "."
    Next

    For getTeachCourseIt2 = 0 To courseCount
        If courseTeach(teacherIndex, getTeachCourseIt2) = True Then
            tempTeachCourse$(getTeachCourseIndex) = course$(getTeachCourseIt2)
            getTeachCourseIndex = getTeachCourseIndex + 1
        End If
    Next
    
    getTeachCourse = tempTeachCourse$
End Function
'---------------------Method for finding student Courses-------------------------------------------------------------
Function getStudCourse(ByVal studentIndex As Integer) As String()
    Dim tempStudCourse$(8)
    getStudCourseIndex = 0
    For getStudCourseIt1 = 0 To 8
        tempStudCourse$(getStudCourseIt1) = "SPARE"
    Next

    For getStudCourseIt2 = 0 To courseCount
        If studentCourse(studentIndex, getStudCourseIt2) = True Then
            tempStudCourse$(getStudCourseIndex) = course$(getStudCourseIt2)
            getStudCourseIndex = getStudCourseIndex + 1
        End If
    Next
    
    getStudCourse = tempStudCourse$
End Function
'----------------------------------------Method for finding lowest cost Column---------------------------
Function getBestCol(ByVal rowIndex As Integer) As Integer
    For getBestIt = 1 To 8
        If grid.TextMatrix(rowIndex, getBestIt) = "" Then
        bestColIndex = getBestIt
        End If
    Next
    
    For i = 1 To 8
        costOfColI = 0
        For getCostIt = 0 To grid.Rows - 2
            If grid.TextMatrix(getCostIt + 1, i) = "SPARE" Then
                costOfColI = costOfColI - 1
            ElseIf grid.TextMatrix(getCostIt + 1, i) <> "" Then
                costOfColI = costOfColI + 1
            End If
        Next
        
        costOfColB = 0
        For getCostIt = 0 To grid.Rows - 2
            If grid.TextMatrix(getCostIt + 1, bestColIndex) = "SPARE" Then
                costOfColB = costOfColB - 1
            ElseIf grid.TextMatrix(getCostIt + 1, bestColIndex) <> "" Then
                costOfColB = costOfColB + 1
            End If
        Next
        
        If costOfColB > costOfColI And grid.TextMatrix(rowIndex, i) = "" Then
            bestColIndex = i
        End If
    Next
    getBestCol = bestColIndex
    
End Function
'----------------------------------------Method for Finding the Cost of a space--------------------------
Function getCost(ByVal ColIndex As Integer) As Integer
    costOfCol = 0
    For getCostIt = 0 To grid.Rows - 2
        If grid.TextMatrix(getCostIt + 1, ColIndex) = "SPARE" Then
            costOfCol = costOfCol - 1
        ElseIf grid.TextMatrix(getCostIt + 1, ColIndex) <> "" Then
            costOfCol = costOfCol + 1
        End If
    Next
    getCost = costOfCol
End Function

Private Sub formattable()
With grid
    .ColWidth(0) = 1500
End With
End Sub

Private Sub studentListBut_Click()
 'open up a form with all of the students listed, and whenever you click it, access the student file and give their classes
studentList.Show
End Sub
