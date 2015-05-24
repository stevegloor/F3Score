VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExportForm 
   AutoRedraw      =   -1  'True
   Caption         =   "Report Selection"
   ClientHeight    =   5235
   ClientLeft      =   780
   ClientTop       =   975
   ClientWidth     =   4455
   HelpContextID   =   340
   LinkTopic       =   "Form3"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5235
   ScaleWidth      =   4455
   Begin VB.CommandButton Command4 
      Caption         =   "Export Report to CSV File"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   3360
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   4560
      WhatsThisHelpID =   340
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Report Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   360
      WhatsThisHelpID =   340
      Width           =   2415
      Begin VB.OptionButton Option1 
         Caption         =   "Full Contest Leader Board"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Round/Slot Allocation "
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Contest Flight Groups"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Full Contest Flight Scores"
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2055
      End
   End
End
Attribute VB_Name = "ExportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Selected As Recordset
Dim ReportType As Integer
Dim Pagelength As Integer
Dim OutputFile As String

Private Sub Check1_Click(Index As Integer)
  ReportType = Index
  
End Sub


Private Sub Command3_Click()
  Unload Me
  
End Sub


Private Sub Command4_Click()
Select Case ReportType
    Case 1
      CommonDialog2.Filter = "CSV Files (*.csv)|*.csv"
      CommonDialog2.FilterIndex = 2
      CommonDialog2.ShowSave
      OutputFile = CommonDialog2.FileName
      If OutputFile = "" Then
      MsgBox ("no file selected")
       Exit Sub
      End If
      If CurrentContestType = "F3B" Then
        Currenttask = InputBox("Enter the Task that you want printed" + Chr$(13) + "Enter A, B or C")
        If Currenttask = "A" Or Currenttask = "a" Then
          Currenttask = "A"
          NumSlots = F3BNumSlots(1)
        ElseIf Currenttask = "B" Or Currenttask = "b" Then
          Currenttask = "B"
          NumSlots = F3BNumSlots(2)
        Else
          Currenttask = "C"
          NumSlots = F3BNumSlots(3)
        End If
        ExportContestCSV
      Else
        Currenttask = "A"
        ExportContestCSV
      End If
    Case 0
      MsgBox ("Not exportable to CSV")
    Case 2
      CommonDialog2.Filter = "CSV Files (*.csv)|*.csv"
      CommonDialog2.FilterIndex = 2
      CommonDialog2.ShowSave
      OutputFile = CommonDialog2.FileName
      If OutputFile = "" Then
      MsgBox ("no file selected")
       Exit Sub
      End If
      ExportSummaryCSV
    Case 3
      CommonDialog2.Filter = "CSV Files (*.csv)|*.csv"
      CommonDialog2.FilterIndex = 2
      CommonDialog2.ShowSave
      OutputFile = CommonDialog2.FileName
      If OutputFile = "" Then
      MsgBox ("no file selected")
       Exit Sub
      End If
      ExportLeaderBoardCSV
    Case 4
      CommonDialog2.Filter = "CSV Files (*.csv)|*.csv"
      CommonDialog2.FilterIndex = 2
      CommonDialog2.ShowSave
      OutputFile = CommonDialog2.FileName
      If OutputFile = "" Then
      MsgBox ("no file selected")
       Exit Sub
      End If
      ExportFlightGroupsCSV
    Case 5
      MsgBox ("Not exportable to CSV")
    Case 6
      MsgBox ("Not exportable to CSV")
    Case 7
      CommonDialog2.Filter = "CSV Files (*.csv)|*.csv"
      CommonDialog2.FilterIndex = 2
      CommonDialog2.ShowSave
      OutputFile = CommonDialog2.FileName
      If OutputFile = "" Then
      MsgBox ("no file selected")
       Exit Sub
      End If
      ExportFlightGroupsCSV
  End Select
End Sub

Private Sub Form_Load()
  Pagelength = 75
End Sub

Private Sub Option1_Click(Index As Integer)
  ReportType = Index
  
End Sub



Public Sub ExportContestCSV()
Dim SQLString3 As String
Dim SelectedSlot As Recordset
Dim RoundNum As Integer
Dim Slot As Integer
Dim Linecount As Integer
Dim ArryLine() As String
Dim Currentslot As Integer
Dim Count As Integer

 Open OutputFile For Output As #1
 
 On Error GoTo Errorhandler
 
 Linecount = 0
 Slot = 1
 Currentslot = 0
 ReDim ArryLine(NumRounds * NumSlots + 20)
 For RoundNum = 1 To NumRounds
     'SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & "Str(Round)")) ORDER BY Pilot.Last_Name;"
     Do Until Slot > (NumSlots + 5)
     SQLString3 = "SELECT DISTINCTROW List.Comp_ID, Pilot.First_Name, Pilot.Last_Name, List.Pilot_ID, Scores.Round, Scores.Slot, Scores.Task, List.FreqUsed, List.Team_ID FROM Pilot INNER JOIN ((Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID)) ON (Pilot.Pilot_ID = List.Pilot_ID) AND (Pilot.Pilot_ID = Scores.Pilot_ID) Where ((List.Comp_ID = " & Str(CurrentContest) & ") And (Scores.Round = " & Str(RoundNum) & ") And (Scores.Slot = " & Str(Slot) & ") AND (Scores.Task = '" & Currenttask & "')) ORDER BY Pilot.Last_Name ASC;"
     Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
     If SelectedSlot.RecordCount = 0 Then
       Exit Do
     End If
     SelectedSlot.MoveFirst
     Do Until SelectedSlot.EOF
       If Slot <> Currentslot Then
         
         Currentslot = Currentslot + 1
         Linecount = Linecount + 1
         ArryLine(Linecount) = ""
         ArryLine(Linecount) = ArryLine(Linecount) & "Round" & Str(RoundNum) & "," & "Slot" & Str(Slot) & ","
       End If
       ArryLine(Linecount) = ArryLine(Linecount) & Trim(SelectedSlot!First_Name) & " " & SelectedSlot!Last_Name & ","
       'Debug.Print Left(Trim(SelectedSlot!First_Name), 1); "."; Left(SelectedSlot!Last_Name, 10)
       SelectedSlot.MoveNext
     Loop
     Slot = Slot + 1
   Loop
   Slot = 1
   Currentslot = 0
 Next RoundNum

 For Count = 1 To Linecount
   Print #1, ArryLine(Count)
 Next Count
 MsgBox ("Report Exported OK")
 SelectedSlot.Close
 Close #1
Exit Sub
Errorhandler:
  MsgBox ("You have not entered any pilots into your contest or you have not allocated the rounds")
End Sub

Public Sub ExportFlightGroupsCSV()
Dim SelectedSlot As Recordset
Dim SQLString3 As String
Dim NumPilots As Integer
Dim CurrentGroup As Integer
Dim Linecount As Integer
Dim RoundNum
Dim Slot
Dim Search As String
Dim Currentslot As Integer
Dim Count As Integer

 Open OutputFile For Output As #1
 On Error GoTo Errorhandler
 Linecount = 0
 Slot = 1
 Currentslot = 0
 CurrentGroup = 0
 ReDim ArryLine(NumRounds * NumSlots * 5)
 If FlightGroups = False Then
   MsgBox ("Task B Flight groups not set up")
   Exit Sub
 Else
   For RoundNum = 1 To NumRounds
     Do While Slot < NumSlots + 5
       SQLString3 = "SELECT DISTINCTROW List.Comp_ID, Pilot.First_Name, Pilot.Last_Name, List.Pilot_ID, Scores.Round, Scores.Slot, Scores.Task, Scores.FlightGroup, List.FreqUsed, List.Team_ID FROM Pilot INNER JOIN ((Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID)) ON (Pilot.Pilot_ID = List.Pilot_ID) AND (Pilot.Pilot_ID = Scores.Pilot_ID) Where ((List.Comp_ID = " & Str(CurrentContest) & ") And (Scores.Round = " & Str(RoundNum) & ") And (Scores.Slot = " & Str(Slot) & ") AND (Scores.Task = 'B')) ORDER BY Scores.FlightGroup ASC;"
       CurrentGroup = 0
       Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
       If SelectedSlot.RecordCount = 0 Then
         Exit Do
       End If
       SelectedSlot.MoveFirst
       Do Until SelectedSlot.EOF
         If Slot <> Currentslot Then
           Debug.Print ArryLine(Linecount)
           Currentslot = Currentslot + 1
           Linecount = Linecount + 1
           ArryLine(Linecount) = ""
           CurrentGroup = SelectedSlot!FlightGroup
           ArryLine(Linecount) = ""
           ArryLine(Linecount) = ArryLine(Linecount) & "Round" & Str(RoundNum) & "," & "Slot" & Str(Slot) & "," & "Group" & Str(CurrentGroup) & ","
         End If
         If CurrentGroup <> SelectedSlot!FlightGroup Then
           Debug.Print ArryLine(Linecount)
           Linecount = Linecount + 1
           CurrentGroup = SelectedSlot!FlightGroup
           ArryLine(Linecount) = ""
           ArryLine(Linecount) = ArryLine(Linecount) & "Round" & Str(RoundNum) & "," & "Slot" & Str(Slot) & "," & "Group" & Str(CurrentGroup) & ","
         End If
         ArryLine(Linecount) = ArryLine(Linecount) & Trim(SelectedSlot!First_Name) & " " & SelectedSlot!Last_Name & ","
         SelectedSlot.MoveNext
       Loop
     Slot = Slot + 1
     Loop
   Slot = 1
   Currentslot = 0
   Next RoundNum
   For Count = 1 To Linecount
     Print #1, ArryLine(Count)
   Next Count
   MsgBox ("Report Exported OK")
   Close #1
   
   SelectedSlot.Close
 End If
Exit Sub
Errorhandler:
  MsgBox ("You have not entered any pilots into your contest or you have not allocated the rounds")
End Sub

Public Sub ExportSummaryCSV()
Dim SQLString3 As String

'SQLString = "SELECT Pilot.First_Name, Pilot.Last_Name, Pilot.Res1, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.LandBonus, Scores.Penalty, Scores.Score, Scores.Task, Scores.RawScore, Scores.Res2, Scores.Comp_ID FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID Where (((Scores.Comp_ID) = " & Str(CurrentContest) & ")) ORDER BY Scores.Round, Scores.Slot, Scores.Task;"

Dim SelectedSlot As Recordset
Dim RoundNum As Integer
Dim Slot As Integer
Dim Linecount As Integer
Dim ArryLine() As String
Dim Currentslot As Integer
Dim Count As Integer
Dim Currenttask As String
Dim Mins As Integer
Dim Secs As Single
 
 On Error GoTo Errorhandler
 Open OutputFile For Output As #1
 SQLString3 = "SELECT Pilot.First_Name, Pilot.Last_Name, Pilot.Res1, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.LandBonus, Scores.Penalty, Scores.Score, Scores.Task, Scores.RawScore, Scores.Res2, Scores.Comp_ID FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID Where (((Scores.Comp_ID) = " & Str(CurrentContest) & ")) ORDER BY Scores.Task ASC, Scores.Round ASC, Scores.Slot ASC, Scores.Score DESC;"
 Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
 ReDim ArryLine(SelectedSlot.RecordCount + 10)
 Linecount = 0
 ArryLine(Linecount) = ArryLine(Linecount) & "Task" & "," & "Round" & "," & "Slot" & "," & "First Name" & "," & "LastName" & "," & "Mins/secs" & "," & "Secs/Hund" & "," & "FlightTime" & "," & "Land Bonus" & "," & "Flight Pen" & "," & "Safety Pen" & "," & "RawScore" & "," & "Normal Score" & ","
 Linecount = 1
 SQLString3 = "SELECT Pilot.First_Name, Pilot.Last_Name, Pilot.Res1, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.LandBonus, Scores.Penalty, Scores.Score, Scores.Task, Scores.RawScore, Scores.Res2, Scores.Comp_ID FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID Where (((Scores.Comp_ID) = " & Str(CurrentContest) & ")) ORDER BY Scores.Task ASC, Scores.Round ASC, Scores.Slot ASC, Scores.Score DESC;"
 Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
 SelectedSlot.MoveFirst
 Do Until SelectedSlot.EOF
   If SelectedSlot!Round > NumRounds Then
     Exit Do
   End If
   If SelectedSlot.RecordCount = 0 Then
     Exit Do
   End If
   
   If SelectedSlot!task = "A" Then
     Mins = SelectedSlot!Flight_Time \ 60
     Secs = SelectedSlot!Flight_Time - ((SelectedSlot!Flight_Time \ 60) * 60)
   ElseIf SelectedSlot!task = "B" Then
     Mins = 0
     Secs = 0
   Else
     Mins = SelectedSlot!Flight_Time \ 100
     Secs = SelectedSlot!Flight_Time Mod 100
   End If
   ArryLine(Linecount) = ArryLine(Linecount) & SelectedSlot!task & "," & Str(SelectedSlot!Round) & "," & Str(SelectedSlot!Slot) & "," & Trim(SelectedSlot!First_Name) & "," & SelectedSlot!Last_Name & "," & Str(Mins) & "," & Format(Secs, "0#.0") & "," & Str(SelectedSlot!Flight_Time) & "," & Str(SelectedSlot!LandBonus) & "," & Str(SelectedSlot!Penalty) & "," & Str(SelectedSlot!Res2) & "," & Str(SelectedSlot!Rawscore) & "," & Str(SelectedSlot!Score) & ","
   'Debug.Print Left(Trim(SelectedSlot!First_Name), 1); "."; Left(SelectedSlot!Last_Name, 10)
   SelectedSlot.MoveNext
   Linecount = Linecount + 1
 Loop
 
 For Count = 0 To Linecount
   Print #1, ArryLine(Count)
 Next Count
 MsgBox ("Report Exported OK")
 SelectedSlot.Close
 Close #1
Exit Sub
Errorhandler:
  MsgBox ("You have not entered any pilots into your contest or you have not allocated the rounds")


End Sub

Public Sub ExportLeaderBoardCSV()
Dim Listset As Recordset
Dim Scoreset As Recordset
Dim SQLString1 As String
Dim rounds As Integer
Dim sco() As Single
Dim ArryLine() As String
Dim Count As Integer
Dim NumPilots As Integer
Dim tabset(20) As Integer
Dim Linecount As Integer
Dim Col As Integer
Dim P As Integer
Dim I As Integer

  rounds = NumRounds
  Call ScoreAddUp(sco(), NumPilots)
  ReDim ArryLine(NumPilots + 5)
  Count = 1
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Pilot.Res1 FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & ")) ORDER BY List.TempScore DESC;"
  Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Listset.MoveFirst
  'Print out the information
  Linecount = Linecount + 1
  Count = 1
  ArryLine(Linecount) = ArryLine(Linecount) & "," & "First Name" & "," & "Last Name" & "," & "Country" & "," & "Total Score" & ","
  If CurrentContestType <> "F3B" Then
    For I = 1 To NumRounds
      ArryLine(Linecount) = ArryLine(Linecount) & "R" & Str(I) & ","
    Next I
    ArryLine(Linecount) = ArryLine(Linecount) & "Throw1" & ","
    ArryLine(Linecount) = ArryLine(Linecount) & "Throw2"
    Linecount = Linecount + 1
  Else
    For I = 1 To NumRounds
      ArryLine(Linecount) = ArryLine(Linecount) & "R" & Str(I) & " Score" & ","
      ArryLine(Linecount) = ArryLine(Linecount) & "R" & Str(I) & "A" & ","
      ArryLine(Linecount) = ArryLine(Linecount) & "R" & Str(I) & "B" & ","
      ArryLine(Linecount) = ArryLine(Linecount) & "R" & Str(I) & "C" & ","
    Next I
    ArryLine(Linecount) = ArryLine(Linecount) & "ThrowA" & ","
    ArryLine(Linecount) = ArryLine(Linecount) & "ThrowB" & ","
    ArryLine(Linecount) = ArryLine(Linecount) & "ThrowC"
    Linecount = Linecount + 1
  End If
  Do Until Listset.EOF
    SQLString1 = "SELECT * From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Pilot_ID=" & Str(Listset!Pilot_ID) & ")) ORDER BY Scores.Round,Scores.Task ASC;"
    Set Scoreset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
    ArryLine(Linecount) = ArryLine(Linecount) & Str(Count) & "," & Listset!First_Name & "," & Listset!Last_Name & "," & Listset!Res1 & "," & Str(sco(Listset!Pilot_ID, 0))
    If CurrentContestType <> "F3B" Then
      For I = 1 To NumRounds
        ArryLine(Linecount) = ArryLine(Linecount) & "," & Str(sco(Listset!Pilot_ID, I))
        Col = Col + 1
      Next I
      ArryLine(Linecount) = ArryLine(Linecount) & "," & Str(throw(Listset!Pilot_ID, 0))
      ArryLine(Linecount) = ArryLine(Linecount) & "," & Str(throw(Listset!Pilot_ID, 1))
    Else
      Scoreset.MoveFirst
      For I = 1 To NumRounds
        ArryLine(Linecount) = ArryLine(Linecount) & "," & Str(sco(Listset!Pilot_ID, I)) & "," & Str(Scoreset!Score)
        Scoreset.MoveNext
        ArryLine(Linecount) = ArryLine(Linecount) & "," & Str(Scoreset!Score)
        Scoreset.MoveNext
        ArryLine(Linecount) = ArryLine(Linecount) & "," & Str(Scoreset!Score)
        Scoreset.MoveNext
      Next I
      ArryLine(Linecount) = ArryLine(Linecount) & "," & Str(throw(Listset!Pilot_ID, 0)) & ","
      ArryLine(Linecount) = ArryLine(Linecount) & Str(throw(Listset!Pilot_ID, 1)) & ","
      ArryLine(Linecount) = ArryLine(Linecount) & Str(throw(Listset!Pilot_ID, 2))
    End If
    Listset.MoveNext
    Count = Count + 1
    Linecount = Linecount + 1
    
  Loop
  Open OutputFile For Output As #1
  For Count = 1 To Linecount
    Print #1, ArryLine(Count)
  Next Count
  MsgBox ("Report Exported OK")
  Close #1
  
 SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, List.TempScore From List WHERE (List.Comp_ID=" & Str(CurrentContest) & ") ORDER BY List.Pilot_ID;"
 Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
 Listset.MoveFirst
 Do Until Listset.EOF
   Listset.Edit
   Listset!TempScore = 0
   Listset.Update
   Listset.MoveNext
 Loop
 Scoreset.Close
 Listset.Close
End Sub
