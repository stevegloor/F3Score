VERSION 5.00
Begin VB.Form DataMaint 
   ClientHeight    =   6720
   ClientLeft      =   585
   ClientTop       =   525
   ClientWidth     =   6750
   HelpContextID   =   390
   LinkTopic       =   "Form5"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6720
   ScaleWidth      =   6750
   Begin VB.CommandButton Command9 
      Caption         =   "Integrity Check Selected Contest"
      Height          =   495
      Left            =   3240
      TabIndex        =   15
      Top             =   5280
      Width           =   3135
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Set Selected Contest to Not Completed"
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   4320
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Delete Pilot"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      WhatsThisHelpID =   390
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Show Current"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   12
      Top             =   3840
      WhatsThisHelpID =   390
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Show Deleted"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      WhatsThisHelpID =   390
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   3000
      TabIndex        =   9
      Top             =   1200
      Width           =   3615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "DONE"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   6000
      WhatsThisHelpID =   390
      Width           =   3135
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Purge Selected Contest"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   4800
      WhatsThisHelpID =   390
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New Pilot"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      WhatsThisHelpID =   390
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Purge Selected Pilot From Database"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      WhatsThisHelpID =   390
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit Selected Pilots Details"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      WhatsThisHelpID =   390
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "** Denotes Currently Selected Contest"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "DataBase Maintenance Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "List of Contests in Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "List of Pilots in Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "DataMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub UpdateFrame2()
Dim TeamSelect As Recordset
Dim SQLString As String
Dim Line As String
 List2.Clear
 SQLString = "SELECT * FROM Competition ORDER BY CompStartDate ASC;"
 Set TeamSelect = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
 If TeamSelect.RecordCount <> 0 Then
   TeamSelect.MoveFirst
 End If
 Do Until TeamSelect.EOF = True
   Line = Format(Str(TeamSelect!Comp_ID), "!@@@@") & " " & TeamSelect!Comp_Name
   If TeamSelect!Comp_ID = CurrentContest Then
     Line = Line & " **"
   End If
   List2.AddItem Line
   TeamSelect.MoveNext
 Loop
TeamSelect.Close
End Sub

Public Sub UpdateFrame1()
Dim TeamSelect As Recordset
Dim SQLString As String
Dim Line As String
 List1.Clear
 If Option1(0) = True Then
   SQLString = "SELECT * FROM Pilot WHERE Current = false ORDER BY Pilot.Last_Name;"
 ElseIf Option1(1) = True Then
   SQLString = "SELECT * FROM Pilot WHERE Current = true ORDER BY Pilot.Last_Name;"
 End If
 Set TeamSelect = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
 If TeamSelect.RecordCount <> 0 Then
   TeamSelect.MoveFirst
 End If
 Do Until TeamSelect.EOF = True
   Line = Format(Str(TeamSelect!Pilot_ID), "!@@@@") & " " & TeamSelect!First_Name & " " & TeamSelect!Last_Name
   List1.AddItem Line
   TeamSelect.MoveNext
 Loop
TeamSelect.Close
End Sub

Private Sub Text1_Change()

End Sub


Private Sub Command1_Click()
  Dim PilotNum As String
  
  PilotNum = Trim(List1.Text)
  PilotNum = Left(PilotNum, 3)
  PilotNum = Trim(PilotNum)
  
  If PilotNum = "" Then
    MsgBox ("Nothing Selected")
  Else
    EditPilotID = PilotNum
    FromChange = True
    PilotEntry.Show
  End If
End Sub

Private Sub Command2_Click()
Dim PilotNum As String
Dim Msg As String
Msg = "You are about to delete all this pilots records." & Chr(13) & "This could affect the results of stored contests" & Chr(13) & "        Are you Sure?"

  If List1.Text <> "" Then
    If MsgBox(Msg, 4) = vbYes Then
      PilotNum = Trim(List1.Text)
      PilotNum = Left(PilotNum, 3)
      PilotNum = Trim(PilotNum)
      DeletePilot PilotNum
      UpdateFrame1
    Else
      MsgBox ("No Records Deleted")
    End If
  Else
    MsgBox ("Nothing Selected")
  End If
End Sub

Private Sub Command3_Click()
  PilotEntry.Show
End Sub

Private Sub Command4_Click()
  
    If CurrentContest = 0 Then
      MsgBox ("No Contest Selected")
    Else
      EditContestNumber = CurrentContest
      FromChange = True
      ContestEnter.Show
    End If
End Sub

Private Sub Command5_Click()
Dim ContestNum As String
  ContestNum = Trim(List2.Text)
  ContestNum = Left(ContestNum, 3)
  ContestNum = Trim(ContestNum)
  
  If ContestNum = CurrentContest Then
    MsgBox ("Cannot Change Current Contest Details")
  Else
    If List2.Text <> "" Then
      If MsgBox("You are about do delete this Contest - Are you Sure?", 4) = vbYes Then
        DeleteContest ContestNum
        UpdateFrame2
      Else
        MsgBox ("No Records Deleted")
      End If
    Else
      MsgBox ("Nothing Selected")
    End If
  End If
End Sub

Private Sub Command6_Click()
Dim PilotNum As String
  If List1.Text <> "" Then
    PilotNum = Trim(List1.Text)
    PilotNum = Left(PilotNum, 3)
    PilotNum = Trim(PilotNum)
    SQLString = "SELECT * FROM Pilot WHERE Pilot_ID =" & PilotNum
    Set SelectedSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    If Option1(0) = True Then
      SelectedSet.Edit
      SelectedSet!current = True
      SelectedSet.Update
    ElseIf Option1(1) = True Then
      SelectedSet.Edit
      SelectedSet!current = False
      SelectedSet.Update
    End If
    UpdateFrame1
  Else
    MsgBox ("Nothing Selected")
  End If
End Sub

Private Sub Command7_Click()
  Unload Me
End Sub


Private Sub Command8_Click()
If List2.Text = "" Then
    MsgBox ("Nothing Selected")
  ElseIf CompRoundsDone = False Then
    MsgBox ("The rounds for this contest have not been set up")
  Else
    UnCompleteContest
  End If
End Sub

Private Sub Command9_Click()
Dim ContestNum As String
  ContestNum = Trim(List2.Text)
  ContestNum = Left(ContestNum, 3)
  ContestNum = Trim(ContestNum)
  
  If ContestNum = CurrentContest Then
    MsgBox ("Cannot Change Current Contest Details")
  Else
    If List2.Text <> "" Then
      If MsgBox("You are about do check this Contest - Are you Sure?", 4) = vbYes Then
        CheckContest ContestNum
      Else
        MsgBox ("No Check Done")
      End If
    Else
      MsgBox ("Nothing Selected")
    End If
  End If
End Sub

Private Sub Form_Load()
  
  Option1(1) = True
  Command6.Caption = "Delete Pilot"
  UpdateFrame1
  UpdateFrame2
End Sub


Private Sub DeletePilot(PilotNumber As String)
Dim DeleteSet As Recordset
Dim SQLString As String
  'Delete all the records for this pilot from the
  'scores table
  SQLString = "SELECT * FROM Scores WHERE Pilot_ID =" & Str(PilotNumber)
  Set DeleteSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If DeleteSet.RecordCount > 0 Then
    DeleteSet.MoveFirst
    Do Until DeleteSet.EOF
      DeleteSet.Delete
      DeleteSet.MoveNext
    Loop
  End If
  'Delete all the records for this pilot from the
  'List table
  SQLString = "SELECT * FROM List WHERE Pilot_ID =" & PilotNumber
  Set DeleteSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If DeleteSet.RecordCount > 0 Then
    DeleteSet.MoveFirst
    Do Until DeleteSet.EOF
      DeleteSet.Delete
      DeleteSet.MoveNext
    Loop
  End If
  'Delete all the records for this pilot from the
  'Pilot table
  SQLString = "SELECT * FROM Pilot WHERE Pilot_ID =" & PilotNumber
  Set DeleteSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If DeleteSet.RecordCount > 0 Then
    DeleteSet.MoveFirst
    Do Until DeleteSet.EOF
      DeleteSet.Delete
      DeleteSet.MoveNext
    Loop
    End If
  DeleteSet.Close
End Sub

Private Sub Option1_Click(Index As Integer)
  UpdateFrame1
  If Option1(0) = True Then
    Command6.Caption = "UnDelete Pilot"
  Else
    Command6.Caption = "Delete Pilot"
  End If
End Sub



Public Sub DeleteContest(ContestNumber)
Dim DeleteSet As Recordset
Dim SQLString As String
  'Delete all the records for this pilot from the
  'scores table
  SQLString = "SELECT * FROM Scores WHERE Comp_ID =" & Str(ContestNumber)
  Set DeleteSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If DeleteSet.RecordCount > 0 Then
    DeleteSet.MoveFirst
    Do Until DeleteSet.EOF
      DeleteSet.Delete
      DeleteSet.MoveNext
    Loop
  End If
  'Delete all the records for this pilot from the
  'List table
  SQLString = "SELECT * FROM List WHERE Comp_ID =" & ContestNumber
  Set DeleteSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If DeleteSet.RecordCount > 0 Then
    DeleteSet.MoveFirst
    Do Until DeleteSet.EOF
      DeleteSet.Delete
      DeleteSet.MoveNext
    Loop
  End If
  'Delete all the records for this pilot from the
  'Pilot table
  SQLString = "SELECT * FROM Competition WHERE Comp_ID =" & ContestNumber
  Set DeleteSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If DeleteSet.RecordCount > 0 Then
    DeleteSet.MoveFirst
    Do Until DeleteSet.EOF
      DeleteSet.Delete
      DeleteSet.MoveNext
    Loop
    End If
  DeleteSet.Close
End Sub


Public Sub UnCompleteContest()
Dim UpdateTable As Recordset
Dim SQLString As String
Dim Listset As Recordset
Dim SQLString1 As String
Dim SQLString2 As String
Dim SQLString3 As String
Dim rounds As Integer
Dim I As Integer
Dim Count As Integer
Dim sco() As Single
Dim maxnum As Integer
Dim MinScore As Single
Dim LeaderSet As Recordset
Dim Place As Integer
Dim NumPilots As Integer
Dim ContestNum As String
  
  ContestNum = Trim(List2.Text)
  ContestNum = Left(ContestNum, 3)
  ContestNum = Trim(ContestNum)
  MsgString = "You are about to set this contest as not Complete - are you sure?"
  If MsgBox(MsgString, 1) = vbOK Then
    rounds = NumRounds
    SQLString = "SELECT * FROM Competition WHERE (Competition.Comp_ID = " & ContestNum & ");"
    Set UpdateTable = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    UpdateTable.Edit
    UpdateTable!Completed = False
    UpdateTable.Update
    Form1.Text3.Text = "No"
    SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, List.TempScore,List.FinalCompScore, List.FlyoffScore From List WHERE (List.Comp_ID=" & Str(CurrentContest) & ") ORDER BY List.TempScore DESC;"
    Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
    Listset.MoveFirst
    Do Until Listset.EOF
      Listset.Edit
      Listset!FinalCompScore = 0
      Listset!TempScore = 0
      Listset!FlyoffScore = 0 'Element Flyoffscore stores the pilots place in the contest
      Listset.Update
      Listset.MoveNext
    Loop
    MsgBox ("The Contest is now not completed")
  Else
    MsgBox ("The Contest is Still Completed")
  End If
End Sub

Public Sub CheckContest(ContestNumber)
Dim CheckSet As Recordset
Dim ThisPilotSet As Recordset
Dim SQLString As String
Dim ThisPilot As String
Dim LastRound As Integer
Dim NumberOfDupes As Integer
  NumberOfDupes = 0
  SQLString = "SELECT * FROM List WHERE Comp_ID =" & ContestNumber
  Set CheckSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If CheckSet.RecordCount > 0 Then
    CheckSet.MoveFirst
    Do Until CheckSet.EOF
      ThisPilot = Str(CheckSet!Pilot_ID)
      'SQLString = "SELECT * FROM Scores WHERE (Comp_ID =" & ContestNumber & " AND Pilot_ID = " & ThisPilot & ") ORDER BY Scores.Pilot_ID, Scores.Task, Scores.Round ASC"
      SQLString = "SELECT Distinctrow Scores.Pilot_ID, Scores.Comp_ID, Scores.Task, Scores.Round From Scores Where (((Scores.Pilot_ID) = " & ThisPilot & ") And ((Scores.Comp_ID) = " & Str(ContestNumber) & ")) ORDER BY Scores.Task, Scores.Round;"
      Set ThisPilotSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
      ThisPilotSet.MoveFirst
      LastRound = ThisPilotSet!Round
      ThisPilotSet.MoveNext
      Do Until ThisPilotSet.EOF
        If ThisPilotSet!Round = LastRound Then
          'duplicate
          NumberOfDupes = NumberOfDupes + 1
          LastRound = ThisPilotSet!Round
          ThisPilotSet.Edit
          ThisPilotSet.Delete
          ThisPilotSet.MoveNext
        Else
          LastRound = ThisPilotSet!Round
          ThisPilotSet.MoveNext
        End If
      Loop
    CheckSet.MoveNext
    Loop
    MsgBox ("There were " & Str(NumberOfDupes) & " Duplicates Found")
  Else
    MsgBox ("Contest had Competitors entered")
  End If
End Sub
