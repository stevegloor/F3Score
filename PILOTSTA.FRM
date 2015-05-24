VERSION 5.00
Begin VB.Form PilotStats 
   Caption         =   "Pilot Statistics"
   ClientHeight    =   4530
   ClientLeft      =   765
   ClientTop       =   1200
   ClientWidth     =   6630
   LinkTopic       =   "Form4"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4530
   ScaleWidth      =   6630
   Begin VB.Frame Frame2 
      Caption         =   "Pilot Statistics"
      Height          =   2895
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   3015
      Begin VB.ListBox List2 
         Height          =   2400
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pilots in Championship"
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3375
      Begin VB.ListBox List1 
         Height          =   3765
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Done"
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print Pilots Stats"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Pilots Stats"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   3120
      Width           =   1335
   End
End
Attribute VB_Name = "PilotStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelectedChampScores As Recordset
Public Sub UpdateList1()

Dim SQLString As String
Dim Currentpilot As Integer
Dim Line As String


Currentpilot = 0
If CurrentChamp <> 0 Then
  'output list of pilots
 SQLString = "SELECT DISTINCTROW Competition.ChampID, List.FlyoffScore, Competition.Completed, Pilot.Pilot_ID, List.FinalCompScore, Pilot.First_Name, Pilot.Last_Name FROM Pilot INNER JOIN ((Championship INNER JOIN Competition ON Championship.ChampID = Competition.ChampID) INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) ON Pilot.Pilot_ID = List.Pilot_ID Where ((Competition.ChampID = " & Trim(Str(CurrentChamp)) & ") And (Competition.Completed = True)) ORDER BY Pilot.Pilot_ID;"
  Set SelectedChampScores = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If SelectedChampScores.RecordCount <> 0 Then
    SelectedChampScores.MoveFirst
    Do Until SelectedChampScores.EOF
      If SelectedChampScores!Pilot_ID <> Currentpilot Then
        Line = Format(Str(SelectedChampScores!Pilot_ID), "!@@@@") & "  " & SelectedChampScores!First_Name & " " & SelectedChampScores!Last_Name
        List1.AddItem Line
        Currentpilot = SelectedChampScores!Pilot_ID
        SelectedChampScores.MoveNext
      Else
        SelectedChampScores.MoveNext
      End If
    Loop
  End If
End If
End Sub

Public Sub ViewPilotStats()
Dim StatSet As Recordset
Dim SQLString As String
Dim Pilot As String
Dim MaxFlight As Integer
Dim Mins As Integer
Dim Secs As Single
Dim Line As String
Dim Search As String
  
  Pilot = Trim(List1.Text)
  Pilot = Left(Pilot, 3)
  Pilot = Trim(Pilot)
  List2.Clear
  SelectedChampScores.MoveFirst
  Search = "Pilot_ID = " & Pilot
  SelectedChampScores.FindFirst Search
  Line = "Statistics for " & SelectedChampScores!First_Name & " " & SelectedChampScores!Last_Name
  List2.AddItem Line
  Line = ""
  List2.AddItem Line
  'Task A Stats
  SQLString = "SELECT MAX(Flight_Time) AS [Res2] FROM (Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID) WHERE ((List.Pilot_ID= " & Pilot & ") AND (Competition.ChampID= " & Trim(Str(CurrentChamp)) & ") AND (Scores.Task='A'));"
  Set StatSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  Mins = StatSet!Res2 / 60
  Secs = StatSet!Res2 Mod 60
  Line = "Longest Duration  -  " & Format(Mins, "0#") & ":" & Format(Secs, "0#")
  List2.AddItem Line
  SQLString = "SELECT MIN(Flight_Time) AS [Res2] FROM (Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID) WHERE ((List.Pilot_ID= " & Pilot & ") AND (Competition.ChampID= " & Trim(Str(CurrentChamp)) & ") AND (Scores.Task='A'));"
  Set StatSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  Mins = StatSet!Res2 / 60
  Secs = StatSet!Res2 Mod 60
  Line = "Shortest Duration  -  " & Format(Mins, "0#") & ":" & Format(Secs, "0#")
  List2.AddItem Line
  SQLString = "SELECT AVG(Flight_Time) AS [Res2] FROM (Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID) WHERE ((List.Pilot_ID= " & Pilot & ") AND (Competition.ChampID= " & Trim(Str(CurrentChamp)) & ") AND (Scores.Task='A'));"
  Set StatSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  Mins = StatSet!Res2 / 60
  Secs = StatSet!Res2 Mod 60
  Line = "Average Duration  -  " & Format(Mins, "0#") & ":" & Format(Secs, "0#")
  List2.AddItem Line
  If SelectedContestType = "F3B" Then
    'Task B stats
    SQLString = "SELECT MAX(Flight_Time) AS [Res2] FROM (Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID) WHERE ((List.Pilot_ID= " & Pilot & ") AND (Competition.ChampID= " & Trim(Str(CurrentChamp)) & ") AND (Scores.Task='B'));"
    Set StatSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    Line = "Most Laps -  " & Str(StatSet!Res2)
    List2.AddItem Line
    SQLString = "SELECT MIN(Flight_Time) AS [Res2] FROM (Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID) WHERE ((List.Pilot_ID= " & Pilot & ") AND (Competition.ChampID= " & Trim(Str(CurrentChamp)) & ") AND (Scores.Task='B'));"
    Set StatSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    Line = "Least Laps  -  " & Str(StatSet!Res2)
    List2.AddItem Line
    SQLString = "SELECT AVG(Flight_Time) AS [Res2] FROM (Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID) WHERE ((List.Pilot_ID= " & Pilot & ") AND (Competition.ChampID= " & Trim(Str(CurrentChamp)) & ") AND (Scores.Task='B'));"
    Set StatSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    Line = "Average Laps  -  " & Str(StatSet!Res2)
    List2.AddItem Line
    'Task C Stats
    SQLString = "SELECT MIN(Flight_Time) AS [Res2] FROM (Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID) WHERE ((List.Pilot_ID= " & Pilot & ") AND (Competition.ChampID= " & Trim(Str(CurrentChamp)) & ") AND (Scores.Task='C'));"
    Set StatSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    Mins = StatSet!Res2 / 100
    Secs = StatSet!Res2 Mod 100
    Line = "Fastest Time  -  " & Format(Mins, "0#") & ":" & Format(Secs, "0#")
    List2.AddItem Line
    SQLString = "SELECT MAX(Flight_Time) AS [Res2] FROM (Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID) WHERE ((List.Pilot_ID= " & Pilot & ") AND (Competition.ChampID= " & Trim(Str(CurrentChamp)) & ") AND (Scores.Task='C'));"
    Set StatSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    Mins = StatSet!Res2 / 100
    Secs = StatSet!Res2 Mod 100
    Line = "Slowest Time  -  " & Format(Mins, "0#") & ":" & Format(Secs, "0#")
    List2.AddItem Line
    SQLString = "SELECT AVG(Flight_Time) AS [Res2] FROM (Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID) WHERE ((List.Pilot_ID= " & Pilot & ") AND (Competition.ChampID= " & Trim(Str(CurrentChamp)) & ") AND (Scores.Task='C'));"
    Set StatSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    Mins = StatSet!Res2 / 100
    Secs = StatSet!Res2 Mod 100
    Line = "Average Time  -  " & Format(Mins, "0#") & ":" & Format(Secs, "0#")
    List2.AddItem Line
  End If
  SQLString = "SELECT COUNT(Flight_Time) AS [Res2] FROM (Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID) WHERE ((List.Pilot_ID= " & Pilot & ") AND (Competition.ChampID= " & Trim(Str(CurrentChamp)) & ") AND (Scores.Task='A'));"
  Set StatSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  Line = "From " & Str(StatSet!Res2) & " Flights"
  List2.AddItem Line
  SQLString = "SELECT MAX(Score) AS [Res2] FROM (Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID) WHERE ((List.Pilot_ID= " & Pilot & ") AND (Competition.ChampID= " & Trim(Str(CurrentChamp)) & "));"
  Set StatSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  Line = "Best Score - " & Str(StatSet!Res2)
  List2.AddItem Line
  SQLString = "SELECT MIN(Score) AS [Res2] FROM (Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID) WHERE ((List.Pilot_ID= " & Pilot & ") AND (Competition.ChampID= " & Trim(Str(CurrentChamp)) & "));"
  Set StatSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  Line = "Worst Score - " & Str(StatSet!Res2)
  List2.AddItem Line
  SQLString = "SELECT COUNT(Score) AS [Res2] FROM (Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID) WHERE ((List.Pilot_ID= " & Pilot & ") AND (Competition.ChampID= " & Trim(Str(CurrentChamp)) & ") AND (Scores.Score = 1000));"
  Set StatSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  Line = "Number of 1000 scores - " & Str(StatSet!Res2)
  List2.AddItem Line
 
 ' SQLString = "SELECT MAX(Flight_Time) AS [Res2] FROM (Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID) WHERE ((List.Pilot_ID= " & Pilot & ") AND (Competition.ChampID= " & Trim(Str(CurrentChamp)) & ") AND (Scores.Task='A'));"
  
  
  
  
  
  'SQLString = "SELECT DISTINCTROW List.Pilot_ID, Competition.ChampID, Competition.CompStartDate, Competition.Comp_Name, List.Comp_ID, List.FinalCompScore, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.LandBonus, Scores.Penalty, Scores.Score, Scores.Task, Scores.RawScore FROM (Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID) WHERE ((List.Pilot_ID= " & Pilot & ") AND (Competition.ChampID= " & Trim(Str(CurrentChamp)) & "));"

  
  
End Sub

Public Sub PrintPilotStats()

End Sub

Private Sub Command1_Click()
  ViewPilotStats
End Sub

Private Sub Command3_Click()
  Unload Me
End Sub


Private Sub Form_Load()
  UpdateList1
End Sub


