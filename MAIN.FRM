VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{625E24A3-B09D-101D-85F5-6EBA1EE93AF4}#3.3#0"; "IniCon32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "F3 Scoring Program"
   ClientHeight    =   7650
   ClientLeft      =   255
   ClientTop       =   795
   ClientWidth     =   8655
   FillColor       =   &H00C0C000&
   ForeColor       =   &H00000000&
   HelpContextID   =   470
   Icon            =   "MAIN.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7650
   ScaleWidth      =   8655
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8040
      TabIndex        =   8
      Top             =   570
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin IniconLib.Init Init1 
      Left            =   2400
      Top             =   6240
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   7080
      WhatsThisHelpID =   10
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   7080
      WhatsThisHelpID =   10
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Display"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   7080
      WhatsThisHelpID =   10
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5895
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   10398
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      Caption         =   "Number of Rounds"
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "COMPLETE"
      Height          =   255
      Left            =   6960
      TabIndex        =   6
      Top             =   240
      WhatsThisHelpID =   10
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Contest "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      WhatsThisHelpID =   10
      Width           =   1095
   End
   Begin VB.Menu File 
      Caption         =   "File"
      HelpContextID   =   210
      Begin VB.Menu Space 
         Caption         =   "-"
      End
      Begin VB.Menu SetupPrint 
         Caption         =   "Print Setup"
      End
      Begin VB.Menu PrintPrinter 
         Caption         =   "Print"
      End
      Begin VB.Menu ExportCsv 
         Caption         =   "Export to CSV"
      End
      Begin VB.Menu FileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Comp 
      Caption         =   "Competition"
      HelpContextID   =   220
      Begin VB.Menu CompNew 
         Caption         =   "New Competition"
      End
      Begin VB.Menu CompOpen 
         Caption         =   "Open Competition"
      End
      Begin VB.Menu Space1 
         Caption         =   "-"
      End
      Begin VB.Menu CompEnter 
         Caption         =   "Enter Competitors"
      End
      Begin VB.Menu SetTeams 
         Caption         =   "Set Up Teams"
      End
      Begin VB.Menu Space2 
         Caption         =   "-"
      End
      Begin VB.Menu FreqManage 
         Caption         =   "Manage Freqs"
      End
      Begin VB.Menu RoundAlt 
         Caption         =   "Set Up Rounds"
      End
      Begin VB.Menu MovePilot 
         Caption         =   "Move Pilot"
      End
      Begin VB.Menu Space10 
         Caption         =   "-"
      End
      Begin VB.Menu LeadBoard 
         Caption         =   "Leader Board"
      End
   End
   Begin VB.Menu F3B 
      Caption         =   "F3B"
      Begin VB.Menu F3Ballloc 
         Caption         =   "F3B Allocation Module"
      End
      Begin VB.Menu SetFlightGoups 
         Caption         =   "Set Up Task B Flight Groups"
      End
      Begin VB.Menu ViewGroups 
         Caption         =   "View Task B Flight Groups"
      End
   End
   Begin VB.Menu View 
      Caption         =   "View"
      HelpContextID   =   230
      Begin VB.Menu ViewPilots 
         Caption         =   "Pilots in Contest"
      End
      Begin VB.Menu ViewAlloc 
         Caption         =   "Contest Allocation"
      End
      Begin VB.Menu ViewTeams 
         Caption         =   "Teams"
      End
   End
   Begin VB.Menu Score 
      Caption         =   "Scores"
      HelpContextID   =   240
      Begin VB.Menu EnterSlot 
         Caption         =   "Enter Slot Flight Times"
      End
      Begin VB.Menu Space9 
         Caption         =   "-"
      End
      Begin VB.Menu VScores 
         Caption         =   "View Scores"
      End
   End
   Begin VB.Menu Competit 
      Caption         =   "DataBase"
      HelpContextID   =   250
      Begin VB.Menu PilotDelete 
         Caption         =   "Edit User/Contests"
      End
      Begin VB.Menu AddNewUser 
         Caption         =   "Add New User to Database"
      End
      Begin VB.Menu SearchDatabase 
         Caption         =   "Search Database"
      End
      Begin VB.Menu hello 
         Caption         =   "-"
      End
      Begin VB.Menu DatabaseBackup 
         Caption         =   "Backup Database"
      End
      Begin VB.Menu DatabaseRepair 
         Caption         =   "Repair Database"
      End
   End
   Begin VB.Menu Tool 
      Caption         =   "Tools"
      HelpContextID   =   260
      Begin VB.Menu SetChange 
         Caption         =   "Change User Settings"
      End
      Begin VB.Menu hello1 
         Caption         =   "-"
      End
      Begin VB.Menu CompContest 
         Caption         =   "Set Current Contest Complete"
      End
      Begin VB.Menu ChangeContest 
         Caption         =   "Change Current Contest Details"
      End
      Begin VB.Menu ConReset 
         Caption         =   "Reset Current Contest"
      End
      Begin VB.Menu hello3 
         Caption         =   "-"
      End
      Begin VB.Menu SetupChampionship 
         Caption         =   "Championship Maintenance"
      End
      Begin VB.Menu FreqLook 
         Caption         =   "Frequency Lookup"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      HelpContextID   =   270
      Begin VB.Menu HlpSearch 
         Caption         =   "Help"
      End
      Begin VB.Menu F3Brules 
         Caption         =   "F3B/F3J Rules"
      End
      Begin VB.Menu AustOpenR 
         Caption         =   "Australian Open Rules"
      End
      Begin VB.Menu MenuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_NORMAL = 1
Public Sub GroupView()
Dim PilotSet As Recordset
Dim Selected As Recordset
Dim SQLString1 As String
Dim NumPilots As Integer
Dim CurrentGroup As Integer
Dim Line As String
Dim RoundNum
Dim Slot
Dim Search As String

  Rem On Error GoTo Errorhandler
  GeneralDisplay.List1.Clear
  CurrentTeam = 0
  'SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & "));"
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID,List.FreqUsed, List.Team_ID, List.InTeam, Pilot.First_Name, Pilot.Last_Name, Pilot.Freq1, Pilot.Freq2, Pilot.Club, Pilot.Number FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & "));"
  Set PilotSet = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  
  PilotSet.MoveFirst
  GeneralDisplay.Caption = "List of Flight Groups in Contest"
  For RoundNum = 1 To NumRounds
    GeneralDisplay.List1.AddItem ""
    Line = "Round" & RoundNum
    GeneralDisplay.List1.AddItem Line
    For Slot = 1 To NumSlots
      GeneralDisplay.List1.AddItem ""
      Line = "Slot" & Slot
      GeneralDisplay.List1.AddItem Line
      Line = ""
      CurrentGroup = 0
      SQLString1 = "SELECT * From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Round= " & Trim(Str(RoundNum)) & ") AND (Scores.Slot=" & Trim(Str(Slot)) & ") AND (Scores.Task = 'B'))ORDER By Scores.FlightGroup;"
      Set Selected = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
      Selected.MoveFirst
      Do Until Selected.EOF
        If CurrentGroup <> Selected!FlightGroup Then
          If Line <> "" Then
            GeneralDisplay.List1.AddItem Line
          End If
          CurrentGroup = Selected!FlightGroup
          Line = ""
          Line = Line & "Flight Group" & Str(CurrentGroup) & "  "
        End If
        Search = "Pilot_ID = " & Str(Selected!Pilot_ID)
        PilotSet.FindFirst Search
        Line = Line & "  " & Trim(Left(PilotSet!First_Name, 1)) & ". " & Trim(PilotSet!Last_Name)
        PilotSet.MoveFirst
        Selected.MoveNext
      Loop
      GeneralDisplay.List1.AddItem Line
    Next Slot
  Next RoundNum
  
  GeneralDisplay.Show
Exit Sub
Errorhandler:
  MsgBox ("No Teams Entered")
End Sub


Private Sub ChampSetUp_Click()
  ChampMaint.Show
End Sub


Private Sub AddNewUser_Click()
PilotEntry.Show
 
End Sub

Private Sub AustOpenR_Click()
Dim strUrl As String
  'uses ShellExecute API - Function defined in declerations
  On Error GoTo errhandle
  
  strUrl = App.Path + "\AustralianOpenThermalRules.pdf"
  ShellExecute Me.hWnd, "open", strUrl, vbNullString, vbNullString, SW_NORMAL
  'MsgBox ("The rule file is in the application directory" + Chr$(13) + "It is called f3jbrules.pdf")
  Exit Sub
errhandle:
  MsgBox ("You do not have Acrobat Reader installed")
End Sub

Private Sub ChangeContest_Click()
  If CurrentContest = 0 Then
     MsgBox ("No Contest Selected")
  Else
    EditContestNumber = CurrentContest
    FromChange = True
    ContestEnter.Show
  End If
End Sub

Private Sub Command2_Click()
  Form3.Show
  
End Sub


Private Sub Command3_Click()
  F3JDb.Close
  Unload Me
  End
End Sub


Private Sub CompContest_Click()
  
  If CurrentContest = 0 Then
    MsgBox ("No Contest Selected")
  ElseIf CompRoundsDone = False Then
    MsgBox ("The rounds for this contest have not been set up")
  Else
    CompleteContest
  End If
End Sub

Private Sub CompetitView_Click()
  MsgBox ("Not Yet Supported")
End Sub

Private Sub ConDelete_Click()
  MsgBox ("Not yet supported")
End Sub

Private Sub ConReset_Click()
  If CurrentContest <> 0 Then
    ContestReset
  Else
    MsgBox ("No Contest Selected")
  End If
  
End Sub

Private Sub EditExist_Click()
  FromEdit = True
  PilotAdd.Show
End Sub


Private Sub DatabaseCompact_Click()
Dim datafile As String
Dim Backupfile As String

  On Error GoTo Errorhandler
  
  datafile = App.Path & "\DATA.MDB"
  Backupfile = App.Path & "\DATABAK.MDB"
 

  F3JDb.Close
  DBEngine.CompactDatabase datafile, Backupfile

  Set F3JDb = DBEngine.Workspaces(0).OpenDatabase(datafile)
  CurrentContest = 0
  FromChange = False
  MsgBox ("DataBase Backed Up")
  Exit Sub

Errorhandler:
  MsgBox ("Can't backup database")
End Sub

Private Sub DatabaseBackup_Click()
Dim datafile As String
Dim Backupfile As String
Dim Exists As String

  Rem On Error GoTo Errorhandler
  
  datafile = App.Path & "\DATA.MDB"
  Backupfile = App.Path & "\DATABAK.MDB"
  Exists = Dir(Backupfile)
  If Exists <> "" Then
    Kill Backupfile
  End If
  F3JDb.Close
  DBEngine.CompactDatabase datafile, Backupfile

  Set F3JDb = DBEngine.Workspaces(0).OpenDatabase(datafile)
  CurrentContest = 0
  FromChange = False
  MsgBox ("DataBase Backed Up")
  Exit Sub

Errorhandler:
  MsgBox ("Can't backup database")
End Sub

Private Sub DatabaseRepair_Click()
Dim datafile As String
  On Error GoTo Errorhandler
  
  datafile = App.Path & "\DATA.MDB"
  If MsgBox("Repairing a database that is not corrupt will corrupt it" & Chr(13) & "Do you want to continue", 19) <> vbYes Then
    MsgBox ("Repair Cancelled")
    Exit Sub
  Else
    F3JDb.Close
    DBEngine.RepairDatabase datafile
    Set F3JDb = DBEngine.Workspaces(0).OpenDatabase(datafile)
    CurrentContest = 0
    FromChange = False
    MsgBox ("Database Repaired")
    Exit Sub
  End If
Errorhandler:
  MsgBox ("Can't repair database")
  
End Sub



Private Sub ExportCsv_Click()
  If CurrentContest = 0 Then
    For I = 1 To 4
      ExportForm.Option1(I).Enabled = False
    Next I
    ExportForm.Show
  ElseIf CompRoundsDone = False Then
    For I = 1 To 4
      ExportForm.Option1(I).Enabled = False
    Next I
    ExportForm.Show
  Else
    For I = 1 To 4
      ExportForm.Option1(I).Enabled = True
    Next I
    ExportForm.Show
  End If
End Sub

Private Sub F3Ballloc_Click()
  If CurrentContest = 0 Then
    MsgBox ("No Contest Selected")
  ElseIf CompRoundsDone = True Then
    MsgBox ("This contest has the rounds set up")
  Else
    AllocateF3B.Show
  End If
  
End Sub

Private Sub F3Brules_Click()
  
  Dim strUrl As String
  'uses ShellExecute API - Function defined in declerations
  On Error GoTo errhandle
  
  strUrl = App.Path + "\f3bjrules.pdf"
  ShellExecute Me.hWnd, "open", strUrl, vbNullString, vbNullString, SW_NORMAL
  'MsgBox ("The rule file is in the application directory" + Chr$(13) + "It is called f3jbrules.pdf")
  Exit Sub
errhandle:
  MsgBox ("You do not have Acrobat Reader installed")
End Sub

Private Sub F3JRules_Click()
Dim RuleFile As String
Dim ApplicationPath As String

Dim ReturnVal As Integer
  ApplicationPath = Chr$(34)
  ApplicationPath = ApplicationPath + App.Path
  ApplicationPath = ApplicationPath + "\f3jrules.wri" + Chr$(34)
  RuleFile = "write " & ApplicationPath
  'MsgBox (RuleFile)
  ReturnVal = Shell(RuleFile, 1)

End Sub

Private Sub FileOpen_Click()
'  Form2.Show
' Cancel Error is True
On Error GoTo errhandler
' Set Filters
CommonDialog1.Filter = "Competition DataBases (*.mdb)|*.MDB"
CommonDialog1.FilterIndex = 2

CommonDialog1.ShowOpen
datafile = CommonDialog1.FileName
Set F3JDb = DBEngine.Workspaces(0).OpenDatabase(datafile)

  CurrentContest = 0
  FromChange = False
errhandler:
  Exit Sub
End Sub


Private Sub Command1_Click()
 
 Cleargrid
End Sub

Private Sub CompEnter_Click()
  If CurrentContest = 0 Then
    MsgBox ("No Contest Selected")
  ElseIf CompRoundsDone = True Then
    PilotSelect.Command1.Enabled = False
    PilotSelect.Command1.Visible = False
    PilotSelect.Command2.Visible = False
    PilotSelect.Command2.Enabled = False
    PilotSelect.Show
  Else
    PilotSelect.Show
  End If
End Sub

Private Sub CompetitEdit_Click()
  PilotEntry.Show
  
End Sub

Private Sub CompNew_Click()
  ContestEnter.Show
End Sub

Private Sub CompOpen_Click()
  Dim MyTable As Recordset
  Set MyTable = F3JDb.OpenRecordset("Competition", dbOpenTable)
  If MyTable.RecordCount = 0 Then
    MsgBox ("There are no Contests in the Database")
    ContestEnter.Show
  Else
    ContestSelect.Show
  End If
  MyTable.Close
End Sub

Private Sub CompSetUp_Click()
  If CurrentContest = 0 Then
    MsgBox ("No Contest Selected")
  ElseIf CompRoundsDone = True Then
    MsgBox ("This contest has the rounds set up")
  Else
    Allocate.Show
  End If
End Sub

Private Sub EnterPilotFlightTime_Click()

End Sub

Private Sub EnterSlot_Click()
On Error GoTo Errorhandler
If CurrentContest = 0 Then
  MsgBox ("No Contest Selected")
ElseIf CompRoundsDone = False Then
  MsgBox ("The rounds for this contest have not been set up")
ElseIf Completed = True Then
  MsgBox ("This contest is complete")
Else
  FromView = False
  If CurrentContestType = "F3B" Then
    EnterF3BScores.Show
  ElseIf CurrentContestType = "Club" Then
    EnterClubScores.Show
  Else
    EnterF3JScores.Show
  End If
End If

Errorhandler:
  Exit Sub
End Sub


Private Sub FileExit_Click()
  F3JDb.Close
  Unload Me
  End
End Sub

Private Sub FileSave_Click()
'  Form2.Show
' Cancel Error is True
On Error GoTo errhandler
' Set Filters
CommonDialog1.Filter = "Competition DataBases (*.mdb)|*.MDB"
CommonDialog1.FilterIndex = 2

CommonDialog1.ShowSave
MsgBox (CommonDialog1.FileName)
errhandler:
  Exit Sub
End Sub



Private Sub Form_Load()
'*** Code added by VB HelpWriter ***
    SetAppHelp Me.hWnd
'***********************************

Dim GridHgt
Dim GridWid
Dim strDatabase As String
  blLocal = True
  LeaderboardDone = False
'On Error GoTo errhandler
  ReDim Place(20)
  FromEdit = False
  Init1.FileName = App.Path & "\f3score.ini"
  Init1.Application = "Database"
  Init1.Parameter = "Location"
  If Init1.Value = "Remote" Then
    blLocal = False
  End If
  'set database location
  If blLocal = True Then
    'database is local
    'set the global variable for the database location
    strDatabaseName = App.Path & "\Data.MDB"
    DatabaseType = "Jet"
  Else
    'database is remote
    Init1.Application = "Database"
    Init1.Parameter = "Type"
    If Init1.Value = "Access" Then 'for an access database on a share
      DatabaseType = "Jet"
      Init1.Parameter = "Path"
      'set the global variable for the database location
      strDatabaseName = Init1.Value
      'use local database class as connection is still access
    End If
  End If
  
  
  
  'If CommonDialog1.FileName <> "" Then
  '  datafile = CommonDialog1.FileName
  'Else
  '
  'End If
  'datafile = App.Path & "\DATA.MDB"
  Set F3JDb = DBEngine.Workspaces(0).OpenDatabase(strDatabaseName)
  
  Init1.Application = "Program"
  Init1.Parameter = "RoundDiscarded"
  DiscardRound = Val(Init1.Value)
  Init1.Parameter = "TenK"
  TenK = Val(Init1.Value)
  Init1.Parameter = "Country"
  Country = Val(Init1.Value)
  Init1.Application = "Max"
  Init1.Parameter = "F3J"
  maxF3J = Val(Init1.Value)
  Init1.Parameter = "F3JFO"
  maxF3JFO = Val(Init1.Value)
  Init1.Parameter = "F3B"
  maxF3B = Val(Init1.Value)
  Init1.Parameter = "AustOpen"
  max811 = Val(Init1.Value)
  Init1.Parameter = "Club"
  maxClub = Val(Init1.Value)
  Init1.Application = "Club"
  Init1.Parameter = "Time"
  ClubTime = Val(Init1.Value)
  Init1.Parameter = "Flights"
  ClubFlights = Val(Init1.Value)
  Init1.Application = "Championship"
  Init1.Parameter = "Cummlative"
  If Init1.Value = 1 Then
    Cummlative = True
  Else
    Cummlative = False
  End If
  For I = 1 To 20
    PlaceNum = "Place" & Trim(Str(I))
    Init1.Parameter = PlaceNum
    Place(I) = Val(Init1.Value)
  Next I
  CurrentContest = 0
  FromChange = False
  GridHgt = MSFlexGrid1.Height - MSFlexGrid1.RowHeight(0) - 30
  GridWid = MSFlexGrid1.Width - MSFlexGrid1.Colwidth(0) - 30
  MSFlexGrid1.Rows = 11
  MSFlexGrid1.Cols = 6
  For I = 1 To 10
    MSFlexGrid1.RowHeight(I) = GridHgt / 10
  Next I
  For I = 1 To 5
    MSFlexGrid1.Colwidth(I) = GridWid / 5
  Next I
  MSFlexGrid1.ScrollBars = 0
  F3B.Visible = False
  FlightGroups = False
  ViewScores = False
Exit Sub
errhandler:
  MsgBox ("Cannot Open Database File" & Chr(13) & "Run Repair Database")
  
  
  
End Sub












Private Sub ScoreViewCompe_Click()

End Sub


Private Sub ScoreViewRound_Click()

End Sub


Private Sub ScoreEnterRound_Click()

End Sub


Private Sub ScoreSlot_Click()

End Sub






Private Sub FreqChange_Click()
  If CurrentContest <> 0 Then
    Form2.Show
  Else
    MsgBox ("No Contest has been Selected")
  End If
End Sub

Private Sub Form_Resize()
If ScaleHeight > 10 Then
  MSFlexGrid1.Top = 960
  MSFlexGrid1.Left = 120
  MSFlexGrid1.Height = ScaleHeight - 2000
  MSFlexGrid1.Width = ScaleWidth - 240
  Command3.Top = ScaleHeight - 600
  Command1.Left = ScaleWidth - (ScaleWidth - 4000)
  Command3.Width = 2175
  Command2.Top = ScaleHeight - 600
  Command1.Left = ScaleWidth - (ScaleWidth - 2000)
  Command3.Width = 2175
  Command1.Top = ScaleHeight - 600
  Command1.Left = ScaleWidth - (ScaleWidth - 120)
  Command3.Width = 2175
End If
If LeaderboardDone = False Then
  Cleargrid
End If
End Sub

Private Sub FreqLook_Click()
  On Error GoTo errhandler
  FreqDisp.Show
  Exit Sub
errhandler:
  
End Sub

Private Sub FreqManage_Click()
  If CurrentContest = 0 Then
    MsgBox ("No Contest Selected")
  
  Else
    'FrequencyManage
    Freq.Show
  End If
End Sub

Private Sub msflexgrid1_SelChange()
'    Msg = "Selection: " & Chr(64 + MSFlexGrid1.SelStartCol)
 '   Msg = Msg & MSFlexGrid1.SelStartRow
'    Msg = Msg & ":" & Chr(64 + MSFlexGrid1.SelEndCol)
'    Msg = Msg & MSFlexGrid1.SelEndRow
'    Label2.Caption = Msg
End Sub


Private Sub HelpContents_Click()

End Sub

Private Sub HelpSearch_Click()

End Sub

Private Sub HlpSearch_Click()
  ShowHelpTopic (10)
End Sub

Private Sub MenuHelpAbout_Click()
  HelpAbout.Show
End Sub

Private Sub HelpAbout_Click()

End Sub


Private Sub HelpManual_Click()
  Dim RuleFile As String
  Dim ReturnVal As Integer
  
    RuleFile = "write " & App.Path & "\readme.wri"
    ReturnVal = Shell(RuleFile, 1)

End Sub

Private Sub LeadBoard_Click()
  If CurrentContest = 0 Then
    MsgBox ("No Contest Selected")
  ElseIf CompRoundsDone = False Then
    MsgBox ("The rounds for this contest have not been set up")
  Else
    Leaderboard
  End If
End Sub

Private Sub MovePilot_Click()
  If CurrentContest = 0 Then
    MsgBox ("No Contest Selected")
  
  ElseIf CurrentContestType = "F3B" Then
    'MsgBox ("You cannot move pilots in F3B contests")
    F3BPilotMaint.Show
  
  ElseIf CompRoundsDone = True Then
    F3JPilotMaint.Show
  
  Else
    MsgBox ("This contest has not has the rounds set up")
  
  End If
End Sub

Private Sub PilotDelete_Click()
  DataMaint.Show
End Sub

Private Sub PrintPrinter_Click()
  Dim I As Integer
  
  If CurrentContest = 0 Then
    For I = 0 To 3
      Form3.Option1(I).Enabled = False
    Next I
    Form3.Option1(4) = True
    Form3.Show
  ElseIf CompRoundsDone = False Then
    Form3.Option1(0).Enabled = False
    Form3.Option1(1).Enabled = False
    For I = 2 To 4
      Form3.Option1(I).Enabled = True
    Next I
    Form3.Show
  Else
    For I = 0 To 4
      Form3.Option1(I).Enabled = True
    Next I
    Form3.Show
  End If
    
End Sub

Private Sub RoundAlt_Click()
If CurrentContest = 0 Then
    MsgBox ("No Contest Selected")
  ElseIf CompRoundsDone = True Then
    MsgBox ("This contest has the rounds set up")
  Else
    Allocate2.Show
  End If
End Sub

Private Sub SearchDatabase_Click()
  PilotAdd.Show
End Sub

Private Sub SetChange_Click()
  EditSettings.Show
End Sub

Private Sub SetFlightGoups_Click()
If CurrentContest = 0 Then
    MsgBox ("No Contest Selected")
  
  ElseIf CompRoundsDone = True Then
    GroupMaint.Show
  
  Else
    MsgBox ("This contest has not had the rounds set up")
  
  End If
End Sub

Private Sub SetTeams_Click()
  If CurrentContest = 0 Then
    MsgBox ("No Contest Selected")
  ElseIf CompRoundsDone = True Then
    TeamSelect.Command1.Enabled = False
    TeamSelect.Command1.Visible = False
    TeamSelect.Command1.Enabled = False
    TeamSelect.Command2.Visible = False
    TeamSelect.Show
  Else
    TeamSelect.Show
  End If
End Sub

Private Sub SetupChampionship_Click()
  ChampMaint.Show
  
End Sub

Private Sub SetupPrint_Click()
  CommonDialog1.ShowPrinter
End Sub



Private Sub ViewAlloc_Click()
If CurrentContest = 0 Then
    MsgBox ("No Contest Selected")
  
  ElseIf CompRoundsDone = True Then
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
      ContestView
    Else
      Currenttask = "A"
      ContestView
    End If
  Else
    MsgBox ("This contest has not had the rounds set up")
  
  End If
End Sub

Private Sub ViewContest_Click()
  
  
  
End Sub

Private Sub ViewGroups_Click()
If CurrentContest = 0 Then
    MsgBox ("No Contest Selected")
  
  ElseIf CompRoundsDone = True Then
    GroupView
  
  Else
    MsgBox ("This contest has not had the rounds set up")
  
  End If
End Sub

Private Sub ViewPilots_Click()
  If CurrentContest = 0 Then
    MsgBox ("No Contest Selected")
  Else
    PilotsView
  End If
End Sub

Private Sub ViewTeams_Click()
If CurrentContest = 0 Then
    MsgBox ("No Contest Selected")
  
  ElseIf CompRoundsDone = True Then
    TeamsView
  
  Else
    MsgBox ("This contest has not had the rounds set up")
  
  End If
End Sub

Private Sub VScores_Click()
  On Error GoTo Errorhandler
  If CurrentContest = 0 Then
    MsgBox ("No Contest Selected")
  ElseIf CompRoundsDone = False Then
    MsgBox ("The rounds for this contest have not been set up")
  Else
    If CurrentContestType = "F3B" Then
      
      ViewF3BScores.Show
    Else
      ViewScores = True
      EnterF3JScores.Show
    End If
    
  End If
  Exit Sub
Errorhandler:
  MsgBox ("An error has occured")
End Sub



Public Sub LBSetGrid(ThisGrid As Control, rounds, NumPilots)
  Dim Col2Fill
  Dim Row2Fill
  Dim NL As String
  Dim GridHgt
  Dim GridWid
  
  GridHgt = MSFlexGrid1.Height
  GridWid = MSFlexGrid1.Width
  Cleargrid
  MSFlexGrid1.Cols = rounds + 3
  MSFlexGrid1.Colwidth(0) = 350
  MSFlexGrid1.Colwidth(1) = 1900
  If NumRounds < 5 Then
    For I = 2 To (rounds + 2)
      MSFlexGrid1.Colwidth(I) = (GridWid - 3000) / (rounds)
    Next I
  Else
    For I = 2 To (rounds + 2)
      MSFlexGrid1.Colwidth(I) = (GridWid - 3065) / (rounds)
    Next I
  End If
  MSFlexGrid1.Rows = (NumPilots + 1)
  If NumPilots > 12 Then
    MSFlexGrid1.ScrollBars = 2
  ElseIf NumRounds > 5 Then
    MSFlexGrid1.ScrollBars = 1
    
  ElseIf NumPilots > 12 And NumRounds > 5 Then
    MSFlexGrid1.ScrollBars = 3
    
  Else
    MSFlexGrid1.ScrollBars = 0
  End If
  
  For I = 1 To NumPilots
    MSFlexGrid1.RowHeight(I) = GridHgt / 15
  Next I
 ' Clear the contents of the grid.
  
  'Initialize the line break variable.
  NL = Chr(10) & Chr(13)
  ' First row = current row.
  MSFlexGrid1.Row = 0
  MSFlexGrid1.Col = 0
  MSFlexGrid1.FixedAlignment(0) = 2
  MSFlexGrid1.Text = ""
  
  MSFlexGrid1.Row = 0
  'Put column headings in the first row.
  MSFlexGrid1.Col = 1
  MSFlexGrid1.FixedAlignment(1) = 2
  MSFlexGrid1.Text = "NAME"
  MSFlexGrid1.Col = 2
  MSFlexGrid1.FixedAlignment(1) = 2
  MSFlexGrid1.Text = "TOTAL"
  For Col2Fill = 3 To (rounds + 2)
     MSFlexGrid1.Col = Col2Fill
     MSFlexGrid1.FixedAlignment(Col2Fill) = 1
     MSFlexGrid1.Text = "R " & (Col2Fill - 2)
  Next Col2Fill
  'Center text in the first column.
  MSFlexGrid1.FixedAlignment(0) = 2
  'Put row headings in the first column.
  MSFlexGrid1.Col = 0
  For Row2Fill = 1 To NumPilots
    MSFlexGrid1.Row = Row2Fill
    MSFlexGrid1.Text = Row2Fill
  Next Row2Fill
  LeaderboardDone = True
End Sub

Public Sub Cleargrid()
Dim GridHgt
 Dim GridWid
 
 'GridHgt = MSFlexGrid1.Height - MSFlexGrid1.RowHeight(0) - 30
 'GridWid = MSFlexGrid1.Width - MSFlexGrid1.Colwidth(0) - 30
 GridHgt = MSFlexGrid1.Height
 GridWid = MSFlexGrid1.Width
 
 MSFlexGrid1.Rows = 11
 MSFlexGrid1.Cols = 6
 For I = 1 To 10
   MSFlexGrid1.RowHeight(I) = GridHgt / 15
 Next I
 For I = 1 To 5
  MSFlexGrid1.Colwidth(I) = GridWid / 5
 Next I
 MSFlexGrid1.ScrollBars = 0
    ' Declare local variables.
    Dim Count As Integer
    ' Clear text from the first column.
    MSFlexGrid1.FixedCols = 1
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = " "
    MSFlexGrid1.Col = 0
    For Count = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = Count
        MSFlexGrid1.Text = " "
    Next Count
    ' Clear text from the first row.
    MSFlexGrid1.Row = 0
    For Count = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Col = Count
        MSFlexGrid1.Text = " "
    Next Count
     MSFlexGrid1.Clear
    ' Clear text from the nonfixed cells.
    'MSFlexGrid1.SelStartCol = 1
    'MSFlexGrid1.SelStartRow = 1
    'MSFlexGrid1.SelEndCol = MSFlexGrid1.Cols - 1
    'MSFlexGrid1.SelEndRow = MSFlexGrid1.Rows - 1
    'MSFlexGrid1.FillStyle = 1
    'MSFlexGrid1.Text = ""
    'MSFlexGrid1.FillStyle = 0
    'MSFlexGrid1.SelEndCol = 1
    'MSFlexGrid1.SelEndRow = 1
End Sub

Public Sub Leaderboard()
  
  Dim Listset As Recordset
  Dim SQLString3 As String
  Dim rounds As Integer
  Dim sco() As Single
  Dim Count As Integer
  Dim NumPilots As Integer
  
  rounds = NumRounds
  Count = 1
  Call ScoreAddUp(sco(), NumPilots)
  ' Size the nonfixed cells.
  LBSetGrid MSFlexGrid1, rounds, NumPilots
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, Pilot.First_Name, Pilot.Last_Name FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & ")) ORDER BY List.TempScore DESC;"
  Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Listset.MoveFirst
  Do Until Listset.EOF
   MSFlexGrid1.Row = Count
   MSFlexGrid1.Col = 1
   MSFlexGrid1.ColAlignment(1) = 0
   MSFlexGrid1.Text = Trim(Listset!First_Name) & " " & Trim(Listset!Last_Name)
   MSFlexGrid1.Row = Count
   MSFlexGrid1.Col = 2
   MSFlexGrid1.ColAlignment(2) = 2
   MSFlexGrid1.Text = Str(sco(Listset!Pilot_ID, 0))
   For I = 1 To rounds
    MSFlexGrid1.Col = I + 2
    MSFlexGrid1.ColAlignment(I + 2) = 2
    MSFlexGrid1.Text = Str(sco(Listset!Pilot_ID, I))
   Next I
   Listset.MoveNext
   Count = Count + 1
  Loop
  
  
  
 SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, List.TempScore,List.FinalCompScore, list.FlyoffScore From List WHERE (List.Comp_ID=" & Str(CurrentContest) & ") ORDER BY List.Pilot_ID;"
 Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
 Listset.MoveFirst
 Do Until Listset.EOF
   Listset.Edit
   Listset!TempScore = 0
   Listset.Update
   Listset.MoveNext
 Loop
 If NumRounds > 5 Then
   MSFlexGrid1.FixedCols = 2
 End If
End Sub


Public Sub FrequencyManage()
  Dim Selected As Recordset
  Dim NumFreq As Integer
  Dim SQLString As String
  Dim CheckFreq
  Dim Freq()
  Dim I As Integer
  Dim Line As String
  Dim MaxFreqUse As Integer
  
  SQLString = "SELECT * from List Where List.Comp_ID =" & Str(CurrentContest) & " ORDER BY FreqUsed ASC;"
  Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If Selected.RecordCount <> 0 Then
    Form4.List1.Clear
    Form4.Label1.Caption = "Frequency Management"
    'Find the number of different frequencys
    Selected.MoveFirst
    CheckFreq = Selected!FreqUsed
    Selected.MoveNext
    NumFreq = 1
    Do Until Selected.EOF
      If Selected!FreqUsed <> CheckFreq Then
         NumFreq = NumFreq + 1
         CheckFreq = Selected!FreqUsed
      End If
      Selected.MoveNext
    Loop
    ReDim Freq(NumFreq, 2)
    MaxFreqUse = 1
    Selected.MoveFirst
    I = 1
    Freq(I, 1) = Selected!FreqUsed
    Do Until Selected.EOF
      If Selected!FreqUsed = Freq(I, 1) Then
         Freq(I, 2) = Freq(I, 2) + 1
         If Freq(I, 2) > MaxFreqUse Then
           MaxFreqUse = Freq(I, 2)
         End If
      Else
         I = I + 1
         Freq(I, 1) = Selected!FreqUsed
         Freq(I, 2) = Freq(I, 2) + 1
         If Freq(I, 2) > MaxFreqUse Then
           MaxFreqUse = Freq(I, 2)
         End If
      End If
      Selected.MoveNext
    Loop
    Line = "Frequency" & "       " & "Usage"
    Form4.List1.AddItem Line
    Line = ""
    For I = 1 To NumFreq
      Line = "  " & Str(Freq(I, 1)) & "            " & Str(Freq(I, 2))
      Form4.List1.AddItem Line
      Line = ""
    Next I
    Form4.List1.AddItem " "
    Line = "Number of Competitors = " & Str(Selected.RecordCount)
    Form4.List1.AddItem Line
    Form4.List1.AddItem Str(MaxFreqUse)
    MostUsedFreq = MaxFreqUse
    Form4.Show
  Else
    MsgBox ("This contest has no pilots entered")
  End If
End Sub

Public Sub ContestReset()
Dim DeleteSet As Recordset
Dim SQLString As String
Dim MsgString As String


  SQLString = "SELECT * FROM Scores WHERE Scores.Comp_ID = " & Str(CurrentContest) & " ORDER BY Pilot_ID;"
  Set DeleteSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If DeleteSet.RecordCount = 0 Then
    MsgBox ("There are no scores in this Contest")
  Else
    MsgString = "You are about to delete all the scores for this contest - are you sure?"
    If MsgBox(MsgString, 1) = vbOK Then
      DeleteSet.MoveFirst
      Do Until DeleteSet.EOF
        DeleteSet.Delete
        DeleteSet.MoveNext
      Loop
      DeleteSet.Close
      SQLString = "SELECT * FROM Competition WHERE (Competition.Comp_ID = " & Str(CurrentContest) & ");"
      Set DeleteSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
      DeleteSet.Edit
      DeleteSet!RoundsDone = False
      DeleteSet!Completed = False
      DeleteSet!RoundsScored = 0
      DeleteSet.Update
      DeleteSet.Close
      Form1.Text3.Text = "NO"
      
      Completed = False
      CompRoundsDone = False
      SQLString = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, List.FinalCompScore From List WHERE (List.Comp_ID=" & Str(CurrentContest) & ") ORDER BY List.Pilot_ID;"
      Set DeleteSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
      DeleteSet.MoveFirst
      Do Until DeleteSet.EOF
        DeleteSet.Edit
        DeleteSet!FinalCompScore = 0
        DeleteSet.Update
        DeleteSet.MoveNext
      Loop
      DeleteSet.Close
      MsgBox ("All Scores Deleted")
    Else
      MsgBox ("No Records were deleted")
    End If
  End If
End Sub

Public Sub CompleteContest()
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
  
  MsgString = "You are about to set this contest as Complete - are you sure?"
  If MsgBox(MsgString, 1) = vbOK Then
    rounds = NumRounds
    Call ScoreAddUp(sco(), NumPilots)
    SQLString = "SELECT * FROM Competition WHERE (Competition.Comp_ID = " & Str(CurrentContest) & ");"
    Set UpdateTable = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    UpdateTable.Edit
    UpdateTable!Completed = True
    UpdateTable.Update
    Form1.Text3.Text = "Yes"
    SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, List.TempScore,List.FinalCompScore, List.FlyoffScore From List WHERE (List.Comp_ID=" & Str(CurrentContest) & ") ORDER BY List.TempScore DESC;"
    Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
    Listset.MoveFirst
    Place = 1
    Do Until Listset.EOF
      Listset.Edit
      Listset!FinalCompScore = Listset!TempScore
      Listset!TempScore = 0
      Listset!FlyoffScore = Place 'Element Flyoffscore stores the pilots place in the contest
      Listset.Update
      Listset.MoveNext
      Place = Place + 1
    Loop
    MsgBox ("The Contest is Finished")
  Else
    MsgBox ("The Contest is not Completed")
  End If
End Sub






Private Sub ContestView()
Dim SQLString As String
Dim SelectedSet As Recordset
Dim RoundNum As Integer
Dim Slot As Integer
Dim Line As String
  
  GeneralDisplay.List1.Clear
  For RoundNum = 1 To NumRounds
    Line = "Round " & Str(RoundNum)
    GeneralDisplay.List1.AddItem Line
    Slot = 1
    Do Until Slot > NumSlots
      Line = "Slot " & Str(Slot)
      GeneralDisplay.List1.AddItem Line
      SQLString = "SELECT DISTINCTROW List.Comp_ID, Pilot.First_Name, Pilot.Last_Name, List.Pilot_ID, Scores.Round, Scores.Slot FROM Pilot INNER JOIN ((Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID)) ON (Pilot.Pilot_ID = List.Pilot_ID) AND (Pilot.Pilot_ID = Scores.Pilot_ID) Where ((List.Comp_ID = " & Str(CurrentContest) & ") And (Scores.Round = " & Str(RoundNum) & ") And (Scores.Slot = " & Str(Slot) & ")AND (Scores.Task = '" & Currenttask & "')) ORDER BY Pilot.Last_Name ASC;"
      Set SelectedSlot = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
      If SelectedSlot.RecordCount > 0 Then
        SelectedSlot.MoveFirst
        Line = ""
        Do Until SelectedSlot.EOF
          Line = Line & "  " & Left(Trim(SelectedSlot!First_Name), 1) & "." & SelectedSlot!Last_Name
          SelectedSlot.MoveNext
        Loop
        GeneralDisplay.List1.AddItem Line
      Else
        Line = "No Pilots"
        GeneralDisplay.List1.AddItem Line
      End If
      Line = ""
      GeneralDisplay.List1.AddItem Line
      Slot = Slot + 1
    Loop
  Next RoundNum
  GeneralDisplay.Caption = "Round/Slot Allocation"
  GeneralDisplay.Show
End Sub

Public Sub PilotsView()
Dim PilotSet As Recordset
Dim SQLString1 As String
Dim NumPilots As Integer
Dim P As Integer
Dim Count As Integer
Dim Line As String

  GeneralDisplay.List1.Clear
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID,List.FreqUsed, List.Team_ID, Pilot.First_Name, Pilot.Last_Name, Pilot.Freq1, Pilot.Freq2, Pilot.Club, Pilot.Number FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & ")) ORDER BY Pilot.Last_Name ASC;"
  Set PilotSet = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  PilotSet.MoveFirst
  GeneralDisplay.Caption = "List of Pilots in Contest"
  Line = "Freq Used          Name"
  GeneralDisplay.List1.AddItem Line
  Line = " "
  GeneralDisplay.List1.AddItem Line
  Do Until PilotSet.EOF
    Line = Format(Str(PilotSet!FreqUsed), "##.000") & "           " & Trim(PilotSet!First_Name) & " " & Trim(PilotSet!Last_Name)
    GeneralDisplay.List1.AddItem Line
    PilotSet.MoveNext
  Loop
  GeneralDisplay.Show

End Sub

Public Sub TeamsView()
Dim PilotSet As Recordset
Dim SQLString1 As String
Dim NumPilots As Integer
Dim CurrentTeam As Integer
Dim Line As String

  On Error GoTo Errorhandler
  GeneralDisplay.List1.Clear
  CurrentTeam = 0
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID,List.FreqUsed, List.Team_ID, List.InTeam, Pilot.First_Name, Pilot.Last_Name, Pilot.Freq1, Pilot.Freq2, Pilot.Club, Pilot.Number FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & ")) ORDER BY List.Team_ID ASC;"
  Set PilotSet = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  PilotSet.MoveFirst
  GeneralDisplay.Caption = "List of Teams in Contest"
  If PilotSet!InTeam = False Then
    Line = "Pilots not in Teams"
    GeneralDisplay.List1.AddItem Line
    Do Until PilotSet!InTeam = True
      Line = "  " & Trim(PilotSet!First_Name) & " " & Trim(PilotSet!Last_Name)
      GeneralDisplay.List1.AddItem Line
      PilotSet.MoveNext
    Loop
  End If
  
  Do Until PilotSet.EOF
    If CurrentTeam <> PilotSet!Team_ID Then
      CurrentTeam = PilotSet!Team_ID
      Line = ""
      GeneralDisplay.List1.AddItem Line
      Line = "Team " & Str(CurrentTeam)
      GeneralDisplay.List1.AddItem Line
    End If
    Line = "  " & Trim(PilotSet!First_Name) & " " & Trim(PilotSet!Last_Name)
    GeneralDisplay.List1.AddItem Line
    PilotSet.MoveNext
  Loop
  GeneralDisplay.Show
Exit Sub
Errorhandler:
  MsgBox ("No Teams Entered")

End Sub
Sub Form_UnLoad(Cancel As Integer)
'*** Code added by VB HelpWriter ***
'*** Subroutine added by VB HelpWriter ***
    QuitHelp
'***********************************
End Sub

Private Sub saveproc()
Dim LeaderSet As Recordset
  Dim Listset As Recordset
  Dim SQLString1 As String
  Dim SQLString2 As String
  Dim SQLString3 As String
  Dim rounds As Integer
  Dim I As Integer
  Dim T As Integer
  Dim Line As String
  Dim Line2 As String
  Dim Place As Integer
  Dim Tab1 As Integer
  Dim Count As Integer
  Dim sco() As Single
  Dim maxnum As Integer
  Dim MinScore As Single
  Dim AMinscore As Single
  Dim BMinscore As Single
  Dim CMinscore As Single
  Dim ThisTask As String
  
  rounds = NumRounds
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, List.TempScore From List WHERE (List.Comp_ID=" & Str(CurrentContest) & ") ORDER BY List.Pilot_ID;"
  Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Listset.MoveFirst
  NumPilots = Listset.RecordCount
  'Find the largest Pilot ID number
  Do Until Listset.EOF
    If Listset!Pilot_ID > maxnum Then
      maxnum = Listset!Pilot_ID
    End If
    Listset.MoveNext
  Loop
  Listset.MoveFirst
  ReDim sco(maxnum, rounds)
  ' Size the nonfixed cells.
  LBSetGrid MSFlexGrid1, rounds, NumPilots
  
  If CurrentContestType <> "F3B" Then
    Do Until Listset.EOF
      MinScore = 1000
      SQLString2 = "SELECT DISTINCTROW Scores.Comp_ID, Scores.Pilot_ID, Scores.Round, Scores.Score, Scores.Task From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Pilot_ID=" & Str(Listset!Pilot_ID) & ") AND (Scores.Task='A')) ORDER BY Scores.Round ASC;"
      Set Scoreset = F3JDb.OpenRecordset(SQLString2, dbOpenDynaset)
      Scoreset.MoveFirst
      I = 1
      Do Until Scoreset.EOF
        If Scoreset!Score < MinScore Then
          MinScore = Scoreset!Score
        End If
        sco(Listset!Pilot_ID, I) = Scoreset!Score
        sco(Listset!Pilot_ID, 0) = sco(Listset!Pilot_ID, 0) + Scoreset!Score
        I = I + 1
        Scoreset.MoveNext
      Loop
      'Discard Lowest score if there are more that 5 rounds flown
      If I > 5 Then
        sco(Listset!Pilot_ID, 0) = sco(Listset!Pilot_ID, 0) - MinScore
      End If
      Listset.Edit
      Listset!TempScore = sco(Listset!Pilot_ID, 0)
      Listset.Update
      Listset.MoveNext
    Loop
  Else
    Do Until Listset.EOF
      AMinscore = 1000
      BMinscore = 1000
      CMinscore = 1000
      For T = 1 To 3
        If T = 1 Then
            ThisTask = "A"
          ElseIf T = 2 Then
            ThisTask = "B"
          ElseIf T = 3 Then
            ThisTask = "C"
        End If
        SQLString2 = "SELECT DISTINCTROW Scores.Comp_ID, Scores.Pilot_ID, Scores.Round, Scores.Score, Scores.Task From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Pilot_ID=" & Str(Listset!Pilot_ID) & ") AND (Scores.Task='" & ThisTask & "'))ORDER BY Scores.Round ASC;"
        Set Scoreset = F3JDb.OpenRecordset(SQLString2, dbOpenDynaset)
        Scoreset.MoveFirst
        I = 1
        Do Until Scoreset.EOF
          If T = 1 Then
            If Scoreset!Score < AMinscore Then
              AMinscore = Scoreset!Score
            End If
          ElseIf T = 2 Then
            If Scoreset!Score < BMinscore Then
              BMinscore = Scoreset!Score
            End If
          ElseIf T = 3 Then
            If Scoreset!Score < CMinscore Then
              CMinscore = Scoreset!Score
            End If
          End If
        
          sco(Listset!Pilot_ID, I) = sco(Listset!Pilot_ID, I) + Scoreset!Score
          sco(Listset!Pilot_ID, 0) = sco(Listset!Pilot_ID, 0) + Scoreset!Score
          Scoreset.MoveNext
          I = I + 1
        Loop
      Next T
        
      'Discard Lowest scores if there are more that 5 rounds flown
      If I > 5 Then
        sco(Listset!Pilot_ID, 0) = sco(Listset!Pilot_ID, 0) - (AMinscore + BMinscore + CMinscore)
      End If
      Listset.Edit
      Listset!TempScore = sco(Listset!Pilot_ID, 0)
      Listset.Update
      Listset.MoveNext
    Loop
  End If
  
  
  Count = 1
  'SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, List.TempScore FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & ")) ORDER BY List.TempScore DESC;"
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, Pilot.First_Name, Pilot.Last_Name FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & ")) ORDER BY List.TempScore DESC;"
  
  Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Listset.MoveFirst
  Do Until Listset.EOF
   MSFlexGrid1.Row = Count
   MSFlexGrid1.Col = 1
   MSFlexGrid1.ColAlignment(1) = 0
   MSFlexGrid1.Text = Trim(Listset!First_Name) & " " & Trim(Listset!Last_Name)
   MSFlexGrid1.Row = Count
   MSFlexGrid1.Col = 2
   MSFlexGrid1.ColAlignment(2) = 2
   'MSFlexGrid1.Text = Str(ListSet!TempScore)
   MSFlexGrid1.Text = Str(sco(Listset!Pilot_ID, 0))
   
   For I = 1 To rounds
    'SQLString2 = "SELECT DISTINCTROW Scores.Comp_ID, Scores.Pilot_ID, Scores.Round, Scores.Score, Scores.Task From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Pilot_ID=" & Str(ListSet!Pilot_ID) & ") AND (Scores.Round=" & Str(I) & ") AND (Scores.Task='A'));"
    'Set ScoreSet = LeaderDb.OpenRecordset(SQLString2, dbOpenDynaset)
    MSFlexGrid1.Col = I + 2
    MSFlexGrid1.ColAlignment(I + 2) = 2
    'MSFlexGrid1.Text = Str(ScoreSet!Score)
    MSFlexGrid1.Text = Str(sco(Listset!Pilot_ID, I))
     
   Next I
   Listset.MoveNext
   Count = Count + 1
  Loop
  
  
  
 SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, List.TempScore,List.FinalCompScore, list.FlyoffScore From List WHERE (List.Comp_ID=" & Str(CurrentContest) & ") ORDER BY List.Pilot_ID;"
 Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
 Listset.MoveFirst
 Do Until Listset.EOF
   Listset.Edit
   Listset!FinalCompScore = Listset!TempScore
   Listset!TempScore = 0
   Listset.Update
   Listset.MoveNext
 Loop
 If NumRounds > 5 Then
   MSFlexGrid1.FixedCols = 2
 End If
End Sub


