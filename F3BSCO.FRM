VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form EnterF3BScores 
   ClientHeight    =   6240
   ClientLeft      =   735
   ClientTop       =   1230
   ClientWidth     =   6825
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   570
   LinkTopic       =   "Form5"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6240
   ScaleWidth      =   6825
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   1080
      WhatsThisHelpID =   310
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   1080
      WhatsThisHelpID =   310
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   5760
      WhatsThisHelpID =   310
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   5760
      WhatsThisHelpID =   310
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3720
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      WhatsThisHelpID =   310
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      WhatsThisHelpID =   310
      Width           =   495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6588
      _Version        =   393216
      Tab             =   1
      TabHeight       =   529
      TabCaption(0)   =   "TASK A"
      TabPicture(0)   =   "F3BSCO.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MSFlexGrid1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "TASK B"
      TabPicture(1)   =   "F3BSCO.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "MSFlexGrid1(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "VScroll4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "TASK C"
      TabPicture(2)   =   "F3BSCO.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSFlexGrid1(2)"
      Tab(2).Control(1)=   "Check1"
      Tab(2).Control(2)=   "Text5"
      Tab(2).ControlCount=   3
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70800
         TabIndex        =   18
         Text            =   "Times are always ss:hh"
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4320
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3120
         WhatsThisHelpID =   310
         Width           =   375
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   495
         Left            =   4680
         TabIndex        =   12
         Top             =   3120
         WhatsThisHelpID =   310
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Score Speed Task with all pilots in round"
         Height          =   255
         Left            =   -74760
         TabIndex        =   11
         Top             =   3360
         Width           =   5655
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2535
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4471
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2775
         Index           =   2
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4895
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3135
         Index           =   0
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5530
         _Version        =   393216
      End
      Begin VB.Label Label2 
         Caption         =   "Flight Group"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   3240
         Width           =   1095
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Slot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   720
      WhatsThisHelpID =   310
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Round"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   720
      WhatsThisHelpID =   310
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   " F3B Flight Time Input Form"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   120
      WhatsThisHelpID =   310
      Width           =   4215
   End
End
Attribute VB_Name = "EnterF3BScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PilotTable As Recordset
Public ScoreTable As Recordset
Public Selected As Recordset
Public SelectedSlot As Recordset
Public SortedSlot As Recordset
Public NumPilots As Integer
Dim RoundNum As String
Dim Slot As String
Dim Group As String
Dim ThisNumSlots As Integer

Function SetTask(SStab As Control) As String
  If SStab.Tab = 0 Then
    SetTask = "A"
  ElseIf SStab.Tab = 1 Then
    SetTask = "B"
  ElseIf SStab.Tab = 2 Then
    SetTask = "C"
  End If
End Function



Private Sub Check1_Click()
  Dim SQLString3 As String
  
  Dim Mins As Integer
  Dim Sec As Single
  Dim Inc As Integer
  Dim Count As Integer
  Dim SQLString4 As String
  Dim ThisSelected As Recordset
  Dim ThisRound As String
  
  
  ThisRound = Text1.Text
  Currenttask = SetTask(SSTab1)
  RoundNum = Text1.Text
  Slot = Text2.Text
  
  If Currenttask = "C" Then
    If Check1.Value = 1 Then
      SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.Task, Scores.RawScore, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Task='" & Currenttask & "'))ORDER BY Pilot.Last_Name;"
    Else
      SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.Task, Scores.RawScore, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & ") AND (Scores.Task='" & Currenttask & "'))ORDER BY Pilot.Last_Name;"
    End If
  Else
    SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.Task, Scores.RawScore FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & ") AND (Scores.Task='" & Currenttask & "'))ORDER BY Pilot.Last_Name;"
  End If
  Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
  If SelectedSlot.RecordCount <> 0 Then
    'checkdupe SelectedSlot
    NumPilots = SelectedSlot.RecordCount
    SelectedSlot.MoveFirst
    SetGrid NumPilots, SSTab1.Tab
    FillGrid SelectedSlot, SSTab1.Tab
        
  Else
   MsgBox ("No Pilots in Slot")
  End If
  If Check1.Value = 1 Then
    VScroll2.Enabled = False
  Else
    VScroll2.Enabled = True
  End If
End Sub

Private Sub Command3_Click()
  
  Call F3BScore(RoundNum, Slot, NumPilots, SelectedSlot)
  
  VScroll1.Enabled = True
  VScroll2.Enabled = True
  Command6.Enabled = True
End Sub



Private Sub Command6_Click()
  If FromView = True Then
    FromView = False
  End If
  Unload Me
End Sub

Public Sub Form_Load()
  Dim SQLString1 As String
  Dim SQLString2 As String
  Dim NumPilots As Integer
  Dim NumScores As Integer
  Dim SQLString3 As String
  Dim Mins As Integer
  Dim Sec As Single
  Dim SQLString4 As String
  Dim ThisSelected As Recordset
  Dim ThisRound As String
  Dim ThisTask As String
  Dim Count As Integer
  Dim Col As Integer
  Dim Row As Integer
  
  Text1.Text = 1
  Text2.Text = 1
  If FlightGroups = False Then
   Text4.Text = 0
  Else
   Text4.Text = 1
  End If
  ThisRound = Text1.Text
  SSTab1.Tab = 0
  Count = 1
  Currenttask = SetTask(SSTab1)
  SQLString4 = "SELECT DISTINCTROW Scores.Comp_ID, Scores.Round, Scores.Slot From Scores WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & Str(ThisRound) & ") AND (Scores.Slot=" & Str(Count) & "));"
  Set ThisSelected = F3JDb.OpenRecordset(SQLString4, dbOpenDynaset)
  Do Until ThisSelected.RecordCount = 0
    Count = Count + 1
    SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Scores.Round, Scores.Slot From Scores WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & Str(ThisRound) & ") AND (Scores.Slot=" & Str(Count) & "));"
    Set ThisSelected = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
  Loop
  ThisNumSlots = Count - 1
  RoundNum = Text1.Text
  Slot = Text2.Text
  
  SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.Task, Scores.RawScore FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & ") AND (Scores.Task='" & Currenttask & "'))ORDER BY Pilot.Last_Name;"
      
  If Currenttask = "B" Then
    Group = Text4.Text
  Else
    Group = 0
  End If
  SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.Task, Scores.RawScore, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & ") AND (Scores.Task='" & Currenttask & "') AND (Scores.FlightGroup = " & Group & "))ORDER BY Pilot.Last_Name;"
  Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
  If SelectedSlot.RecordCount <> 0 Then
    'checkdupe SelectedSlot
    'Set Data1.Recordset = SelectedSlot
    'DBGrid1.Refresh
    NumPilots = SelectedSlot.RecordCount
    SelectedSlot.MoveFirst
    SetGrid NumPilots, SSTab1.Tab
    FillGrid SelectedSlot, SSTab1.Tab
  Else
    MsgBox ("No Pilots in Slot")
  End If
  
  Command3.Caption = "Store Scores for Task"
  Command6.Caption = "Done"
  If FlightGroups = False Then
    Label2.Visible = False
    Text4.Visible = False
    VScroll4.Visible = False
  End If
  
    
 End Sub


Public Sub F3BScore(RoundNum, Slot, NumPilots, SelectedSlot As Recordset)
Dim String1 As String
Dim FirstScore
Dim PilotScore
Dim UpdateTable As Recordset
Dim Count As Integer
Dim SQLString As String
Dim FlightPenalty() As Single
Dim LandBonus As Boolean
Dim ManualPenalty() As Single
Dim Row As Integer
Dim FlightG As Boolean
Dim Column As Integer
Dim MaxScore As Single
Dim MinScore As Single
Dim NormScore() As Single
Dim FlightTime() As Single
Dim ValLandBonus() As Single
Dim Rawscore() As Single
Dim Tagged As Boolean
Dim Refly As Boolean
Dim Mins As Integer
Dim Secs As Single
Dim Pilots As Integer
Dim T As Integer
  If Currenttask = "C" And NumPilots > 20 Then
    ReDim Rawscore(NumPilots + 2)
    ReDim NormScore(NumPilots + 2)
    ReDim FlightTime(NumPilots + 2)
    ReDim ValLandBonus(NumPilots + 2)
    ReDim FlightPenalty(NumPilots + 2)
    ReDim ManualPenalty(NumPilots + 2)
  Else
    ReDim Rawscore(20)
    ReDim NormScore(20)
    ReDim FlightTime(20)
    ReDim ValLandBonus(20)
    ReDim FlightPenalty(20)
    ReDim ManualPenalty(20)
  End If
  
  FlightG = False
  MinScore = 10000
  T = SSTab1.Tab
  Pilots = NumPilots
  Row = 0
  MSFlexGrid1(T).Row = Row
  NumPilots = SelectedSlot.RecordCount
  SelectedSlot.MoveFirst
  Refly = False
  Do Until SelectedSlot.EOF
    If SelectedSlot!Score > 0 Then
      Refly = True
    End If
    SelectedSlot.MoveNext
  Loop
  SelectedSlot.MoveFirst
  
  Do Until Row = NumPilots
    Row = Row + 1
    MSFlexGrid1(T).Row = Row
    If Currenttask = "A" Then
      LandBonus = True
      MSFlexGrid1(T).Col = 1
      FlightTime(Row) = Val(MSFlexGrid1(T).Text) * 60
      MSFlexGrid1(T).Col = 2
      FlightTime(Row) = FlightTime(Row) + Val(Trim(MSFlexGrid1(T).Text))
      MSFlexGrid1(T).Col = 4
      ManualPenalty(Row) = Val(MSFlexGrid1(T).Text)
      If FlightTime(Row) > (max * 60) And FlightTime(Row) < ((max + 0.5) * 60) Then
        FlightPenalty(Row) = FlightPenalty(Row) + ((FlightTime(Row) - (max * 60)) * 2)
        'FlightTime(Row) = (max * 60)
        LandBonus = True
      ElseIf FlightTime(Row) >= ((max + 0.5) * 60) Then
        FlightPenalty(Row) = FlightPenalty(Row) + ((FlightTime(Row) - (max * 60)) * 2)
        LandBonus = False
      End If
    
    ElseIf Currenttask = "B" Then
      MSFlexGrid1(T).Col = 1
      FlightTime(Row) = Val(Trim(MSFlexGrid1(T).Text))
      MSFlexGrid1(T).Col = 4
      ManualPenalty(Row) = ManualPenalty(Row) + Val(MSFlexGrid1(T).Text)
      LandBonus = False
    ElseIf Currenttask = "C" Then
      MSFlexGrid1(T).Col = 1
      FlightTime(Row) = Val(Trim(MSFlexGrid1(T).Text)) * 100
      MSFlexGrid1(T).Col = 2
      FlightTime(Row) = FlightTime(Row) + Val(Trim(MSFlexGrid1(T).Text))
      MSFlexGrid1(T).Col = 4
      ManualPenalty(Row) = ManualPenalty(Row) + Val(MSFlexGrid1(T).Text)
      LandBonus = False
    End If
    MSFlexGrid1(T).Col = 3
    If Val(MSFlexGrid1(T).Text) < 30 Then
      If LandBonus = True Then
        MSFlexGrid1(T).Col = 3
        Select Case Val(MSFlexGrid1(T).Text)
        Case 1
          ValLandBonus(Row) = 100
        Case 2
          ValLandBonus(Row) = 95
        Case 3
          ValLandBonus(Row) = 90
        Case 4
          ValLandBonus(Row) = 85
        Case 5
          ValLandBonus(Row) = 80
        Case 6
          ValLandBonus(Row) = 75
        Case 7
          ValLandBonus(Row) = 70
        Case 8
          ValLandBonus(Row) = 65
        Case 9
          ValLandBonus(Row) = 60
        Case 10
          ValLandBonus(Row) = 55
        Case 11
          ValLandBonus(Row) = 50
        Case 12
          ValLandBonus(Row) = 45
        Case 13
          ValLandBonus(Row) = 40
        Case 14
          ValLandBonus(Row) = 34
        Case 15
          ValLandBonus(Row) = 30
        Case Else
          ValLandBonus(Row) = 0
        End Select
        
      Else
        ValLandBonus(Row) = 0
        
      End If
    Else
      MSFlexGrid1(T).Col = 3
      ValLandBonus(Row) = Val(MSFlexGrid1(T).Text)
    End If
    Rawscore(Row) = FlightTime(Row) + ValLandBonus(Row) - FlightPenalty(Row)
    If Rawscore(Row) > MaxScore Then
      MaxScore = Rawscore(Row)
    End If
    If Rawscore(Row) < MinScore Then
      If Rawscore(Row) <> 0 Then
        MinScore = Rawscore(Row)
      End If
    End If
  Loop
  Rem normalise scores
  If Currenttask = "C" Then
    For Row = 1 To NumPilots
      If Rawscore(Row) <> 0 Then
        NormScore(Row) = Round((MinScore / Rawscore(Row)) * 1000, 4)
      Else
        NormScore(Row) = 0
      End If
      Next Row
  Else
    If MaxScore <> 0 Then
      For Row = 1 To NumPilots
        NormScore(Row) = Round((Rawscore(Row) / MaxScore) * 1000, 4)
      Next Row
    End If
  End If
  Row = 1
  If Refly = True Then
    If MsgBox("This slot has already been done" & Chr(13) & "Do You want to rescore this slot", 19) <> vbYes Then
      MsgBox ("Scoring Cancelled")
    Else
      Tagged = False
      MSFlexGrid1(T).Col = 6
      For Row = 1 To NumPilots
        MSFlexGrid1(T).Row = Row
        If MSFlexGrid1(T).Text = "*" Then
          Tagged = True
        End If
      Next Row
      If Tagged = False Then
        If MsgBox("You have to tag a pilot to indicate" & Chr(13) & "who asked for the refly" & Chr(13) & "Do you want to rescore the round anyway?", 19) <> vbYes Then
          Exit Sub
        Else
          SelectedSlot.MoveFirst
          Row = 1
          Do Until SelectedSlot.EOF
            SelectedSlot.Edit
            SelectedSlot!Score = Round(NormScore(Row), 0)
            SelectedSlot!Rawscore = Round(Rawscore(Row), 0)
            SelectedSlot!Flight_Time = FlightTime(Row)
            SelectedSlot!LandBonus = ValLandBonus(Row)
            SelectedSlot!Penalty = FlightPenalty(Row)
            SelectedSlot!Res2 = ManualPenalty(Row)
            SelectedSlot.Update
            Row = Row + 1
            If Row < NumPilots Then
              MSFlexGrid1(T).Row = Row
            End If
            SelectedSlot.MoveNext
          Loop
          FillGrid SelectedSlot, T
          MsgBox ("Slot Rescored")
        End If
      Else
        SelectedSlot.MoveFirst
        Row = 1
        MSFlexGrid1(T).Row = Row
        Do Until SelectedSlot.EOF
          MSFlexGrid1(T).Col = 6
          MSFlexGrid1(T).Row = Row
          If MSFlexGrid1(T).Text <> "*" Then
            If SelectedSlot!Score < Round(NormScore(Row), 0) Then
              SelectedSlot.Edit
              SelectedSlot!Score = Round(NormScore(Row), 0)
              SelectedSlot!Rawscore = Round(Rawscore(Row), 0)
              SelectedSlot!Flight_Time = FlightTime(Row)
              SelectedSlot!LandBonus = ValLandBonus(Row)
              SelectedSlot!Penalty = FlightPenalty(Row)
              SelectedSlot!Res2 = ManualPenalty(Row)
              SelectedSlot.Update
            End If
          Else
              SelectedSlot.Edit
              SelectedSlot!Score = Round(NormScore(Row), 0)
              SelectedSlot!Rawscore = Round(Rawscore(Row), 0)
              SelectedSlot!Flight_Time = FlightTime(Row)
              SelectedSlot!LandBonus = ValLandBonus(Row)
              SelectedSlot!Penalty = FlightPenalty(Row)
              SelectedSlot!Res2 = ManualPenalty(Row)
              SelectedSlot.Update
          End If
          Row = Row + 1
          If Row < NumPilots Then
            MSFlexGrid1(T).Row = Row
          End If
          SelectedSlot.MoveNext
        Loop
        FillGrid SelectedSlot, T
        SelectedSlot.MoveFirst
        MsgBox ("Slot Rescored")
      End If
    End If
  Else
    Row = 1
    SelectedSlot.MoveFirst
    Do Until SelectedSlot.EOF
      SelectedSlot.Edit
      SelectedSlot!Score = Round(NormScore(Row), 0)
      SelectedSlot!Rawscore = Round(Rawscore(Row), 0)
      SelectedSlot!Flight_Time = FlightTime(Row)
      SelectedSlot!LandBonus = ValLandBonus(Row)
      SelectedSlot!Penalty = FlightPenalty(Row)
      SelectedSlot!Res2 = ManualPenalty(Row)
      SelectedSlot.Update
      Row = Row + 1
      SelectedSlot.MoveNext
    Loop
    FillGrid SelectedSlot, T
    MsgBox ("Slot Scored")
  End If 'refly
End Sub

Public Sub ClubScore(ByVal SelectedSlot)
  SelectedSlot.MoveFirst
  Do Until SelectedSlot.EOF = True
    SelectedSlot.Edit
    SelectedSlot!Score = SelectedSlot!Flight_Time
    SelectedSlot.Update
    SelectedSlot.MoveNext
  Loop
  End Sub

Private Sub msflexgrid1_DblClick(Index As Integer)
  Dim R As Integer
  
  MSFlexGrid1(Index).Col = 6
  If MSFlexGrid1(Index).Text = "" Then
    MSFlexGrid1(Index).Text = "*"
    MSFlexGrid1(Index).Col = 4
    For R = 1 To (MSFlexGrid1(Index).Rows - 1)
      MSFlexGrid1(Index).Row = R
      MSFlexGrid1(Index).Text = ""
    Next R
    If MsgBox("Do you want all the scores cleared", 19) = vbYes Then
      ClearScores SSTab1.Tab
    End If
  Else
    MSFlexGrid1(Index).Text = ""
  End If
End Sub

Private Sub msflexgrid1_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Text3.Text <> "" Then
      MSFlexGrid1(Index).Text = Text3.Text
      Text3.Text = ""
    End If
  ElseIf KeyAscii = 27 Then
    Text3.Text = ""
  Else
    Text3.Text = Text3.Text & Chr(KeyAscii)
    Srow = MSFlexGrid1(Index).Row
    Scol = MSFlexGrid1(Index).Col
  End If
End Sub


Private Sub msflexgrid1_SelChange(Index As Integer)
  Dim Pcol As Integer
  Dim Prow As Integer

If Updated = True Then
  If Text3.Text <> "" Then
    Pcol = MSFlexGrid1(Index).Col
    Prow = MSFlexGrid1(Index).Row
    MSFlexGrid1(Index).Row = Srow
    MSFlexGrid1(Index).Col = Scol
    If MSFlexGrid1(Index).Col = 2 Then
      If Index <> 2 Then
        If Val(Text3.Text) >= 60 Then
        MsgBox ("Entry must be less than 60")
        Text3.Text = ""
        MSFlexGrid1(Index).Text = ""
        MSFlexGrid1(Index).Row = Prow
        MSFlexGrid1(Index).Col = Pcol
        Exit Sub
      End If
      'ElseIf Index = 2 Then
      '  If Val(Text3.Text) > 0 And Val(Text3.Text) < 10 Then
      '    MsgBox ("Ensure that seconds are entered as 2 digits")
      '
      '    MSFlexGrid1(Index).Row = Prow
      '    MSFlexGrid1(Index).Col = Pcol
       '   Exit Sub
      '  End If
      
      Else
        If Val(Text3.Text) > 100 Then
          MsgBox ("Entry must be less that 100")
          Text3.Text = ""
          MSFlexGrid1(Index).Row = Prow
          MSFlexGrid1(Index).Col = Pcol
          Exit Sub
        End If
      End If
    End If
    If MSFlexGrid1(Index).Col = 3 Then
      If Val(Text3.Text) > 100 Then
        MsgBox ("You cannot enter more than 100 points")
        Text3.Text = ""
        MSFlexGrid1(Index).Row = Prow
        MSFlexGrid1(Index).Col = Pcol
      End If
    End If
    MSFlexGrid1(Index).Text = Text3.Text
    Text3.Text = ""
    MSFlexGrid1(Index).Row = Prow
    MSFlexGrid1(Index).Col = Pcol
  End If
End If
End Sub




Private Sub SSTab1_Click(PreviousTab As Integer)
  Dim SQLString3 As String
  
  Dim Mins As Integer
  Dim Sec As Single
  Dim Inc As Integer
  Dim Count As Integer
  Dim SQLString4 As String
  Dim ThisSelected As Recordset
  Dim ThisRound As String
  
  
  ThisRound = Text1.Text
  Currenttask = SetTask(SSTab1)
  RoundNum = Text1.Text
  Slot = Text2.Text
  If Currenttask = "B" Then
    Group = Text4.Text
  Else
    Group = 0
  End If
  
  If Currenttask = "C" Then
    If Check1.Value = 1 Then
      SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.Task, Scores.RawScore, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Task='" & Currenttask & "'))ORDER BY Pilot.Last_Name;"
    Else
      SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.Task, Scores.RawScore, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & ") AND (Scores.Task='" & Currenttask & "'))ORDER BY Pilot.Last_Name;"
    End If
  Else
    SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.Task, Scores.RawScore, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & ") AND (Scores.Task='" & Currenttask & "') AND (Scores.FlightGroup = " & Group & "))ORDER BY Pilot.Last_Name;"
  End If
  Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
  If SelectedSlot.RecordCount <> 0 Then
    'checkdupe SelectedSlot
    NumPilots = SelectedSlot.RecordCount
    SelectedSlot.MoveFirst
    SetGrid NumPilots, SSTab1.Tab
    FillGrid SelectedSlot, SSTab1.Tab
        
  Else
   MsgBox ("No Pilots in Slot")
  End If
  
End Sub

Private Sub Text1_Change()
Dim Count As Integer
Dim SQLString3 As String
Dim ThisSelected As Recordset

 Count = 1
 SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Scores.Round, Scores.Slot From Scores WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & Str(ThisRound) & ") AND (Scores.Slot=" & Str(Count) & "));"
 Set ThisSelected = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
 Do Until ThisSelected.RecordCount = 0
    Count = Count + 1
    SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Scores.Round, Scores.Slot From Scores WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & Str(ThisRound) & ") AND (Scores.Slot=" & Str(Count) & "));"
    Set ThisSelected = MyDb.OpenRecordset(SQLString3, dbOpenDynaset)
 Loop
 ThisNumSlots = Count - 1
End Sub


Private Sub Text3_Change()
  VScroll1.Enabled = False
  VScroll2.Enabled = False
  Command6.Enabled = False
End Sub

Private Sub VScroll1_Change()
Dim SQLString3 As String
  
  Dim Mins As Integer
  Dim Sec As Single
  
  VScroll1.max = NumRounds
  VScroll1.Min = 1
  Text1.Text = VScroll1.Value
  
  Currenttask = SetTask(SSTab1)
  RoundNum = Text1.Text
  Slot = Text2.Text
  If Currenttask = "B" Then
    Group = Text4.Text
  Else
    Group = 0
  End If
  SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.Task, Scores.RawScore, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & ") AND (Scores.Task='" & Currenttask & "') AND (Scores.FlightGroup = " & Group & "))ORDER BY Pilot.Last_Name;"
  Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
  If SelectedSlot.RecordCount <> 0 Then
    'checkdupe SelectedSlot
    NumPilots = SelectedSlot.RecordCount
    SelectedSlot.MoveFirst
    SetGrid NumPilots, SSTab1.Tab
    FillGrid SelectedSlot, SSTab1.Tab
  Else
    MsgBox ("No Pilots in Slot")
  End If
  End Sub

Private Sub VScroll2_Change()
  Dim SQLString3 As String
  
  Dim Mins As Integer
  Dim Sec As Single
  Dim Inc As Integer
  Dim Count As Integer
  Dim SQLString4 As String
  Dim ThisSelected As Recordset
  Dim ThisRound As String
  
  
  ThisRound = Text1.Text
  Currenttask = SetTask(SSTab1)
  
    
  Count = 1
  SQLString4 = "SELECT DISTINCTROW Scores.Comp_ID, Scores.Round, Scores.Slot From Scores WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & Str(ThisRound) & ") AND (Scores.Slot=" & Str(Count) & "));"
  Set ThisSelected = F3JDb.OpenRecordset(SQLString4, dbOpenDynaset)
  Do Until ThisSelected.RecordCount = 0
    Count = Count + 1
    SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Scores.Round, Scores.Slot From Scores WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & Str(ThisRound) & ") AND (Scores.Slot=" & Str(Count) & "));"
    Set ThisSelected = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
  Loop
  ThisNumSlots = Count - 1
  
  VScroll2.max = ThisNumSlots
  VScroll2.Min = 1
  Text2.Text = VScroll2.Value
  
  
  RoundNum = Text1.Text
  Slot = Text2.Text
  If Currenttask = "B" Then
    Group = Text4.Text
  Else
    Group = 0
  End If
  SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.Task, Scores.RawScore, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & ") AND (Scores.Task='" & Currenttask & "') AND (Scores.FlightGroup = " & Group & "))ORDER BY Pilot.Last_Name;"
  Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
  If SelectedSlot.RecordCount <> 0 Then
    'checkdupe SelectedSlot
    NumPilots = SelectedSlot.RecordCount
    SelectedSlot.MoveFirst
    
    SetGrid NumPilots, SSTab1.Tab
    FillGrid SelectedSlot, SSTab1.Tab
  Else
    MsgBox ("No Pilots in Slot")
  End If
  End Sub




Sub SetGrid(NumPilots As Integer, Index As Integer)
  Dim Col2Fill
  Dim Row2Fill
  Dim NL As String
  Dim GridHgt
  Dim GridWid
  Dim I
  
  Updated = False
  GridHgt = MSFlexGrid1(Index).Height
  GridWid = MSFlexGrid1(Index).Width
  
  MSFlexGrid1(Index).Cols = 7
  MSFlexGrid1(Index).Colwidth(0) = GridWid / 3
  
  GridWid = GridWid - MSFlexGrid1(Index).Colwidth(0) - 60
  MSFlexGrid1(Index).Colwidth(6) = (GridWid / 6)
  MSFlexGrid1(Index).Colwidth(1) = (GridWid / 6)
  MSFlexGrid1(Index).Colwidth(2) = (GridWid / 6)
  MSFlexGrid1(Index).Colwidth(3) = (GridWid / 6)
  MSFlexGrid1(Index).Colwidth(4) = (GridWid / 6)
  MSFlexGrid1(Index).Colwidth(5) = (GridWid / 6)
  MSFlexGrid1(Index).Colwidth(6) = (GridWid / 6)
      
    
  MSFlexGrid1(Index).Rows = (NumPilots + 1)
  If NumPilots > 7 Then
    MSFlexGrid1(Index).ScrollBars = 2
  Else
    MSFlexGrid1(Index).ScrollBars = 0
  End If
  MSFlexGrid1(Index).RowHeight(0) = GridHgt / 7
  For I = 1 To NumPilots
    MSFlexGrid1(Index).RowHeight(I) = GridHgt / 9
  Next I
 ' Clear the contents of the grid.
  Call Cleargrid(SSTab1.Tab)
  'Initialize the line break variable.
  NL = Chr(10) & Chr(13)
  ' First row = current row.
  MSFlexGrid1(Index).Row = 0
  MSFlexGrid1(Index).Col = 0
  MSFlexGrid1(Index).FixedAlignment(0) = 0
  MSFlexGrid1(Index).Text = "Name"
  
  MSFlexGrid1(Index).Row = 0
  'Put column headings in the first row.
  If Currenttask = "A" Then
    MSFlexGrid1(Index).Col = 1
    MSFlexGrid1(Index).FixedAlignment(1) = 2
    MSFlexGrid1(Index).Text = "Mins"
    MSFlexGrid1(Index).Col = 2
    MSFlexGrid1(Index).FixedAlignment(2) = 2
    MSFlexGrid1(Index).Text = "Sec"
    MSFlexGrid1(Index).Col = 3
    MSFlexGrid1(Index).FixedAlignment(3) = 2
    MSFlexGrid1(Index).Text = "Land"
    MSFlexGrid1(Index).Col = 4
    MSFlexGrid1(Index).FixedAlignment(4) = 2
    MSFlexGrid1(Index).Text = "Pen"
    MSFlexGrid1(Index).Col = 5
    MSFlexGrid1(Index).FixedAlignment(5) = 2
    MSFlexGrid1(Index).Text = "Score"
    MSFlexGrid1(Index).Col = 6
    MSFlexGrid1(Index).FixedAlignment(6) = 0
    MSFlexGrid1(Index).Text = "Tag"
  ElseIf Currenttask = "B" Then
    MSFlexGrid1(Index).Col = 1
    MSFlexGrid1(Index).FixedAlignment(1) = 2
    MSFlexGrid1(Index).Text = "Laps"
    MSFlexGrid1(Index).Col = 4
    MSFlexGrid1(Index).FixedAlignment(4) = 2
    MSFlexGrid1(Index).Text = "Pen"
    MSFlexGrid1(Index).Col = 5
    MSFlexGrid1(Index).FixedAlignment(5) = 2
    MSFlexGrid1(Index).Text = "Score"
    MSFlexGrid1(Index).Col = 6
    MSFlexGrid1(Index).FixedAlignment(6) = 0
    MSFlexGrid1(Index).Text = "Tag"
  ElseIf Currenttask = "C" Then
    MSFlexGrid1(Index).Col = 1
    MSFlexGrid1(Index).FixedAlignment(1) = 2
    MSFlexGrid1(Index).Text = "Secs"
    MSFlexGrid1(Index).Col = 2
    MSFlexGrid1(Index).FixedAlignment(2) = 2
    MSFlexGrid1(Index).Text = "100ths"
    MSFlexGrid1(Index).Col = 4
    MSFlexGrid1(Index).FixedAlignment(4) = 2
    MSFlexGrid1(Index).Text = "Pen"
    MSFlexGrid1(Index).Col = 5
    MSFlexGrid1(Index).FixedAlignment(5) = 2
    MSFlexGrid1(Index).Text = "Score"
    MSFlexGrid1(Index).Col = 6
    MSFlexGrid1(Index).FixedAlignment(6) = 0
    MSFlexGrid1(Index).Text = "Tag"
  End If
  For I = 1 To 5
    MSFlexGrid1(Index).ColAlignment(I) = 2
  Next I
  
  Updated = True
End Sub

Sub Cleargrid(Index As Integer)
  Dim I As Integer
  Dim J As Integer
  
  Updated = False
   ' Declare local variables.
    Dim Count As Integer
    ' Clear text from the first column.
    MSFlexGrid1(Index).FixedCols = 1
    MSFlexGrid1(Index).Col = 0
    MSFlexGrid1(Index).Row = 0
    MSFlexGrid1(Index).Text = " "
    MSFlexGrid1(Index).Col = 0
    For Count = 1 To MSFlexGrid1(Index).Rows - 1
        MSFlexGrid1(Index).Row = Count
        MSFlexGrid1(Index).Text = " "
    Next Count
    ' Clear text from the first row.
    MSFlexGrid1(Index).Row = 0
    For Count = 1 To MSFlexGrid1(Index).Cols - 1
        MSFlexGrid1(Index).Col = Count
        MSFlexGrid1(Index).Text = " "
    Next Count
    ' Clear text from the nonfixed cells.
    'MSFlexGrid1(Index).Clear
    For I = 1 To MSFlexGrid1(Index).Cols - 1
      MSFlexGrid1(Index).Col = I
      For J = 1 To MSFlexGrid1(Index).Rows - 1
        MSFlexGrid1(Index).Row = J
        MSFlexGrid1(Index).FillStyle = 1
        MSFlexGrid1(Index).Text = ""
        MSFlexGrid1(Index).FillStyle = 0
      Next J
    Next I
   'MSFlexGrid1(Index).Col = 1
      
   ' MSFlexGrid1(Index).Row = 1
   ' MSFlexGrid1(Index).SelEndCol = MSFlexGrid1(Index).Cols - 1
   ' MSFlexGrid1(Index).selendRow = MSFlexGrid1(Index).Rows - 1
   ' MSFlexGrid1(Index).FillStyle = 1
   ' MSFlexGrid1(Index).Text = ""
   ' MSFlexGrid1(Index).FillStyle = 0
    MSFlexGrid1(Index).ColSel = 1
    MSFlexGrid1(Index).RowSel = 1
  Updated = True
End Sub

Public Sub FillGrid(SelectedSlot As Recordset, Index As Integer)
Dim Row As Integer
Dim Mins As Integer
Dim Secs As Single
Dim Name As String

Updated = False
SelectedSlot.MoveFirst
Row = 1
If Currenttask = "A" Then
  Do Until SelectedSlot.EOF
    MSFlexGrid1(Index).Row = Row
    Mins = SelectedSlot!Flight_Time \ 60
    Sec = SelectedSlot!Flight_Time Mod 60
    MSFlexGrid1(Index).Col = 1
    MSFlexGrid1(Index).Text = Trim(Format(Mins, "0#"))
    MSFlexGrid1(Index).Col = 2
    MSFlexGrid1(Index).Text = Trim(Format(Sec, "0#"))
    MSFlexGrid1(Index).Col = 3
    MSFlexGrid1(Index).Text = Trim(Str(SelectedSlot!LandBonus))
    MSFlexGrid1(Index).Col = 4
    MSFlexGrid1(Index).Text = Trim(Str(SelectedSlot!Res2))
    MSFlexGrid1(Index).Col = 0
    MSFlexGrid1(Index).Text = Trim(SelectedSlot!First_Name) & " " & Trim(SelectedSlot!Last_Name)
    If SelectedSlot!Score > 0 Then
      MSFlexGrid1(Index).Col = 5
      MSFlexGrid1(Index).Text = Trim(Str(SelectedSlot!Score))
    Else
      MSFlexGrid1(Index).Col = 5
      MSFlexGrid1(Index).Text = "0"
    End If
    Row = Row + 1
    SelectedSlot.MoveNext
  Loop
ElseIf Currenttask = "B" Then
  Do Until SelectedSlot.EOF
    MSFlexGrid1(Index).Row = Row
    MSFlexGrid1(Index).Col = 1
    MSFlexGrid1(Index).Text = Trim(Format(SelectedSlot!Flight_Time, "0#"))
    MSFlexGrid1(Index).Col = 0
    MSFlexGrid1(Index).Text = Trim(SelectedSlot!First_Name) & " " & Trim(SelectedSlot!Last_Name)
    MSFlexGrid1(Index).Col = 4
    MSFlexGrid1(Index).Text = Trim(Str(SelectedSlot!Res2))
    If SelectedSlot!Score > 0 Then
      MSFlexGrid1(Index).Col = 5
      MSFlexGrid1(Index).Text = Trim(Str(SelectedSlot!Score))
    End If
    Row = Row + 1
    SelectedSlot.MoveNext
  Loop
ElseIf Currenttask = "C" Then
  Do Until SelectedSlot.EOF
    MSFlexGrid1(Index).Row = Row
    Mins = SelectedSlot!Flight_Time \ 100 'actually the seconds
    Sec = SelectedSlot!Flight_Time Mod 100 'actually the tenths
    MSFlexGrid1(Index).Col = 1
    MSFlexGrid1(Index).Text = Trim(Format(Mins, "0#"))
    MSFlexGrid1(Index).Col = 2
    MSFlexGrid1(Index).Text = Trim(Format(Sec, "0#"))
    MSFlexGrid1(Index).Col = 0
    MSFlexGrid1(Index).Text = Trim(SelectedSlot!First_Name) & " " & Trim(SelectedSlot!Last_Name)
    MSFlexGrid1(Index).Col = 4
    MSFlexGrid1(Index).Text = Trim(Str(SelectedSlot!Res2))
    If SelectedSlot!Score > 0 Then
      MSFlexGrid1(Index).Col = 5
      MSFlexGrid1(Index).Text = Format(SelectedSlot!Score, "##0.0")
      'MSFlexGrid1(Index).Text = Trim(Str(SelectedSlot!Score))
    Else
      MSFlexGrid1(Index).Col = 5
      MSFlexGrid1(Index).Text = "0"
    End If
    Row = Row + 1
    SelectedSlot.MoveNext
  Loop
End If
Updated = True
End Sub

Public Sub ClearScores(Index As Integer)
   ' Declare local variables.
    Dim Count As Integer
    Dim I As Integer
    Dim J As Integer
    
    
    ' Clear text from the first column.
    
   If Index = 0 Then
    For I = 1 To MSFlexGrid1(Index).Cols - 2
      MSFlexGrid1(Index).Col = I
      For J = 1 To MSFlexGrid1(Index).Rows - 1
        MSFlexGrid1(Index).Row = J
        MSFlexGrid1(Index).FillStyle = 1
        MSFlexGrid1(Index).Text = ""
        MSFlexGrid1(Index).FillStyle = 0
      Next J
    Next I

  ElseIf Index = 1 Then
    MSFlexGrid1(Index).Col = 1
    For Count = 1 To (MSFlexGrid1(Index).Rows - 1)
      MSFlexGrid1(Index).Row = Count
      MSFlexGrid1(Index).Text = "0"
    Next Count
    MSFlexGrid1(Index).Col = 5
    For Count = 1 To (MSFlexGrid1(Index).Rows - 1)
      MSFlexGrid1(Index).Row = Count
      MSFlexGrid1(Index).Text = "0"
    Next Count
    MSFlexGrid1(Index).FillStyle = 0
    MSFlexGrid1(Index).ColSel = 1
    MSFlexGrid1(Index).RowSel = 1
  ElseIf Index = 2 Then
    MSFlexGrid1(Index).Col = 1
    For Count = 1 To (MSFlexGrid1(Index).Rows - 1)
      MSFlexGrid1(Index).Row = Count
      MSFlexGrid1(Index).Text = "0"
    Next Count
    MSFlexGrid1(Index).Col = 2
    For Count = 1 To (MSFlexGrid1(Index).Rows - 1)
      MSFlexGrid1(Index).Row = Count
      MSFlexGrid1(Index).Text = "0"
    Next Count
    MSFlexGrid1(Index).Col = 5
    For Count = 1 To (MSFlexGrid1(Index).Rows - 1)
      MSFlexGrid1(Index).Row = Count
      MSFlexGrid1(Index).Text = "0"
    Next Count
    MSFlexGrid1(Index).FillStyle = 0
    MSFlexGrid1(Index).ColSel = 1
    MSFlexGrid1(Index).RowSel = 1
  End If
End Sub

Public Sub checkdupe(SelectedSlot As Recordset)
Dim CurrPilot As Integer
Dim DeleteSet As Recordset
Dim SQLString As String

  SelectedSlot.MoveFirst
  CurrPilot = SelectedSlot!Pilot_ID
  SelectedSlot.MoveNext
  Do Until SelectedSlot.EOF
    If SelectedSlot!Pilot_ID = CurrPilot Then
      SQLString = "SELECT DISTINCTROW Scores.Comp_ID, Scores.Round, Scores.Slot, Scores.Pilot_ID From Scores WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & ") AND (Scores.Pilot_ID = " & Str(CurrPilot) & "));"
      Set DeleteSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
      SelectedSlot.Delete
      If SelectedSlot.EOF = False Then
        SelectedSlot.MoveNext
      Else
        SelectedSlot.MoveLast
      End If
    End If
    If SelectedSlot.EOF = False Then
      CurrPilot = SelectedSlot!Pilot_ID
      SelectedSlot.MoveNext
    Else
      SelectedSlot.MoveLast
    End If
    
  Loop
End Sub


Private Sub VScroll3_Change()
Dim SQLString3 As String
  
  
  
  VScroll1.Min = 1
  Text3.Text = VScroll3.Value
  
  Currenttask = SetTask(SSTab1)
  RoundNum = Text1.Text
  Slot = Text2.Text
  If Currenttask = "B" Then
    Group = Text3.Text
  Else
    Group = 0
  End If
  SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.Task, Scores.RawScore, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & ") AND (Scores.Task='" & Currenttask & "') AND (Scores.FlightGroup = " & Group & "))ORDER BY Pilot.Last_Name;"
  Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
  If SelectedSlot.RecordCount <> 0 Then
    'checkdupe SelectedSlot
    NumPilots = SelectedSlot.RecordCount
    SelectedSlot.MoveFirst
    SetGrid NumPilots, SSTab1.Tab
    FillGrid SelectedSlot, SSTab1.Tab
  Else
    MsgBox ("No Pilots in Slot")
  End If
End Sub


Private Sub VScroll4_Change()
Dim SQLString3 As String
  
  
  
  VScroll4.Min = 0
  Text4.Text = VScroll4.Value
  
  Currenttask = SetTask(SSTab1)
  RoundNum = Text1.Text
  Slot = Text2.Text
  If Currenttask = "B" Then
    Group = Text4.Text
  Else
    Group = 0
  End If
  SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.Task, Scores.RawScore, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & ") AND (Scores.Task='" & Currenttask & "') AND (Scores.FlightGroup = " & Group & "))ORDER BY Pilot.Last_Name;"
  Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
  If SelectedSlot.RecordCount <> 0 Then
    'checkdupe SelectedSlot
    NumPilots = SelectedSlot.RecordCount
    SelectedSlot.MoveFirst
    SetGrid NumPilots, SSTab1.Tab
    FillGrid SelectedSlot, SSTab1.Tab
  Else
    MsgBox ("No Pilots in Slot")
  End If
End Sub


