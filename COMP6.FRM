VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form EnterF3JScores 
   ClientHeight    =   6090
   ClientLeft      =   660
   ClientTop       =   705
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   310
   LinkTopic       =   "Form5"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6090
   ScaleWidth      =   7200
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pilots in Selected Slot"
      Height          =   3735
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      WhatsThisHelpID =   310
      Width           =   6975
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2895
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5106
         _Version        =   393216
      End
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   495
      Left            =   3720
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
      Left            =   4320
      TabIndex        =   3
      Top             =   5520
      WhatsThisHelpID =   310
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   5520
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
      Left            =   3240
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
   Begin VB.Label Label3 
      Caption         =   "S Pen - Safety Penalty is subtracted from final score"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "F Pen - Flight penalty affects normalised score"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   720
      Width           =   2655
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
      Left            =   3360
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
      Caption         =   " Flight Time Input Form"
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
      Left            =   960
      TabIndex        =   8
      Top             =   120
      WhatsThisHelpID =   310
      Width           =   4215
   End
End
Attribute VB_Name = "EnterF3JScores"
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
Dim ThisNumSlots As Integer



Private Sub Command3_Click()
  If CurrentContestType = "F3J" Then
    Call F3JScore(RoundNum, Slot, NumPilots, SelectedSlot)
  ElseIf CurrentContestType = "F3JFO" Then
    Call F3JScore(RoundNum, Slot, NumPilots, SelectedSlot)
  ElseIf CurrentContestType = "AustOpen" Then
    Call AusDurScore(RoundNum, Slot, NumPilots, SelectedSlot)
  End If
  
  VScroll1.Enabled = True
  VScroll2.Enabled = True
  Command6.Enabled = True
End Sub



Private Sub Command6_Click()
  If FromView = True Then
    FromView = False
  End If
  If ViewScores = True Then
    ViewScores = False
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
  
  Text1.Text = 1
  Text2.Text = 1
  Dim SQLString4 As String
  Dim ThisSelected As Recordset
  Dim ThisRound As String
  Dim Count As Integer
  Dim Col As Integer
  Dim Row As Integer
  
  ThisRound = Text1.Text
  If ViewScores = True Then
    Command3.Visible = False
  End If
  Count = 1
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
  
  SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.RawScore, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & "))ORDER BY Pilot.Last_Name;"
  Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
  If SelectedSlot.RecordCount <> 0 Then
    'checkdupe SelectedSlot
    NumPilots = SelectedSlot.RecordCount
    SetGrid NumPilots
    SelectedSlot.MoveFirst
    FillGrid SelectedSlot
  Else
    MsgBox ("No Pilots in Slot")
    Unload Me
    Exit Sub
  End If
  
  Command3.Caption = "Store Scores"
  Command6.Caption = "Done"
    
 End Sub


Public Sub F3JScore(RoundNum, Slot, NumPilots, SelectedSlot As Recordset)
Dim String1 As String
Dim FirstScore
Dim PilotScore
Dim UpdateTable As Recordset
Dim Count As Integer
Dim SQLString As String
Dim FlightPenalty(20) As Single
Dim LandBonus As Boolean
Dim ManualPenalty(20) As Single
Dim Row As Integer
Dim FlightTime(20) As Single
Dim Column As Integer
Dim MaxScore As Single
Dim NormScore(20) As Single
Dim ValLandBonus(20) As Single
Dim Rawscore(20) As Single
Dim Tagged As Boolean
Dim Refly As Boolean
Dim Mins As Integer
Dim Secs As Single
Dim Pilots As Integer
  
  Pilots = NumPilots
  Row = 0
  MSFlexGrid1.Row = Row
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
    MSFlexGrid1.Row = Row
    LandBonus = True
    MSFlexGrid1.Col = 1
    FlightTime(Row) = Val(MSFlexGrid1.Text) * 60
    MSFlexGrid1.Col = 2
    FlightTime(Row) = FlightTime(Row) + Val(Trim(MSFlexGrid1.Text))
    MSFlexGrid1.Col = 4
    If FlightTime(Row) > (max * 60) And FlightTime(Row) < ((max + 1) * 60) Then
      FlightPenalty(Row) = FlightPenalty(Row) + (30 + Round((FlightTime(Row) - (max * 60)), 1))
      'FlightTime(Row) = (max * 60)
      LandBonus = False
    ElseIf FlightTime(Row) >= ((max + 1) * 60) Then
      'SelectedSlot!Flight_Time = 0
      FlightPenalty(Row) = FlightPenalty(Row) + FlightTime(Row)
      LandBonus = False
    End If
    MSFlexGrid1.Col = 5
    ManualPenalty(Row) = Val(MSFlexGrid1.Text)
    MSFlexGrid1.Col = 3
    If Val(MSFlexGrid1.Text) < 30 Then
      If LandBonus = True Then
        MSFlexGrid1.Col = 3
        Select Case Val(MSFlexGrid1.Text)
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
          ValLandBonus(Row) = 35
        Case 15
          ValLandBonus(Row) = 30
        Case Else
          ValLandBonus(Row) = 0
        End Select
        
      Else
        ValLandBonus(Row) = 0
        
      End If
    Else
      MSFlexGrid1.Col = 3
      ValLandBonus(Row) = Val(MSFlexGrid1.Text)
    End If
    Rawscore(Row) = FlightTime(Row) + ValLandBonus(Row) - FlightPenalty(Row)
    If Rawscore(Row) > MaxScore Then
      MaxScore = Rawscore(Row)
    End If
  Loop
  Rem normalise scores
  If MaxScore <> 0 Then
    For Row = 1 To NumPilots
      NormScore(Row) = Round((Rawscore(Row) / MaxScore) * 1000, 4)
    Next Row
  End If
  Row = 1
  
  If Refly = True Then
    If MsgBox("This slot has already been done" & Chr(13) & "Do You want to rescore this slot", 19) <> vbYes Then
      MsgBox ("Scoring Cancelled")
    Else
      Tagged = False
      MSFlexGrid1.Col = 7
      For Row = 1 To NumPilots
        MSFlexGrid1.Row = Row
        If MSFlexGrid1.Text = "*" Then
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
            SelectedSlot!Score = Round(NormScore(Row), 1)
            SelectedSlot!Rawscore = Round(Rawscore(Row), 1)
            SelectedSlot!Flight_Time = FlightTime(Row)
            SelectedSlot!LandBonus = ValLandBonus(Row)
            SelectedSlot!Penalty = FlightPenalty(Row)
            SelectedSlot!Res2 = ManualPenalty(Row)
            SelectedSlot.Update
            Row = Row + 1
            If Row < NumPilots Then
              MSFlexGrid1.Row = Row
            End If
            SelectedSlot.MoveNext
          Loop
          FillGrid SelectedSlot
          MsgBox ("Slot Rescored")
        End If
      Else
        SelectedSlot.MoveFirst
        Row = 1
        MSFlexGrid1.Row = Row
        Do Until SelectedSlot.EOF
          MSFlexGrid1.Col = 7
          MSFlexGrid1.Row = Row
          If MSFlexGrid1.Text <> "*" Then
            If SelectedSlot!Score < Round(NormScore(Row), 1) Then
              SelectedSlot.Edit
              SelectedSlot!Score = Round(NormScore(Row), 1)
              SelectedSlot!Rawscore = Round(Rawscore(Row), 1)
              SelectedSlot!Flight_Time = FlightTime(Row)
              SelectedSlot!LandBonus = ValLandBonus(Row)
              SelectedSlot!Penalty = FlightPenalty(Row)
              SelectedSlot!Res2 = ManualPenalty(Row)
              SelectedSlot.Update
            End If
          Else
            SelectedSlot.Edit
            SelectedSlot!Score = Round(NormScore(Row), 1)
            SelectedSlot!Rawscore = Round(Rawscore(Row), 1)
            SelectedSlot!Flight_Time = FlightTime(Row)
            SelectedSlot!LandBonus = ValLandBonus(Row)
            SelectedSlot!Penalty = FlightPenalty(Row)
            SelectedSlot!Res2 = ManualPenalty(Row)
            SelectedSlot.Update
          End If
          Row = Row + 1
          If Row < NumPilots Then
            MSFlexGrid1.Row = Row
          End If
          SelectedSlot.MoveNext
        Loop
        FillGrid SelectedSlot
        SelectedSlot.MoveFirst
        MsgBox ("Slot Rescored")
      End If
    End If
  Else
    Row = 1
    SelectedSlot.MoveFirst
    Do Until SelectedSlot.EOF
      SelectedSlot.Edit
      SelectedSlot!Score = Round(NormScore(Row), 1)
      SelectedSlot!Rawscore = Round(Rawscore(Row), 1)
      SelectedSlot!Flight_Time = FlightTime(Row)
      SelectedSlot!LandBonus = ValLandBonus(Row)
      SelectedSlot!Penalty = FlightPenalty(Row)
      SelectedSlot!Res2 = ManualPenalty(Row)
      SelectedSlot.Update
      Row = Row + 1
      SelectedSlot.MoveNext
    Loop
    FillGrid SelectedSlot
    MsgBox ("Slot Scored")
  End If 'refly
  
End Sub


Private Sub msflexgrid1_DblClick()
  Dim R As Integer
  
  MSFlexGrid1.Col = 7
  If MSFlexGrid1.Text = "" Then
    MSFlexGrid1.Text = "*"
    'MSFlexGrid1.Col = 4
    'For R = 1 To (MSFlexGrid1.Rows - 1)
    '  MSFlexGrid1.Row = R
    '  MSFlexGrid1.Text = ""
    'Next R
    If MsgBox("Do you want all the scores cleared", 19) = vbYes Then
      ClearScores
    End If
  Else
    MSFlexGrid1.Text = ""
  End If
End Sub

Private Sub msflexgrid1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Text3.Text <> "" Then
      MSFlexGrid1.Text = Text3.Text
    End If
  ElseIf KeyAscii = 27 Then
    Text3.Text = ""
  Else
    Text3.Text = Text3.Text & Chr(KeyAscii)
    Srow = MSFlexGrid1.Row
    Scol = MSFlexGrid1.Col
  End If
End Sub


Private Sub msflexgrid1_SelChange()

Dim Pcol As Integer
Dim Prow As Integer

If Updated = True Then
  If Text3.Text <> "" Then
    Pcol = MSFlexGrid1.Col
    Prow = MSFlexGrid1.Row
    MSFlexGrid1.Row = Srow
    MSFlexGrid1.Col = Scol
    If MSFlexGrid1.Col = 2 Then
      If Val(Text3.Text) >= 60 Then
        MsgBox ("Entry must be less that 60")
        Text3.Text = ""
        MSFlexGrid1.Text = ""
        MSFlexGrid1.Row = Prow
        MSFlexGrid1.Col = Pcol
        Exit Sub
      End If
    End If
    If MSFlexGrid1.Col = 3 Then
      If Val(Text3.Text) > 100 Then
        MsgBox ("You cannot enter more than 100 points")
        Text3.Text = ""
        MSFlexGrid1.Row = Prow
        MSFlexGrid1.Col = Pcol
      End If
    End If
    MSFlexGrid1.Text = Text3.Text
    Text3.Text = ""
    MSFlexGrid1.Row = Prow
    MSFlexGrid1.Col = Pcol
  End If
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

Private Sub Text4_Change()
  If Val(Text4.Text) > 60 Then
    MsgBox ("Please re-enter seconds")
    Text4.Text = ""
  End If
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
  
  
  RoundNum = Text1.Text
  Slot = Text2.Text
  
  
  SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.RawScore, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & "))ORDER BY Pilot.Last_Name;"
  Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
  If SelectedSlot.RecordCount <> 0 Then
    'checkdupe SelectedSlot
    NumPilots = SelectedSlot.RecordCount
    SelectedSlot.MoveFirst
    SetGrid NumPilots
    FillGrid SelectedSlot
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
  
  SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.RawScore, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & "))ORDER BY Pilot.Last_Name;"
  Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
  If SelectedSlot.RecordCount <> 0 Then
    'checkdupe SelectedSlot
    NumPilots = SelectedSlot.RecordCount
    SelectedSlot.MoveFirst
    SetGrid NumPilots
    FillGrid SelectedSlot
        
  Else
   MsgBox ("No Pilots in Slot")
  End If
  End Sub




Sub SetGrid(NumPilots As Integer)
  Dim Col2Fill
  Dim Row2Fill
  Dim NL As String
  Dim GridHgt
  Dim GridWid
  Dim I
  Updated = False
  
  GridHgt = MSFlexGrid1.Height
  GridWid = MSFlexGrid1.Width
  If ViewScores = False Then
    MSFlexGrid1.Cols = 8
  Else
    MSFlexGrid1.Cols = 7
  End If
  MSFlexGrid1.Colwidth(0) = GridWid / 3
  If NumPilots > 7 Then
    MSFlexGrid1.ScrollBars = 2
    GridWid = GridWid - MSFlexGrid1.Colwidth(0) - 120
  Else
    MSFlexGrid1.ScrollBars = 0
    GridWid = GridWid - MSFlexGrid1.Colwidth(0) - 65
  End If
  MSFlexGrid1.Colwidth(1) = (GridWid / 8)
  MSFlexGrid1.Colwidth(2) = (GridWid / 7)
  MSFlexGrid1.Colwidth(3) = (GridWid / 8)
  MSFlexGrid1.Colwidth(4) = (GridWid / 7)
  MSFlexGrid1.Colwidth(5) = (GridWid / 7)
  MSFlexGrid1.Colwidth(6) = (GridWid / 5)
  If ViewScores = False Then
    MSFlexGrid1.Colwidth(7) = (GridWid / 9)
  End If
    
    
  MSFlexGrid1.Rows = (NumPilots + 1)
  If NumPilots > 7 Then
    MSFlexGrid1.ScrollBars = 2
  Else
    MSFlexGrid1.ScrollBars = 0
  End If
  MSFlexGrid1.RowHeight(0) = GridHgt / 7
  For I = 1 To NumPilots
    MSFlexGrid1.RowHeight(I) = GridHgt / 9
  Next I
 ' Clear the contents of the grid.
  Cleargrid
  'Initialize the line break variable.
  NL = Chr(10) & Chr(13)
  ' First row = current row.
  MSFlexGrid1.Row = 0
  MSFlexGrid1.Col = 0
  MSFlexGrid1.FixedAlignment(0) = 0
  MSFlexGrid1.Text = "Name"
  
  MSFlexGrid1.Row = 0
  'Put column headings in the first row.
  MSFlexGrid1.Col = 1
  MSFlexGrid1.FixedAlignment(1) = 2
  MSFlexGrid1.Text = "Mins"
  MSFlexGrid1.Col = 2
  MSFlexGrid1.FixedAlignment(2) = 2
  MSFlexGrid1.Text = "Sec"
  MSFlexGrid1.Col = 3
  MSFlexGrid1.FixedAlignment(3) = 2
  MSFlexGrid1.Text = "Land"
  MSFlexGrid1.Col = 4
  MSFlexGrid1.FixedAlignment(4) = 2
  MSFlexGrid1.Text = "F Pen"
  MSFlexGrid1.Col = 5
  MSFlexGrid1.FixedAlignment(5) = 2
  MSFlexGrid1.Text = "S Pen"
  MSFlexGrid1.Col = 6
  MSFlexGrid1.FixedAlignment(5) = 2
  MSFlexGrid1.Text = "Score"
  If ViewScores = False Then
    MSFlexGrid1.Col = 7
    MSFlexGrid1.FixedAlignment(6) = 2
    MSFlexGrid1.Text = "Tag"
  End If
  For I = 1 To 5
    MSFlexGrid1.ColAlignment(I) = 2
  Next I
  
  Updated = True
  
End Sub

Sub Cleargrid()
   ' Declare local variables.
    Dim Count As Integer
    Dim I As Integer
    Dim J As Integer
    
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
    ' Clear text from the nonfixed cells.
    MSFlexGrid1.Clear
    For I = 1 To MSFlexGrid1.Cols - 2
      MSFlexGrid1.Col = I
      For J = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = J
        MSFlexGrid1.FillStyle = 1
        MSFlexGrid1.Text = ""
        MSFlexGrid1.FillStyle = 0
      Next J
    Next I
    MSFlexGrid1.ColSel = 1
    MSFlexGrid1.RowSel = 1
End Sub

Public Sub FillGrid(SelectedSlot As Recordset)
Dim Row As Integer
Dim Mins As Integer
Dim Secs As Single
Dim Name As String
Dim ActualPenalty As Integer

Updated = False
SelectedSlot.MoveFirst
Row = 1
Do Until SelectedSlot.EOF
  ActualPenalty = 0
  MSFlexGrid1.Row = Row
  Mins = SelectedSlot!Flight_Time \ 60
  'Sec = SelectedSlot!Flight_Time - ((SelectedSlot!Flight_Time \ 60) * 60)
  Sec = SelectedSlot!Flight_Time - ((SelectedSlot!Flight_Time \ 60) * 60)
  MSFlexGrid1.Col = 1
  MSFlexGrid1.Text = Trim(Format(Mins, "0#"))
  MSFlexGrid1.Col = 2
  MSFlexGrid1.Text = Trim(Format(Sec, "0#.0"))
  MSFlexGrid1.Col = 3
  MSFlexGrid1.Text = Trim(Str(SelectedSlot!LandBonus))
  MSFlexGrid1.Col = 4
  If CurrentContestType = "F3J" Then
    If SelectedSlot!Penalty > 0 Then
      ActualPenalty = SelectedSlot!Penalty - (SelectedSlot!Flight_Time - (max * 60))
      MSFlexGrid1.Text = Trim(Str(ActualPenalty))
    Else
      MSFlexGrid1.Text = Trim(Str(ActualPenalty))
    End If
  End If
  
  MSFlexGrid1.Col = 5
  MSFlexGrid1.Text = Trim(Str(SelectedSlot!Res2))
  MSFlexGrid1.Col = 0
  MSFlexGrid1.Text = Trim(SelectedSlot!First_Name) & " " & Trim(SelectedSlot!Last_Name)
  If SelectedSlot!Score > 0 Then
    MSFlexGrid1.Col = 6
    MSFlexGrid1.Text = Format(SelectedSlot!Score, "##0.0")
  Else
    MSFlexGrid1.Col = 6
    MSFlexGrid1.Text = "0"
  End If
  Row = Row + 1
  SelectedSlot.MoveNext
Loop
Updated = True
End Sub

Public Sub ClearScores()
   ' Declare local variables.
    Dim Count As Integer
    Dim I As Integer
    Dim J As Integer
    ' Clear text from the first column.
    
   
    ' Clear text from the nonfixed cells.
    
    For I = 1 To MSFlexGrid1.Cols - 2
      MSFlexGrid1.Col = I
      For J = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = J
        MSFlexGrid1.FillStyle = 1
        MSFlexGrid1.Text = ""
        MSFlexGrid1.FillStyle = 0
      Next J
    Next I
    MSFlexGrid1.ColSel = 1
    MSFlexGrid1.RowSel = 1
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

Public Sub AusDurScore(RoundNum, Slot, NumPilots, SelectedSlot As Recordset)
Dim String1 As String
Dim FirstScore
Dim PilotScore
Dim UpdateTable As Recordset
Dim Count As Integer
Dim SQLString As String
Dim FlightPenalty(20) As Single
Dim LandBonus As Boolean
Dim ManualPenalty(20) As Single
Dim Row As Integer
Dim FlightTime(20) As Single
Dim Column As Integer
Dim MaxScore As Single
Dim NormScore(20) As Single
Dim ValLandBonus(20) As Single
Dim Rawscore(20) As Single
Dim Tagged As Boolean
Dim Refly As Boolean
Dim Mins As Integer
Dim Secs As Single
Dim Pilots As Integer
Dim addpenalty As Boolean

  Pilots = NumPilots
  Row = 0
  MSFlexGrid1.Row = Row
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

    MSFlexGrid1.Row = Row
    LandBonus = True
    MSFlexGrid1.Col = 1
    FlightTime(Row) = Val(MSFlexGrid1.Text) * 60
    MSFlexGrid1.Col = 2
    FlightTime(Row) = FlightTime(Row) + Val(Trim(MSFlexGrid1.Text))
    MSFlexGrid1.Col = 5
    ManualPenalty(Row) = Val(MSFlexGrid1.Text)
    If FlightTime(Row) > (max * 60) And FlightTime(Row) < ((max + 0.5) * 60) Then
      FlightPenalty(Row) = FlightPenalty(Row) + ((FlightTime(Row) - (max * 60)) * 2)
      'FlightTime(Row) = (max * 60)
      LandBonus = True
    ElseIf FlightTime(Row) >= ((max + 0.5) * 60) Then
      FlightPenalty(Row) = FlightPenalty(Row) + ((FlightTime(Row) - (max * 60)) * 2)
      LandBonus = False
    End If
    MSFlexGrid1.Col = 3
    If Val(MSFlexGrid1.Text) < 30 Then
      If LandBonus = True Then
        MSFlexGrid1.Col = 3
        Select Case Val(MSFlexGrid1.Text)
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
          ValLandBonus(Row) = 35
        Case 15
          ValLandBonus(Row) = 30
        Case Else
          ValLandBonus(Row) = 0
        End Select
        
      Else
        ValLandBonus(Row) = 0
        
      End If
    Else
      MSFlexGrid1.Col = 3
      ValLandBonus(Row) = Val(MSFlexGrid1.Text)
    End If
    Rawscore(Row) = FlightTime(Row) + ValLandBonus(Row) - FlightPenalty(Row)
    If Rawscore(Row) > MaxScore Then
      MaxScore = Rawscore(Row)
    End If
  Loop
  Rem normalise scores
  
  If MaxScore <> 0 Then
    For Row = 1 To NumPilots
      NormScore(Row) = Round((Rawscore(Row) / MaxScore) * 1000, 4)
    Next Row
  End If
  Row = 1
  
  If Refly = True Then
    If MsgBox("This slot has already been done" & Chr(13) & "Do You want to rescore this slot", 19) <> vbYes Then
      MsgBox ("Scoring Cancelled")
    Else
      Tagged = False
      MSFlexGrid1.Col = 7
      For Row = 1 To NumPilots
        MSFlexGrid1.Row = Row
        If MSFlexGrid1.Text = "*" Then
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
            SelectedSlot!Score = Round(NormScore(Row), 1)
            SelectedSlot!Rawscore = Round(Rawscore(Row), 1)
            SelectedSlot!Flight_Time = FlightTime(Row)
            SelectedSlot!LandBonus = ValLandBonus(Row)
            SelectedSlot!Penalty = FlightPenalty(Row)
            SelectedSlot!Res2 = ManualPenalty(Row)
            SelectedSlot.Update
            Row = Row + 1
            If Row < NumPilots Then
              MSFlexGrid1.Row = Row
            End If
            SelectedSlot.MoveNext
          Loop
          FillGrid SelectedSlot
          MsgBox ("Slot Rescored")
        End If
      Else
        SelectedSlot.MoveFirst
        Row = 1
        MSFlexGrid1.Row = Row
        Do Until SelectedSlot.EOF
          MSFlexGrid1.Col = 7
          MSFlexGrid1.Row = Row
          If MSFlexGrid1.Text <> "*" Then
            If SelectedSlot!Score < Round(NormScore(Row), 1) Then
              SelectedSlot.Edit
              SelectedSlot!Score = Round(NormScore(Row), 1)
              SelectedSlot!Rawscore = Round(Rawscore(Row), 1)
              SelectedSlot!Flight_Time = FlightTime(Row)
              SelectedSlot!LandBonus = ValLandBonus(Row)
              SelectedSlot!Penalty = FlightPenalty(Row)
              SelectedSlot!Res2 = ManualPenalty(Row)
              SelectedSlot.Update
            End If
          Else
            SelectedSlot.Edit
            SelectedSlot!Score = Round(NormScore(Row), 1)
            SelectedSlot!Rawscore = Round(Rawscore(Row), 1)
            SelectedSlot!Flight_Time = FlightTime(Row)
            SelectedSlot!LandBonus = ValLandBonus(Row)
            SelectedSlot!Penalty = FlightPenalty(Row)
            SelectedSlot!Res2 = ManualPenalty(Row)
            SelectedSlot.Update
          End If
          Row = Row + 1
          If Row < NumPilots Then
            MSFlexGrid1.Row = Row
          End If
          SelectedSlot.MoveNext
        Loop
        FillGrid SelectedSlot
        SelectedSlot.MoveFirst
        MsgBox ("Slot Rescored")
      End If
    End If
  Else
    Row = 1
    SelectedSlot.MoveFirst
    Do Until SelectedSlot.EOF
      SelectedSlot.Edit
      SelectedSlot!Score = Round(NormScore(Row), 1)
      SelectedSlot!Rawscore = Round(Rawscore(Row), 1)
      SelectedSlot!Flight_Time = FlightTime(Row)
      SelectedSlot!LandBonus = ValLandBonus(Row)
      SelectedSlot!Penalty = FlightPenalty(Row)
      SelectedSlot!Res2 = ManualPenalty(Row)
      SelectedSlot.Update
      Row = Row + 1
      SelectedSlot.MoveNext
    Loop
    FillGrid SelectedSlot
    MsgBox ("Slot Scored")
  End If

End Sub
