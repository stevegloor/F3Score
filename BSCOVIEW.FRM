VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form ViewF3BScores 
   ClientHeight    =   6690
   ClientLeft      =   660
   ClientTop       =   1080
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
   HelpContextID   =   420
   LinkTopic       =   "Form5"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6690
   ScaleWidth      =   6825
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   6120
      Width           =   2775
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1200
      WhatsThisHelpID =   420
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      WhatsThisHelpID =   420
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "DONE"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   6120
      WhatsThisHelpID =   420
      Width           =   2655
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1200
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1200
      Width           =   495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7223
      _Version        =   393216
      TabHeight       =   529
      TabCaption(0)   =   "Task A"
      TabPicture(0)   =   "BSCOVIEW.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "List1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "List2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Task B"
      TabPicture(1)   =   "BSCOVIEW.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "List1(1)"
      Tab(1).Control(1)=   "List2(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Task C"
      TabPicture(2)   =   "BSCOVIEW.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "List1(2)"
      Tab(2).Control(1)=   "List2(2)"
      Tab(2).ControlCount=   2
      Begin VB.ListBox List2 
         Height          =   3180
         Index           =   2
         Left            =   -72480
         TabIndex        =   15
         Top             =   600
         Width           =   3735
      End
      Begin VB.ListBox List1 
         Height          =   3180
         Index           =   2
         Left            =   -74760
         TabIndex        =   14
         Top             =   600
         Width           =   2295
      End
      Begin VB.ListBox List2 
         Height          =   3180
         Index           =   1
         Left            =   -72480
         TabIndex        =   13
         Top             =   600
         Width           =   3735
      End
      Begin VB.ListBox List1 
         Height          =   3180
         Index           =   1
         Left            =   -74760
         TabIndex        =   12
         Top             =   600
         Width           =   2295
      End
      Begin VB.ListBox List2 
         Height          =   3180
         Index           =   0
         Left            =   2520
         TabIndex        =   11
         Top             =   600
         Width           =   3735
      End
      Begin VB.ListBox List1 
         Height          =   3180
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   2295
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
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   840
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
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   " View F3B Form"
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
      Height          =   615
      Left            =   960
      TabIndex        =   7
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "ViewF3BScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PilotTable As Recordset
Public ScoreTable As Recordset
Public Selected As Recordset
Public SelectedSlot As Recordset
Public SortedSlot As Recordset
Dim RoundNum As String, Slot As String
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




Private Sub Command1_Click()
  PrintForm
End Sub


Private Sub Command6_Click()
  If FromView = True Then
    FromView = False
  End If
  Unload Me
End Sub

Public Sub Form_Load()
  Dim RoundNum As String
  Dim Slot As String
  
  Text1.Text = 1
  Text2.Text = 1
  
  RoundNum = Text1.Text
  Slot = Text2.Text
  Call UpdateDisplay(RoundNum, Slot)
  
End Sub





Private Sub SSTab1_Click(PreviousTab As Integer)
  RoundNum = Text1.Text
  Slot = Text2.Text
  Call UpdateDisplay(RoundNum, Slot)
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


Private Sub VScroll1_Change()
Dim RoundNum As String
Dim Slot As String
  
  VScroll1.max = NumRounds
  VScroll1.Min = 1
  Text1.Text = VScroll1.Value
  
  
  RoundNum = Text1.Text
  Slot = Text2.Text
  Call UpdateDisplay(RoundNum, Slot)
  
End Sub

Private Sub VScroll2_Change()
  Dim SQLString3 As String
  
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
  ThisSelected.Close
  
  RoundNum = Text1.Text
  Slot = Text2.Text
  Call UpdateDisplay(RoundNum, Slot)
  
  End Sub




Private Sub UpdateDisplay(ByVal RoundNum, ByVal Slot)
Dim SQLString As String
Dim SelectedSlot As Recordset
Dim Mins As Integer
Dim Sec As Single
Dim Line As String
Dim T As Integer

  Currenttask = SetTask(SSTab1)
  T = SSTab1.Tab
  List1(T).Clear
  List2(T).Clear
  SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.Task, Scores.RawScore, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & ") AND (Scores.Task='" & Currenttask & "'))ORDER BY Scores.Score DESC;"
  
  Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
  If SelectedSlot.RecordCount <> 0 Then
    SelectedSlot.MoveFirst
    If Currenttask = "A" Then
      Line = "    Pilot Name"
      List1(T).AddItem Line
      Line = "Mins" & "   " & "Secs" & "  " & "Land" & "    " & "Pen" & "     " & "Score"
      List2(T).AddItem Line
      Do Until SelectedSlot.EOF
        If SelectedSlot!Flight_Time <> 0 Then
           Mins = SelectedSlot!Flight_Time \ 60
           Sec = SelectedSlot!Flight_Time Mod 60
        End If
        Line = SelectedSlot!First_Name & " " & SelectedSlot!Last_Name
        List1(T).AddItem Line
        Line = "  " & Format(Mins, "0#") & "      " & Format(Sec, "0#") & "     " & Format(SelectedSlot!LandBonus, "0##") & "      " & Format(SelectedSlot!Res2, "0##") & "      " & Format(SelectedSlot!Score, "####.0")
        List2(T).AddItem Line
        SelectedSlot.MoveNext
        Mins = 0
        Sec = 0
      Loop
    ElseIf Currenttask = "B" Then
      Line = "    Pilot Name"
      List1(T).AddItem Line
      Line = "Laps" & "     " & "Score"
      List2(T).AddItem Line
      Do Until SelectedSlot.EOF
        Line = SelectedSlot!First_Name & " " & SelectedSlot!Last_Name
        List1(T).AddItem Line
        Line = "  " & Format(SelectedSlot!Flight_Time, "0#") & "      " & Format(SelectedSlot!Score, "0###.0")
        List2(T).AddItem Line
        SelectedSlot.MoveNext
        Mins = 0
        Sec = 0
      Loop
    ElseIf Currenttask = "C" Then
      Line = "    Pilot Name"
      List1(T).AddItem Line
      Line = "   Time" & "         " & "Score"
      List2(T).AddItem Line
      Do Until SelectedSlot.EOF
        If SelectedSlot!Flight_Time <> 0 Then
           Mins = SelectedSlot!Flight_Time \ 100
           Sec = SelectedSlot!Flight_Time Mod 100
        End If
        Line = SelectedSlot!First_Name & " " & SelectedSlot!Last_Name
        List1(T).AddItem Line
        Line = "  " & Format(Mins, "0#") & " : " & Format(Sec, "0#") & "       " & Format(SelectedSlot!Score, "0###.0")
        List2(T).AddItem Line
        SelectedSlot.MoveNext
        Mins = 0
        Sec = 0
      Loop
    End If
  Else
    MsgBox ("No Pilots in Slot")
  End If
End Sub
