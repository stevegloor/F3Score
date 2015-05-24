VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AllocateF3B 
   Caption         =   "Slot Allocation Form"
   ClientHeight    =   6465
   ClientLeft      =   480
   ClientTop       =   750
   ClientWidth     =   8715
   HelpContextID   =   350
   LinkTopic       =   "Form4"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   26.938
   ScaleMode       =   4  'Character
   ScaleWidth      =   72.625
   Begin VB.CommandButton Command9 
      Caption         =   "Change Task C"
      Height          =   375
      Left            =   7200
      TabIndex        =   31
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Change Task B"
      Height          =   375
      Left            =   5520
      TabIndex        =   30
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Change Task A"
      Height          =   375
      Left            =   3840
      TabIndex        =   29
      Top             =   5520
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6588
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Task A"
      TabPicture(0)   =   "FORM4b.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "List1(1)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Task B"
      TabPicture(1)   =   "FORM4b.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "List1(2)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Task C"
      TabPicture(2)   =   "FORM4b.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "List1(3)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.ListBox List1 
         Height          =   3180
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   8055
      End
      Begin VB.ListBox List1 
         Height          =   3180
         Index           =   2
         Left            =   -74880
         TabIndex        =   26
         Top             =   360
         Width           =   8055
      End
      Begin VB.ListBox List1 
         Height          =   3180
         Index           =   1
         Left            =   -74880
         TabIndex        =   25
         Top             =   360
         Width           =   8055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Do not try to edit these"
      Height          =   855
      Left            =   3840
      TabIndex        =   12
      Top             =   4320
      Width           =   4575
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   4080
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Number of Pilots"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Number of times pilots could fly against each other - Smaller is better "
         Height          =   615
         Left            =   1800
         TabIndex        =   13
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Change these to suit your Contest"
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   3615
      Begin VB.CheckBox Check1 
         Caption         =   "Check to make Task C one slot"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   28
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   3
         Left            =   2520
         TabIndex        =   21
         Text            =   "Text3"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   2
         Left            =   2520
         TabIndex        =   20
         Text            =   "Text3"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   3
         Left            =   1080
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   8
         Text            =   "Text3"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Task C"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Task B"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Task A"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Number of Slots"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   " Number of Pilots in each Slot"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Finish Allocation Manually"
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   5880
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   2160
      ScaleHeight     =   195
      ScaleWidth      =   6075
      TabIndex        =   5
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   5160
      WhatsThisHelpID =   350
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Allocate Slots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   5160
      WhatsThisHelpID =   350
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Accept Allocation and Set Up Scores In Database"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   5880
      WhatsThisHelpID =   350
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Allocations"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   5160
      WhatsThisHelpID =   350
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Allocation Progress"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      WhatsThisHelpID =   350
      Width           =   1935
   End
End
Attribute VB_Name = "AllocateF3B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OK As Boolean
Public NumberOfPilots As Integer
Public allocated As Boolean
Public Maxtasks As Integer

Private Sub Command1_Click()
  Dim I As Integer
  
  Command4.Enabled = True
  'Command2.Enabled = True
  Command3.Enabled = True
  For I = 1 To 3
    List1(I).Clear
  Next I
  Picture1.Cls
  ReDim alloc(4, 0, 0, 0)
End Sub


Public Sub MainAllocate()
  
  If CurrentContestType = "F3B" Then
    AllocateSlots
  
  
  Else
    MsgBox ("Contest is not yet supported")
  End If
End Sub

Public Sub AllocateSlots()
  'On Error GoTo errhandler
  Dim Flight() As Integer
  Dim Freq() As Single
  Dim Team() As Integer
  Dim dupes() As Integer
  Dim SQLString(3) As String
  Dim NumPilots As Integer
  Dim MaxPilotsInSlot As Integer
  Dim I As Integer
  Dim J As Integer
  Dim K As Integer
  Dim Task As Integer
  Dim P As Integer
  
  Dim RoundNum As Integer
  Dim OldString As Integer
  Dim Line1 As String
  Dim Selected As Recordset
  Dim MaxSlots As Integer
  Dim maxnum As Integer
  Dim OldValue As Integer
  Dim RoundTries As Integer
  Dim Message As String
  Dim Response As Boolean
  Dim SQLString1 As String
  'Dim lngValue1    As Long
  'Dim lngValue2    As Long
  'Dim dblData()    As Double
  'Dim dblTestData  As Double
  'Dim cRndGen      As clsRndGen
  'Set cRndGen = New clsRndGen
  
  ReDim F3BNumSlots(3)
  Picture1.ForeColor = RGB(0, 0, 255)     'use blue bar
  OK = True
  RoundTries = 0
  
  Text4.Text = ""
  MaxSlots = 0
  SQLString(2) = "SELECT * FROM List Where ((List.Comp_ID =" & Str(CurrentContest) & "));"
  SQLString(1) = "SELECT * FROM List Where ((List.Comp_ID =" & Str(CurrentContest) & ")) ORDER BY List.Rnd1 ASC;"
  SQLString(3) = "SELECT * FROM List Where ((List.Comp_ID =" & Str(CurrentContest) & ")) ORDER BY List.FreqUsed ASC,List.Rnd1 ASC;"
  
  Set Selected = F3JDb.OpenRecordset(SQLString(2), dbOpenDynaset)
  
  NumPilots = Val(Text1.Text)
  
  'Find out the maximum pilot ID
  Selected.MoveFirst
  Do Until Selected.EOF
    If Selected!Pilot_ID > maxnum Then
      maxnum = Selected!Pilot_ID
    End If
    Selected.MoveNext
  Loop
  'Dimension the arrays to the right sizes
  For I = 1 To 3
    F3BNumSlots(I) = Val(Text2(I).Text)
  Next I
  
  MaxPilotsInSlot = Val(Text3(1).Text)
  ReDim Flight(4, maxnum, maxnum)
  For I = 1 To 4
    For J = 1 To maxnum
      For K = 1 To maxnum
        Flight(I, J, K) = 0
      Next K
    Next J
  Next I
  ReDim alloc(4, NumRounds, F3BNumSlots(1) + 10, MaxPilotsInSlot + 10)
 
  ReDim Freq(maxnum)
  ReDim Team(maxnum)
  ReDim dupes(maxnum)
  'fill the arrays
  Selected.MoveFirst
  Do Until Selected.EOF
    Freq(Selected!Pilot_ID) = Selected!FreqUsed
    Team(Selected!Pilot_ID) = Selected!Team_ID
    Selected.MoveNext
  Loop
  
  'allocat pilots to rounds
  If Check1.Value = 0 Then
    Maxtasks = 3
  Else
    F3BNumSlots(3) = 1
    ReDim alloc(4, NumRounds, F3BNumSlots(1) + 10, Selected.RecordCount + 20)
    Maxtasks = 2
    For RoundNum = 1 To NumRounds
      P = 1
      Selected.Close
      Set Selected = F3JDb.OpenRecordset(SQLString(2), dbOpenDynaset)
      Selected.MoveFirst
      Do Until Selected.EOF
        Randomize (CDbl(Now()) + Timer)
        Selected.Edit
        Selected!Rnd1 = Int((10000 * Rnd) + 1)
        Selected.Update
        Selected.MoveNext
      Loop
      Selected.Close
      Set Selected = F3JDb.OpenRecordset(SQLString(1), dbOpenDynaset)
      Selected.MoveFirst
      Do Until Selected.EOF
        alloc(3, RoundNum, 1, P) = Selected!Pilot_ID
        Selected.MoveNext
        P = P + 1
      Loop
    Next RoundNum
  End If
  For Task = 1 To Maxtasks
    'ReDim Flight(maxnum, maxnum)
    'F3BNumSlots(Task) = Val(Text2(Task).Text)
    'OldValue = F3BNumSlots(task)
    MaxPilotsInSlot = Val(Text3(Task).Text)
    RoundNum = 1
    OK = True
    RoundTries = 0
    maxtimes = F3BNumSlots(Task) / 2
    Do Until RoundNum > NumRounds
    Selected.Close
    Set Selected = F3JDb.OpenRecordset(SQLString(2), dbOpenDynaset)
    Selected.MoveFirst
    Do Until Selected.EOF
      Randomize (CDbl(Now()) + Timer)
      Selected.Edit
      Selected!Rnd1 = Int((10000 * Rnd) + 1)
      Selected.Update
      Selected.MoveNext
    Loop
    Selected.Close
    'pull out names in random order
    Set Selected = F3JDb.OpenRecordset(SQLString(1), dbOpenDynaset)
    NumPilots = Selected.RecordCount
    'If MaxSlots <= NumSlots Then
    '  MaxSlots = NumSlots
    'End If
    'NumSlots = OldValue
    Selected.MoveFirst
    Call RoundSlotAlloc2(Task, Selected, Flight(), Freq(), Team(), RoundNum, MaxPilotsInSlot) 'Call slot allocation procedure
    'If the round has been drawn correctly
    If OK = True Then
      updateprogress Picture1, (RoundNum / NumRounds) * 100
      'updating flight matrix with pilot information
         'after a successful allocation
      P = 1
      J = 1
      For I = 1 To F3BNumSlots(Task)
        For J = 1 To MaxPilotsInSlot
          P = 1
          Currentpilot = alloc(Task, RoundNum, I, J) 'Take the first pilot
          Do Until alloc(Task, RoundNum, I, P) = 0 'update
            If Currentpilot <> alloc(Task, RoundNum, I, P) Then
              Flight(Task, Currentpilot, alloc(Task, RoundNum, I, P)) = Flight(Task, Currentpilot, alloc(Task, RoundNum, I, P)) + 1
              
            End If
            If P <> MaxPilotsInSlot Then
              P = P + 1
            Else
              Exit Do
            End If
          Loop
        Next J
      P = 1
      Next I
      
      RoundNum = RoundNum + 1
      RoundsCorrect = RoundNum
    ElseIf RoundTries > (NumPilots * 5) And RoundTries < (NumPilots * 10) Then
      maxtimes = maxtimes + 1
      Text4.Text = Str(maxtimes)
      RoundTries = RoundTries + 1
      For I = 1 To F3BNumSlots(Task)
        For J = 1 To (MaxPilotsInSlot + 10)
         alloc(Task, RoundNum, I, J) = 0
        Next J
      Next I
      OK = True
    Else
      'Text4.Text = Str(maxtimes)
      RoundTries = RoundTries + 1
      For I = 1 To F3BNumSlots(Task)
        For J = 1 To (MaxPilotsInSlot + 10)
         alloc(Task, RoundNum, I, J) = 0
        Next J
      Next I
      OK = True
    End If
    If RoundTries > (NumPilots * 10) Then
      Message = "      TOO MANY FREQUENCY OR TEAM CONFLICTS" & Chr(10) & Chr(13) & Chr(13) & "Either shift some pilots off one frequency or ensure that the teams do not have too many members and the team members are on the same frequency"
      Response = MsgBox(Message, 16, "Allocate Program")
      Command4.Enabled = False
      'Command2.Enabled = False
      Command3.Enabled = False
      Exit Sub
    End If
    Selected.MoveFirst
  Loop
  Next Task
  'F3BNumSlots(Task) = MaxSlots
  'Call CheckDupes(dupes(), maxnum, NumPilots)
  UpdateBox
  'Picture1.Cls
  'Set cRndGen = Nothing
Exit Sub
errhandler:
  MsgBox ("You have not entered any competitors")

End Sub



Private Sub Command2_Click()
  RoundMaint.Show
  Command3.Visible = True
End Sub

Private Sub Command3_Click()
  If CurrentContestType = "F3J" Then
    'F3JMakeScores
  ElseIf CurrentContestType = "AustOpen" Then
    'F3JMakeScores
  ElseIf CurrentContestType = "F3B" Then
    F3BMakeScores
  End If
End Sub

Private Sub Command4_Click()
  'Command3.Visible = False
  Text4.Text = ""
  MainAllocate
End Sub


Private Sub Command5_Click()
  Unload Me
  
End Sub


Private Sub UpdateBox()
 Dim RoundNum As Integer
 Dim Slot As Integer
 Dim P As Integer
 Dim Line As String
 Dim SQLString As String
 Dim Selected As Recordset
 Dim I As Integer
 Dim Task As Integer
 Dim LineC As Integer
 
 'Print the contest out to a form
 For I = 1 To 3
   List1(I).Clear
 Next I
 For Task = 1 To 3
   P = 1
   LineC = 1
   For RoundNum = 1 To NumRounds
     List1(Task).AddItem " "
     List1(Task).AddItem "Round " & Str(RoundNum)
     For Slot = 1 To F3BNumSlots(Task)
       Line = "Slot " & Str(Slot) & "  "
       Do Until alloc(Task, RoundNum, Slot, P) = 0
         SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, Pilot.First_Name, Pilot.Last_Name, List.FreqUsed, List.Flyoff FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID WHERE ((List.Pilot_ID=" & Str(alloc(Task, RoundNum, Slot, P)) & ") AND (List.Comp_ID=" & Str(CurrentContest) & "));"
         Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
         Line = Line + Left(Trim(Selected!First_Name), 1) & "." & Selected!Last_Name & " "
         P = P + 1
         LineC = LineC + 1
         If LineC > 8 Then
           List1(Task).AddItem Line
           Line = "           "
           LineC = 1
         End If
       Loop
       List1(Task).AddItem Line
       Line = ""
       P = 1
       LineC = 1
     Next Slot
    P = 1
    LineC = 1
   Next RoundNum
 Next Task
End Sub

Private Sub Command6_Click()
  allocated = False
  ThisTask = 1
  F3BRoundMaint.Show
End Sub


Public Sub F3BMakeScores()
Dim RoundNum As Integer
Dim Slot As Integer
Dim Task As Integer
Dim P As Integer
Dim SetTable As Recordset
Dim ScoreTable As Recordset
Dim SearchStr As String
Dim Ver As Integer

  Ver = 0
  Set ScoreTable = F3JDb.OpenRecordset("Scores", dbOpenTable)
  'Construct all the scores in the scores database
    P = 1
    For Task = 1 To 3
      For RoundNum = 1 To NumRounds
        For Slot = 1 To F3BNumSlots(Task)
          Do Until alloc(Task, RoundNum, Slot, P) = 0
            'SearchStr = "Pilot_ID = " + Str(alloc(RoundNum, Slot, P))
            'Selected.FindFirst SearchStr
             ScoreTable.AddNew
             ScoreTable!Comp_ID = CurrentContest
             'ScoreTable!Pilot_ID = Selected!Pilot_ID
             ScoreTable!Pilot_ID = alloc(Task, RoundNum, Slot, P)
             ScoreTable!Round = RoundNum
             ScoreTable!Slot = Slot
             If Task = 1 Then
               ScoreTable!Task = "A"
             ElseIf Task = 2 Then
               ScoreTable!Task = "B"
             ElseIf Task = 3 Then
               ScoreTable!Task = "C"
             End If
             ScoreTable.Update
             'Selected.MoveFirst
             P = P + 1
             Ver = Ver + 1
          Loop
        P = 1
        Next Slot
      P = 1
      Next RoundNum
    Next Task
    If Ver <> NumRounds * NumberOfPilots * 3 Then
      MsgBox ("The Round allocation is corrupt" & Chr(13) & "Please Reset Contest and Re-Allocate")
      Form1.ContestReset
    Else
      Set SetTable = F3JDb.OpenRecordset("Competition", dbOpenDynaset)
      SearchStr = "Comp_ID = " & Str(CurrentContest)
      SetTable.FindFirst SearchStr
      SetTable.Edit
      SetTable!RoundsDone = True
      CompRoundsDone = True
      SetTable!NumSlots = F3BNumSlots(1)
      SetTable!Res1 = Trim(Str(F3BNumSlots(2)))
      SetTable!Res2 = F3BNumSlots(3)
      SetTable.Update
      SetTable.Close
      MsgBox ("Pilots set up in scores database")
      ScoreTable.Close
   End If
End Sub


Public Sub RoundSlotAlloc2(ByVal Task, Selected, ByRef Flight, ByRef Freq, ByRef Team, RoundNum As Integer, ByVal MaxPilotsInSlot As Integer)
Dim SearchStr As String
Dim Currentpilot As Integer
Dim FirstFreq As String
Dim SecondFreq As String
Dim Slot As Integer
Dim P As Integer
Dim Tries As Integer
Dim Line1 As String
Dim CurrentPilotTeam As Integer
Dim PilotCount As Integer
Dim SecondTeam As Integer
Dim SecondPilot As Integer
Dim Check As Recordset
Dim testvar As Boolean
  
  testvar = False
  Finish = False
  PilotCount = 0
  Tries = 1
  Slot = 1
  
  Selected.MoveFirst
  Do Until Selected.EOF  'loop until there are no more pilots
    Slot = 0
    allocated = False
    Currentpilot = Selected!Pilot_ID
    CurrentPilotTeam = Selected!Team_ID
    FirstFreq = Selected!FreqUsed 'current record
    Do Until Slot > F3BNumSlots(Task) + 1 Or allocated = True
      Slot = Slot + 1 'Pilot will not go into slot - go onto next slot
      If Slot > F3BNumSlots(Task) Then 'if the pilot will not fit into any slots,
        If Tries > 16 Then
          OK = False
          Exit Sub
        Else
          Slot = 1 'before increasing the number of slots, try the pilot in the previous slots
          'MaxPilotsInSlot = MaxPilotsInSlot + 1
          Tries = Tries + 1
        End If
      End If
      P = 1
      OK = True
      Do Until P > MaxPilotsInSlot 'do until a freq conflict
        'if we have got to a blank entry in the alloc matrix
        'then it must be alright to alocate the pilot to this spot
        If alloc(Task, RoundNum, Slot, P) = 0 Then
          'allocate pilot to slot
           alloc(Task, RoundNum, Slot, P) = Currentpilot
           P = 1
          'updating flight matrix with pilot information
          'after a successful allocation
          allocated = True
          Exit Do
        End If
        'we are not at a blank entry so we must check for
        'frequency conflicts
        'Debug.Print CurrentPilotTeam, " ", Team(alloc(RoundNum, Slot, P))
        If Abs(FirstFreq - Freq(alloc(Task, RoundNum, Slot, P))) <= 0.011 Then
        
          If TenK = False Then
            Exit Do 'Frequency is within 10KC and this is not allowed
          ElseIf FirstFreq = Freq(alloc(Task, RoundNum, Slot, P)) Then
            Exit Do 'Freq Conflict discovered break out of loop where 10Kc splits are allowed
          Else
            If Flight(Task, Currentpilot, alloc(Task, RoundNum, Slot, P)) >= maxtimes Then
              Exit Do 'If the pilot has flown against this pilot too many times
          
            ElseIf CurrentPilotTeam = Team(alloc(Task, RoundNum, Slot, P)) Then
              If Selected!Team_ID > 0 Then 'If Team_ID > 0 then pilot is not in team
                Exit Do 'Pilot is in same team as others
              End If
            End If
            P = P + 1
          End If
        ElseIf Flight(Task, Currentpilot, alloc(Task, RoundNum, Slot, P)) >= maxtimes Then
          Exit Do 'If the pilot has flown against this pilot too many times
      
         'ElseIf CurrentPilotTeam = Selected!Team_ID Then 'If Team_ID is equal then break
      
        ElseIf CurrentPilotTeam = Team(alloc(Task, RoundNum, Slot, P)) Then
        
          If Selected!Team_ID = 0 Then 'If Team_ID > 0 then pilot is not in team
            P = P + 1 'go on the next pilot
          Else
            Exit Do 'Pilot is in same team as others
          End If
      
        Else
          P = P + 1 'Go onto next pilot in slot
      
        End If
      Loop
    Loop
    Selected.MoveNext ' move to the next pilot
    PilotCount = PilotCount + 1

  Loop
  alloc(Task, RoundNum, 0, 0) = PilotCount
  If PilotCount <> Selected.RecordCount Then
    OK = False
  End If

End Sub


Public Sub CheckDupes(dupes() As Integer, maxnum As Integer, NumPilots As Integer)
Dim num() As Integer
Dim RoundNum As Integer
Dim Slot As Integer
Dim PCount As Integer
Dim I As Integer
Dim P As Integer
Dim Count As Integer
Dim Found As Boolean
Dim Skip As Integer
Dim Task As Integer

 ReDim num(NumSlots)
 
  P = 1
  PCount = 0
  For Task = 1 To 3
  For RoundNum = 1 To NumRounds
    For Slot = 1 To NumSlots
      Do Until alloc(Task, RoundNum, Slot, P) = 0
        dupes(alloc(Task, RoundNum, Slot, P)) = dupes(alloc(Task, RoundNum, Slot, P)) + 1
        num(Slot) = num(Slot) + 1
        P = P + 1
      Loop
      PCount = PCount + num(Slot)
      P = 1
    Next Slot
    P = 1
  
    Skip = 0
    For I = 1 To maxnum
      If dupes(I) > 1 Then
      Debug.Print "duplicate", " ", I
        Slot = 1
        Do Until Slot > NumSlots
          P = 1
          Do Until alloc(Task, RoundNum, Slot, P) = 0
            If alloc(Task, RoundNum, Slot, P) = dupes(I) Then
              If Skip > 1 Then
                For Count = P To num(Slot)
                  alloc(Task, RoundNum, Slot, Count) = alloc(RoundNum, Slot, Count + 1)
                Next Count
              Else
               Skip = Skip + 1
              End If
            End If
            P = P + 1
          Loop
          P = 1
          Slot = Slot + 1
        Loop
      End If
    Skip = 0
    Next I
    For I = 1 To maxnum
      dupes(I) = 0
    Next I
 Next RoundNum
 Next Task
End Sub

Private Sub Command7_Click()
  ThisTask = 1
  F3BRoundMaint.Show
  
End Sub

Private Sub Command8_Click()
 ThisTask = 2
  F3BRoundMaint.Show
End Sub


Private Sub Command9_Click()
  ThisTask = 3
  F3BRoundMaint.Show
End Sub


Private Sub Form_Load()

Dim Selected As Recordset
Dim SQLString As String
Dim NumPilots As Integer

Dim MaxPilotsInSlot As Integer
Dim I As Integer
  
  SQLString = "SELECT * FROM List Where ((List.Comp_ID =" & Str(CurrentContest) & ")) ORDER BY List.Res2 DESC;"
  
  Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)

  NumPilots = Selected.RecordCount
  NumberOfPilots = Selected.RecordCount
  NumberPilots = Selected.RecordCount
  If NumberOfPilots > 100 Then
    NumSlots = (Int(Sqr(NumPilots))) + 4
    MaxPilotsInSlot = NumPilots / NumSlots
    If MaxPilotsInSlot > 10 Then
      MaxPilotsInSlot = 10
    End If
    Do Until NumSlots * MaxPilotsInSlot > NumPilots
      NumSlots = NumSlots + 1
    Loop
    Text1.Text = NumPilots
    For I = 1 To 3
      Text2(I).Text = NumSlots
      Text3(I).Text = MaxPilotsInSlot
    Next I
    
  Else
    NumSlots = Int(Sqr(NumPilots))
    MaxPilotsInSlot = Int(NumPilots / NumSlots)
    Do Until NumSlots * MaxPilotsInSlot > NumPilots
      NumSlots = NumSlots + 1
    Loop
    Text1.Text = NumPilots
    For I = 1 To 3
      Text2(I).Text = NumSlots
      Text3(I).Text = MaxPilotsInSlot
    Next I
  End If
   
End Sub


