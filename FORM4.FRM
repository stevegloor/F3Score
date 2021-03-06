VERSION 5.00
Begin VB.Form Allocate 
   Caption         =   "Slot Allocation Form"
   ClientHeight    =   5790
   ClientLeft      =   195
   ClientTop       =   525
   ClientWidth     =   8685
   HelpContextID   =   350
   LinkTopic       =   "Form4"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   24.125
   ScaleMode       =   4  'Character
   ScaleWidth      =   72.375
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Finish Allocation Manually"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   5280
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   5835
      TabIndex        =   8
      Top             =   120
      Width           =   5895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contest Slot Allocation"
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   600
      WhatsThisHelpID =   350
      Width           =   8415
      Begin VB.ListBox List1 
         Height          =   3765
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   8175
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   5280
      WhatsThisHelpID =   350
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Allocate Slots"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      WhatsThisHelpID =   350
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Accept Allocation and Set Up Scores In Database"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      WhatsThisHelpID =   350
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change Allocations"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   4800
      WhatsThisHelpID =   350
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Don't Accept Allocations"
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   4800
      WhatsThisHelpID =   350
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Allocation Progress"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      WhatsThisHelpID =   350
      Width           =   1935
   End
End
Attribute VB_Name = "Allocate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OK As Boolean
Public NumberOfPilots As Integer
Public allocated As Boolean

Private Sub Command1_Click()
  Command4.Enabled = True
  Command2.Enabled = True
  Command3.Enabled = True
  List1.Clear
  Picture1.Cls
  ReDim alloc(0, 0, 0)
End Sub


Public Sub MainAllocate()
  If CurrentContestType = "F3J" Then
    AllocateSlots
  ElseIf CurrentContestType = "F3B" Then
    AllocateSlots
  ElseIf CurrentContestType = "Club" Then
    AllocateClub
  ElseIf CurrentContestType = "F3JFO" Then
    AllocateClub
  ElseIf CurrentContestType = "AustOpen" Then
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
 
  On Error GoTo errhandler
  
  Picture1.ForeColor = RGB(0, 0, 255)     'use blue bar
  OK = True
  RoundTries = 0
  maxtimes = 1
  MaxSlots = 0
  
  SQLString(2) = "SELECT * FROM List Where ((List.Comp_ID =" & Str(CurrentContest) & ")) ORDER BY List.FreqUsed ASC;"
  SQLString(1) = "SELECT * FROM List Where ((List.Comp_ID =" & Str(CurrentContest) & ")) ORDER BY List.Rnd1 ASC;"
  
  Set Selected = F3JDb.OpenRecordset(SQLString(1), dbOpenDynaset)
  NumPilots = Selected.RecordCount
  NumberOfPilots = Selected.RecordCount
  If NumberOfPilots > 100 Then
    NumSlots = (Int(Sqr(NumPilots))) + 4
    MaxPilotsInSlot = NumPilots / NumSlots
    If MaxPilotsInSlot > 10 Then
      MaxPilotsInSlot = 10
    End If
  Else
    NumSlots = Int(Sqr(NumPilots))
    MaxPilotsInSlot = Int(NumPilots / NumSlots)
  
  End If
  
  
  OldValue = NumSlots
  'Find out the maximum pilot ID
  Selected.MoveFirst
  Do Until Selected.EOF
    If Selected!Pilot_ID > maxnum Then
      maxnum = Selected!Pilot_ID
    End If
    Selected.MoveNext
  Loop
  'Dimension the arrays to the right sizes
  ReDim Flight(maxnum, maxnum)
  ReDim alloc(NumRounds, NumSlots + 10, MaxPilotsInSlot + 10)
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
  
  RoundNum = 1
  'allocat pilots to rounds
  Do Until RoundNum > NumRounds
    Selected.Close
    'put in new random numbers
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
    If MaxSlots <= NumSlots Then
      MaxSlots = NumSlots
    End If
    NumSlots = OldValue
    Selected.MoveFirst
    Call RoundSlotAlloc2(Selected, Flight(), Freq(), Team(), RoundNum, MaxPilotsInSlot) 'Call slot allocation procedure
    'If the round has been drawn correctly
    If OK = True Then
      updateprogress Picture1, (RoundNum / NumRounds) * 100
      RoundNum = RoundNum + 1
      RoundsCorrect = RoundNum
    ElseIf RoundTries > 100 And RoundTries < 200 Then
      maxtimes = maxtimes + 1
      Text4.Text = Str(maxtimes)
      RoundTries = RoundTries + 1
      For I = 1 To NumSlots
        For J = 1 To (MaxPilotsInSlot + 10)
         alloc(RoundNum, I, J) = 0
        Next J
      Next I
      OK = True
    Else
      Text4.Text = Str(maxtimes)
      RoundTries = RoundTries + 1
      For I = 1 To NumSlots
        For J = 1 To (MaxPilotsInSlot + 10)
         alloc(RoundNum, I, J) = 0
        Next J
      Next I
      OK = True
    End If
    If RoundTries > 200 Then
      Message = "      TOO MANY FREQUENCY OR TEAM CONFLICTS" & Chr(10) & Chr(13) & Chr(13) & "Either shift some pilots off one frequency or ensure that the teams do not have too many members and the team members are on the same frequency"
      Response = MsgBox(Message, 16, "Allocate Program")
      Command4.Enabled = False
      Command2.Enabled = False
      Command3.Enabled = False
      Exit Sub
    End If
    Selected.MoveFirst
  Loop
  NumSlots = MaxSlots
  'Call CheckDupes(dupes(), maxnum, NumPilots)
  UpdateBox
  'Picture1.Cls
  
Exit Sub
errhandler:
  MsgBox ("You have not entered any competitors")

End Sub

Public Sub AllocateClub()
  Dim NumPilots As Integer
  Dim ThisTask As String
  Dim Count As Integer
  Dim RoundCount As Integer
  Dim SearchStr As String
  Dim SQLString As String
  Dim Selected As Recordset
  Dim ScoreTable As Recordset
  Dim SetTable As Recordset
  
  SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, Pilot.Freq1, Pilot.Freq2 FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID =" & Str(CurrentContest) & ")) ORDER BY Pilot.Freq1 ASC;"
  Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  Set ScoreTable = F3JDb.OpenRecordset("Scores", dbOpenTable)
  
  If Selected.RecordCount = 0 Then
    MsgBox ("You have not entered any Pilots")
    Exit Sub
  Else
    NumPilots = Selected.RecordCount
    Selected.MoveFirst
    For RoundCount = 1 To NumRounds
      Do Until Selected.EOF = True
        ScoreTable.AddNew
        ScoreTable!Round = RoundCount
        ScoreTable!Slot = 1
        ScoreTable!Comp_ID = CurrentContest
        ScoreTable!Pilot_ID = Selected!Pilot_ID
        ScoreTable!task = "A"
        ScoreTable.Update
        Selected.MoveNext
      Loop
    Selected.MoveFirst
    Next RoundCount
    NumSlots = 1
    Set SetTable = F3JDb.OpenRecordset("Competition", dbOpenDynaset)
    SearchStr = "Comp_ID = " & Str(CurrentContest)
    SetTable.FindFirst SearchStr
    SetTable.Edit
    SetTable!RoundsDone = True
    CompRoundsDone = True
    SetTable!NumSlots = 1
    SetTable.Update
    MsgBox ("Rounds Set Up In Database")
  End If
End Sub



Private Sub RoundSlotAlloc(ByVal Selected, ByRef Flight, RoundNum, ByVal MaxPilotsInSlot)
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
Do Until Selected.EOF Or Finish = True 'loop until there are no more pilots
  Slot = 0
  allocated = False
  Currentpilot = Selected!Pilot_ID
  CurrentPilotTeam = Selected!Team_ID
  FirstFreq = Selected!FreqUsed 'current record
  Do Until Slot > NumSlots Or allocated = True Or Finish = True
    Slot = Slot + 1 'Pilot will not go into slot - go onto next slot
      If Slot > NumSlots Then 'if the pilot will not fit into any slots,
        If Tries = 2 Then 'if we have had 2 tries at fitting the pilot in
          NumSlots = NumSlots + 1 'Increase the number of slots
          MaxPilotsInSlot = MaxPilotsInSlot - 2 'decrease the number of pilots in the slots
          Tries = Tries + 1
       
        ElseIf Tries > 4 Then
          OK = False
          Exit Sub
     
        Else
          Slot = 1 'before increasing the number of slots, try the pilot in the previous slots
          MaxPilotsInSlot = MaxPilotsInSlot + 1
          Tries = Tries + 1
        End If
      End If
    P = 1
    OK = True
    Do Until P > MaxPilotsInSlot 'do until a freq conflict
      'if we have got to a blank entry in the alloc matrix
      'then it must be alright to alocate the pilot to this spot
      If alloc(RoundNum, Slot, P) = 0 Then
        SearchStr = "Pilot_ID = " + Str(Currentpilot) 'or no more pilots in slot
        Selected.FindFirst SearchStr 'then allocate pilot to slot
        alloc(RoundNum, Slot, P) = Selected!Pilot_ID
        P = 1
        'udating flight matrix with pilot information
        'after a successful allocation
        Do Until alloc(RoundNum, Slot, P) = 0 'update
          SearchStr = "Pilot_ID =" + Str(alloc(RoundNum, Slot, P))
          Selected.FindFirst SearchStr
          If Currentpilot <> Selected!Pilot_ID Then
            Flight(Currentpilot, Selected!Pilot_ID) = Flight(Currentpilot, Selected!Pilot_ID) + 1
          End If
          If P <> MaxPilotsInSlot Then
            P = P + 1
          Else
            Exit Do
          End If
        Loop 'update
        allocated = True
        Exit Do
      End If
      'we are not at a blank entry so we must check for
      'frequency conflicts
      SearchStr = "Pilot_ID =" + Str(alloc(RoundNum, Slot, P))
      Selected.FindFirst SearchStr
      If FirstFreq = Selected!FreqUsed Then
      'If Abs(FirstFreq - Selected!FreqUsed) <= 0.011 Then
        Exit Do 'Freq Conflict discovered break out of loop
      
      ElseIf Flight(Currentpilot, Selected!Pilot_ID) >= maxtimes Then
        Exit Do 'If the pilot has flown against this pilot too many times
      
      ElseIf CurrentPilotTeam = Selected!Team_ID Then 'If Team_ID is equal then break
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
alloc(RoundNum, 0, 0) = maxtimes
If PilotCount <> Selected.RecordCount Then
  OK = False
End If
End Sub

Public Sub F3JMakeScores()
Dim RoundNum As Integer
Dim Slot As Integer
Dim P As Integer
Dim SetTable As Recordset
Dim ScoreTable As Recordset
Dim SearchStr As String
Dim Ver As Integer

  Ver = 0
  Set ScoreTable = F3JDb.OpenRecordset("Scores", dbOpenTable)
  'Construct all the scores in the scores database
    P = 1
    For RoundNum = 1 To NumRounds
     For Slot = 1 To NumSlots
      Do Until alloc(RoundNum, Slot, P) = 0
        'SearchStr = "Pilot_ID = " + Str(alloc(RoundNum, Slot, P))
        'Selected.FindFirst SearchStr
        ScoreTable.AddNew
        ScoreTable!Comp_ID = CurrentContest
        'ScoreTable!Pilot_ID = Selected!Pilot_ID
        ScoreTable!Pilot_ID = alloc(RoundNum, Slot, P)
        ScoreTable!Round = RoundNum
        ScoreTable!Slot = Slot
        ScoreTable!task = "A"
        ScoreTable.Update
        'Selected.MoveFirst
        P = P + 1
        Ver = Ver + 1
      Loop
     P = 1
     Next Slot
    P = 1
    Next RoundNum
    If Ver <> NumRounds * NumberOfPilots Then
      MsgBox ("The RoundNum allocation is corrupt" & Chr(13) & "Please Reset Contest and Re-Allocate")
      Form1.ContestReset
    Else
      Set SetTable = F3JDb.OpenRecordset("Competition", dbOpenDynaset)
      SearchStr = "Comp_ID = " & Str(CurrentContest)
      SetTable.FindFirst SearchStr
      SetTable.Edit
      SetTable!RoundsDone = True
      CompRoundsDone = True
      SetTable!NumSlots = NumSlots
      SetTable.Update
      SetTable.Close
      MsgBox ("Pilots set up in scores database")
      ScoreTable.Close
   End If
End Sub


Private Sub Command2_Click()
  RoundMaint.Show
End Sub

Private Sub Command3_Click()
  If CurrentContestType = "F3J" Then
    F3JMakeScores
  ElseIf CurrentContestType = "AustOpen" Then
    F3JMakeScores
  ElseIf CurrentContestType = "F3B" Then
    F3BMakeScores
  End If
End Sub

Private Sub Command4_Click()
  
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
 Dim LineC As Integer
 
 'Print the contest out to a form
 List1.Clear
 P = 1
 For RoundNum = 1 To NumRounds
  List1.AddItem " "
   List1.AddItem "Round " & Str(RoundNum)
   For Slot = 1 To NumSlots
   Line = "Slot " & Str(Slot) & "  "
    Do Until alloc(RoundNum, Slot, P) = 0
      SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, Pilot.First_Name, Pilot.Last_Name, List.FreqUsed, List.Flyoff FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID WHERE ((List.Pilot_ID=" & Str(alloc(RoundNum, Slot, P)) & ") AND (List.Comp_ID=" & Str(CurrentContest) & "));"
      Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
      Line = Line + Left(Trim(Selected!First_Name), 1) & "." & Selected!Last_Name & " "
      LineC = LineC + 1
         If LineC > 8 Then
           List1.AddItem Line
           Line = "           "
           LineC = 1
         End If
      P = P + 1
      LineC = LineC + 1
    Loop
   List1.AddItem Line
   Line = ""
   P = 1
   LineC = 1
   Next Slot
  P = 1
  LineC = 1
 Next RoundNum
End Sub

Private Sub Command6_Click()
  allocated = False
  RoundMaint.Show
End Sub


Public Sub F3BMakeScores()
Dim RoundNum As Integer
Dim Slot As Integer
Dim task As Integer
Dim P As Integer
Dim SetTable As Recordset
Dim ScoreTable As Recordset
Dim SearchStr As String
Dim Ver As Integer

  Ver = 0
  Set ScoreTable = F3JDb.OpenRecordset("Scores", dbOpenTable)
  'Construct all the scores in the scores database
    P = 1
    For task = 1 To 3
      For RoundNum = 1 To NumRounds
        For Slot = 1 To NumSlots
          Do Until alloc(RoundNum, Slot, P) = 0
            'SearchStr = "Pilot_ID = " + Str(alloc(RoundNum, Slot, P))
            'Selected.FindFirst SearchStr
             ScoreTable.AddNew
             ScoreTable!Comp_ID = CurrentContest
             'ScoreTable!Pilot_ID = Selected!Pilot_ID
             ScoreTable!Pilot_ID = alloc(RoundNum, Slot, P)
             ScoreTable!Round = RoundNum
             ScoreTable!Slot = Slot
             If task = 1 Then
               ScoreTable!task = "A"
             ElseIf task = 2 Then
               ScoreTable!task = "B"
             ElseIf task = 3 Then
               ScoreTable!task = "C"
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
    Next task
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
      SetTable!NumSlots = NumSlots
      SetTable.Update
      SetTable.Close
      MsgBox ("Pilots set up in scores database")
      ScoreTable.Close
   End If
End Sub


Public Sub RoundSlotAlloc2(ByVal Selected, ByRef Flight, ByRef Freq, ByRef Team, RoundNum As Integer, ByVal MaxPilotsInSlot As Integer)
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
  Do Until Selected.EOF Or Finish = True 'loop until there are no more pilots
  Slot = 0
  allocated = False
  Currentpilot = Selected!Pilot_ID
  CurrentPilotTeam = Selected!Team_ID
  FirstFreq = Selected!FreqUsed 'current record
  Do Until Slot > NumSlots + 1 Or allocated = True Or Finish = True
    Slot = Slot + 1 'Pilot will not go into slot - go onto next slot
      If Slot > NumSlots Then 'if the pilot will not fit into any slots,
        If Tries = 2 Then 'if we have had 2 tries at fitting the pilot in
          NumSlots = NumSlots + 1 'Increase the number of slots
          MaxPilotsInSlot = MaxPilotsInSlot - 2 'decrease the number of pilots in the slots
          Tries = Tries + 1
       
        ElseIf Tries > 8 Then
          OK = False
          Exit Sub
     
        Else
          Slot = 1 'before increasing the number of slots, try the pilot in the previous slots
          MaxPilotsInSlot = MaxPilotsInSlot + 1
          Tries = Tries + 1
        End If
      End If
    P = 1
    OK = True
    Do Until P > MaxPilotsInSlot 'do until a freq conflict
      'if we have got to a blank entry in the alloc matrix
      'then it must be alright to alocate the pilot to this spot
      If alloc(RoundNum, Slot, P) = 0 Then
        'allocate pilot to slot
        alloc(RoundNum, Slot, P) = Currentpilot
        P = 1
        'updating flight matrix with pilot information
        'after a successful allocation
        Do Until alloc(RoundNum, Slot, P) = 0 'update
          If Currentpilot <> alloc(RoundNum, Slot, P) Then
             Flight(Currentpilot, alloc(RoundNum, Slot, P)) = Flight(Currentpilot, alloc(RoundNum, Slot, P)) + 1
          End If
          If P <> MaxPilotsInSlot Then
            P = P + 1
          Else
            Exit Do
          End If
        Loop 'update
        allocated = True
        Exit Do
      End If
      'we are not at a blank entry so we must check for
      'frequency conflicts
      'Debug.Print CurrentPilotTeam, " ", Team(alloc(RoundNum, Slot, P))
      If Abs(FirstFreq - Freq(alloc(RoundNum, Slot, P))) <= 0.011 Then
        If TenK = False Then
          Exit Do 'Frequency is within 10KC and this is not allowed
        ElseIf FirstFreq = Freq(alloc(RoundNum, Slot, P)) Then
          Exit Do 'Freq Conflict discovered break out of loop where 10Kc splits are allowed
        Else
          If Flight(Currentpilot, alloc(RoundNum, Slot, P)) >= maxtimes Then
            Exit Do 'If the pilot has flown against this pilot too many times
          
          ElseIf CurrentPilotTeam = Team(alloc(RoundNum, Slot, P)) Then
            If Selected!Team_ID = 0 Then 'If Team_ID > 0 then pilot is not in team
              P = P + 1 'go on the next pilot
            Else
              Exit Do 'Pilot is in same team as others
            End If
          
          End If
        P = P + 1
        End If
      ElseIf Flight(Currentpilot, alloc(RoundNum, Slot, P)) >= maxtimes Then
        Exit Do 'If the pilot has flown against this pilot too many times
      
      'ElseIf CurrentPilotTeam = Selected!Team_ID Then 'If Team_ID is equal then break
      
      ElseIf CurrentPilotTeam = Team(alloc(RoundNum, Slot, P)) Then
        
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
alloc(RoundNum, 0, 0) = PilotCount
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

 ReDim num(NumSlots)
 
  P = 1
  PCount = 0
  For RoundNum = 1 To NumRounds
    For Slot = 1 To NumSlots
      Do Until alloc(RoundNum, Slot, P) = 0
        dupes(alloc(RoundNum, Slot, P)) = dupes(alloc(RoundNum, Slot, P)) + 1
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
          Do Until alloc(RoundNum, Slot, P) = 0
            If alloc(RoundNum, Slot, P) = dupes(I) Then
              If Skip > 1 Then
                For Count = P To num(Slot)
                  alloc(RoundNum, Slot, Count) = alloc(RoundNum, Slot, Count + 1)
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
End Sub

Private Sub Form_Load()
Dim Selected As Recordset
Dim SQLString As String


  SQLString = "SELECT * FROM List Where ((List.Comp_ID =" & Str(CurrentContest) & ")) ORDER BY List.Res2 DESC;"
  Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  NumberPilots = Selected.RecordCount
  Selected.Close
  
End Sub


