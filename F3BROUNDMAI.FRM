VERSION 5.00
Begin VB.Form F3BRoundMaint 
   Caption         =   "Round/Slot Allocation Maintenance"
   ClientHeight    =   6120
   ClientLeft      =   960
   ClientTop       =   690
   ClientWidth     =   5250
   HelpContextID   =   400
   LinkTopic       =   "Form6"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6120
   ScaleWidth      =   5250
   Begin VB.CommandButton Command4 
      Caption         =   "CANCEL- Don't Save Changes"
      Height          =   495
      Left            =   3600
      TabIndex        =   16
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update Allocation"
      Height          =   495
      Left            =   3600
      TabIndex        =   14
      Top             =   4800
      WhatsThisHelpID =   400
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "OverView"
      Height          =   4695
      Left            =   3600
      TabIndex        =   12
      Top             =   0
      WhatsThisHelpID =   400
      Width           =   1575
      Begin VB.ListBox List3 
         Height          =   4350
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Pilot From Slot"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      WhatsThisHelpID =   400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Pilot to Slot"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   3240
      WhatsThisHelpID =   400
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pilots Not Allocated"
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      WhatsThisHelpID =   400
      Width           =   3255
      Begin VB.ListBox List2 
         Height          =   1035
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pilots in Selected Slot"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      WhatsThisHelpID =   400
      Width           =   3375
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   3135
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
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   495
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
         Left            =   1920
         TabIndex        =   3
         Top             =   600
         Width           =   495
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   495
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         WhatsThisHelpID =   400
         Width           =   255
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   495
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         WhatsThisHelpID =   400
         Width           =   255
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
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   855
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
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ensure in Overview that all rounds have the same number of pilots - edit as necessary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   5280
      Width           =   3495
   End
End
Attribute VB_Name = "F3BRoundMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pilotcheck(40) As Integer
Dim roundcheck As Boolean
Public Sub UpdateList1()
Dim RoundNum As Integer
Dim Slot As Integer
Dim FoundSlot As Integer
Dim P As Integer
Dim SearchStr As String
Dim Selected As Recordset
Dim Line As String
Dim SQLString As String
Dim Currentpilot As Integer
Dim Match As Integer
Dim Pos As Integer
Dim num As Integer
Dim Count As Integer
Dim I As Integer

 If allocated = True Then
    List1.Clear
    RoundNum = Val(Text1.Text)
    Slot = Val(Text2.Text)
    P = 1
    If alloc(ThisTask, RoundNum, Slot, P) <> 0 Then
      Do Until alloc(ThisTask, RoundNum, Slot, P) = 0
        SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, Pilot.First_Name, Pilot.Last_Name, List.FreqUsed, List.Flyoff FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID WHERE ((List.Pilot_ID=" & Str(alloc(RoundNum, Slot, P)) & ") AND (List.Comp_ID=" & Str(CurrentContest) & "));"
        Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
        Line = Format(Str(Selected!Pilot_ID), "!@@@@") & "   " & Selected!FreqUsed & "   " & Selected!First_Name & " " & Selected!Last_Name
        List1.AddItem Line
        P = P + 1
      Loop
      Selected.Close
    Else
      List1.AddItem "No Pilots"
    End If
  Else
    SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, Pilot.First_Name, Pilot.Last_Name, List.FreqUsed, List.Flyoff FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID WHERE ((List.Comp_ID=" & Str(CurrentContest) & "));"
    Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    Do Until Selected.EOF
      Selected.Edit
      Selected!Flyoff = False
      Selected.Update
      Selected.MoveNext
    Loop
    List1.Clear
    RoundNum = Val(Text1.Text)
    Slot = Val(Text2.Text)
    P = 1
    SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, Pilot.First_Name, Pilot.Last_Name, List.FreqUsed, List.Flyoff FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID WHERE ((List.Comp_ID=" & Str(CurrentContest) & "));"
    Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    
    P = 1
    Selected.MoveFirst
    Currentpilot = Selected!Pilot_ID
    Match = 0
    FoundSlot = 0
    Do Until Selected.EOF
      For I = 1 To F3BNumSlots(ThisTask)
        Do Until alloc(ThisTask, RoundNum, I, P) = 0
          If alloc(ThisTask, RoundNum, I, P) = Currentpilot Then
            Match = Match + 1   'Record a match in the allocation table - should be 1
            FoundSlot = I       'Record where the match was
          End If
          P = P + 1
        Loop
        P = 1
      Next I
      If Match = 0 Then         'if a pilot has no matches he is not allocated
        Selected.Edit           'so set the flag to true
        Selected!Flyoff = True
        Selected.Update
      
      ElseIf Match > 1 Then   'if pilot has more than one match then there is a
        P = Selected!Pilot_ID 'duplication so eliminate
        Pos = 1
        num = 0
        Do Until alloc(ThisTask, RoundNum, FoundSlot, Pos) = 0 'Find the number of pilots in slot
          Pos = Pos + 1
          num = num + 1
        Loop
        Pos = 1
        Do Until alloc(ThisTask, RoundNum, FoundSlot, Pos) = P 'Find the position of the pilot
          Pos = Pos + 1
        Loop
        For Count = Pos To num
          alloc(ThisTask, RoundNum, FoundSlot, Count) = alloc(RoundNum, Slot, Count + 1)
        Next Count
      End If
      P = 1
      FoundSlot = 0
      Match = 0
      Selected.MoveNext
      If Selected.EOF = False Then
        Currentpilot = Selected!Pilot_ID
      End If
    Loop
    
    P = 1
    Do Until alloc(ThisTask, RoundNum, Slot, P) = 0
      SearchStr = "Pilot_ID = " & Str(alloc(ThisTask, RoundNum, Slot, P))
      Selected.FindFirst SearchStr
      Line = Format(Str(Selected!Pilot_ID), "!@@@@") & "   " & Format(Selected!FreqUsed, "##.000") & "   " & Selected!First_Name & " " & Selected!Last_Name
      List1.AddItem Line
      P = P + 1
    Loop
    Selected.Close
    
  End If
End Sub

Public Sub UpdateList2()
Dim SQLString As String
Dim Selected As Recordset
Dim Line As String
  
  List2.Clear
  SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, Pilot.First_Name, Pilot.Last_Name, List.FreqUsed, List.Flyoff FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID WHERE ((List.FlyOff = true) AND (List.Comp_ID=" & Str(CurrentContest) & "));"
  Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If Selected.RecordCount = 0 Then
    List2.AddItem "No Pilots"
    VScroll1.Enabled = True
  Else
    Selected.MoveFirst
    Do Until Selected.EOF
      Line = Format(Str(Selected!Pilot_ID), "!@@@@") & "  " & Format(Selected!FreqUsed, "##.000") & "  " & Selected!First_Name & " " & Selected!Last_Name
      List2.AddItem Line
      Selected.MoveNext
    Loop
    VScroll1.Enabled = False
  End If
  Selected.Close
End Sub

Private Sub UpdateList3()
Dim RoundNum As Integer
Dim Slot As Integer
Dim P As Integer
Dim num As Integer
Dim Line As String
Dim PCount As Integer

  List3.Clear
  num = 0
  P = 1
  PCount = 0
  For RoundNum = 1 To NumRounds
    Line = "Round " & Str(RoundNum)
    List3.AddItem Line
    For Slot = 1 To F3BNumSlots(ThisTask)
      Do Until alloc(ThisTask, RoundNum, Slot, P) = 0
        num = num + 1
        P = P + 1
      Loop
    PCount = PCount + num
    Line = "Slot" & Str(Slot) & "    " & Str(num)
    List3.AddItem Line
    num = 0
    P = 1
    Next Slot
  P = 1
  Line = "Pilots " & Str(PCount)
  pilotcheck(RoundNum) = PCount
  List3.AddItem Line
  PCount = 0
  Next RoundNum
End Sub

Private Sub Command1_Click()
  If List2.Text = "" Then
    MsgBox ("Nothing Selected")
  Else
    AddPilot
  End If

End Sub

Private Sub Command2_Click()
  If List1.Text = "" Then
    MsgBox ("Nothing Selected")
  Else
    
    RemovePilot
  End If
End Sub

Private Sub Command3_Click()
  'Allocate.Command4.Enabled = True
  'Allocate.Command2.Enabled = True
  'Allocate.Command3.Enabled = True
  Cleanup
  
  UpdateBox
  If roundcheck = True Then
    Unload Me
  End If
End Sub


Private Sub Command4_Click()
  Cleanup
  Unload Me
End Sub

Private Sub Form_Load()
 Text1.Text = 1
 Text2.Text = 1
 UpdateList1
 UpdateList2
 UpdateList3

End Sub

Private Sub VScroll1_Change()
  Dim SQLString As String
  Dim Selected As Recordset
  
  VScroll1.max = NumRounds
  VScroll1.Min = 1
  Text1.Text = VScroll1.Value
  
  If allocated = False Then
    
    UpdateList1
    UpdateList2
    UpdateList3
  Else
    UpdateList1
  End If
  
End Sub


Private Sub VScroll2_Change()
  VScroll2.max = F3BNumSlots(ThisTask)
  VScroll2.Min = 1
  Text2.Text = VScroll2.Value
  UpdateList1
End Sub



Private Sub RemovePilot()
Dim RoundNum As Integer
Dim Slot As Integer
Dim P As Integer
Dim Pos As Integer
Dim Pilot As String
Dim num As Integer
Dim Count As Integer
Dim SQLString As String
Dim Selected As Recordset

  RoundNum = Val(Text1.Text)
  Slot = Val(Text2.Text)
  Pilot = Trim(List1.Text)
  Pilot = Left(Pilot, 3)
  Pilot = Trim(Pilot)
  SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, List.Flyoff FROM list WHERE ((List.Pilot_ID=" & Pilot & ") AND (List.Comp_ID=" & Str(CurrentContest) & "));"
  Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  Selected.Edit
  Selected!Flyoff = True
  Selected.Update
  P = Val(Pilot)
  Pos = 1
  num = 0
  Do Until alloc(ThisTask, RoundNum, Slot, Pos) = 0
    Pos = Pos + 1
    num = num + 1
  Loop
  Pos = 1
  Do Until alloc(ThisTask, RoundNum, Slot, Pos) = P
    Pos = Pos + 1
  Loop
  For Count = Pos To num
    alloc(ThisTask, RoundNum, Slot, Count) = alloc(ThisTask, RoundNum, Slot, Count + 1)
  Next Count
UpdateList1
UpdateList2
UpdateList3
End Sub

Private Sub AddPilot()
Dim RoundNum As Integer
Dim Slot As Integer
Dim P As Integer
Dim Pos As Integer
Dim Pilot As String
Dim num As Integer
Dim Count As Integer
Dim SQLString As String
Dim Selected As Recordset
Dim FirstFreq As Single
Dim CurrentPilotTeam
Dim Change As Boolean
Dim Response As Boolean
Dim Msg As String
Dim SearchStr As String
Dim DebugVar As Double


  Change = True
  RoundNum = Val(Text1.Text)
  Slot = Val(Text2.Text)
  Pilot = Trim(List2.Text)
  Pilot = Left(Pilot, 3)
  Pilot = Trim(Pilot)
  'First get the selected Pilots details
  SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID,List.Team_ID, Pilot.First_Name, Pilot.Last_Name, List.FreqUsed, List.Flyoff FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID WHERE List.Comp_ID=" & Str(CurrentContest) & ";"
  Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  SearchStr = "Pilot_ID =" & Pilot
  Selected.FindFirst SearchStr
  FirstFreq = Selected!FreqUsed
  CurrentPilotTeam = Selected!Team_ID
  
  Count = 1
  'Then go through all the other pilots to check for conflicts
  Do Until alloc(ThisTask, RoundNum, Slot, Count) = 0
    SearchStr = "Pilot_ID =" + Str(alloc(ThisTask, RoundNum, Slot, Count))
    Selected.FindFirst SearchStr
    If Abs(FirstFreq - Selected!FreqUsed) <= 0.011 Then
        If TenK = False Then
          MsgBox ("Cannot allocate Pilot to this slot due to a Frequency Conflict")
          Change = False
          Exit Do 'Freq Conflict discovered break out of loop
        End If
    ElseIf FirstFreq = Selected!FreqUsed Then
          MsgBox ("Cannot allocate Pilot to this slot due to a Frequency Conflict")
          Change = False
          Exit Do 'Freq Conflict discovered break out of loop where 10Kc splits are allowed
     
    ElseIf CurrentPilotTeam = Selected!Team_ID Then 'only check the teams if tries is less than 4
      If Selected!Team_ID <> 0 Then 'If Team_ID > 0 then pilot is not in team
          Msg = "Team Conflict" & Chr(13) & "Allocate Anyway ? "
          If MsgBox(Msg, 19, "Allocate Program") <> vbYes Then
            Change = False
            Exit Do 'Pilot is in same team as others
          End If
        End If
    End If
    Count = Count + 1
  Loop
  Selected.Close
  If Change = True Then 'if there are no conflicts then add to alloc
    Pos = 1
    num = 0
    'Find the last entry in alloc
    Do Until alloc(ThisTask, RoundNum, Slot, Pos) = 0
      Pos = Pos + 1
      num = num + 1
    Loop
    'allocate pilot to alloc
    alloc(ThisTask, RoundNum, Slot, num + 1) = Val(Pilot)
    'set flag off so pilot does not appear in bottom window
    SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, List.Flyoff FROM list WHERE ((List.Pilot_ID=" & Pilot & ") AND (List.Comp_ID=" & Str(CurrentContest) & "));"
    Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    Selected.Edit
    Selected!Flyoff = False
    Selected.Update
    Selected.Close
  End If
  UpdateList1
  UpdateList2
  UpdateList3
End Sub

Public Sub UpdateBox()
Dim RoundNum As Integer
 Dim Slot As Integer
 Dim task As Integer
 Dim P As Integer
 Dim Line As String
 Dim SQLString As String
 Dim Selected As Recordset
 Dim LineC As Integer
 
 roundcheck = True
 'Print the contest out to a form
 AllocateF3B.List1(1).Clear
 AllocateF3B.List1(2).Clear
 AllocateF3B.List1(3).Clear
 
 P = 1
 LineC = 1
 For task = 1 To 3
   LineC = 1
   P = 1
   For RoundNum = 1 To NumRounds
     AllocateF3B.List1(task).AddItem " "
     AllocateF3B.List1(task).AddItem "Round " & Str(RoundNum) & "  " & Str(alloc(task, RoundNum, 0, 0))
     For Slot = 1 To F3BNumSlots(task)
       LineC = 1
       Line = "Slot " & Str(Slot) & "  "
       Do Until alloc(task, RoundNum, Slot, P) = 0
         SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, Pilot.First_Name, Pilot.Last_Name, List.FreqUsed, List.Flyoff FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID WHERE ((List.Pilot_ID=" & Str(alloc(task, RoundNum, Slot, P)) & ") AND (List.Comp_ID=" & Str(CurrentContest) & "));"
         Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
         Line = Line + Left(Selected!First_Name, 1) & "." & Selected!Last_Name & " "
         P = P + 1
         LineC = LineC + 1
         If LineC > 8 Then
           AllocateF3B.List1(task).AddItem Line
           Line = "           "
           LineC = 1
         End If
       Loop
       AllocateF3B.List1(task).AddItem Line
       Line = ""
       P = 1
       LineC = 1
     Next Slot
     P = 1
     LineC = 1
   Next RoundNum
 Next task
End Sub

Public Sub Cleanup()
Dim RoundNum As Integer
Dim Slotnum As Integer
Dim ZeroSlot As Integer
  
  ZeroSlot = 1
  For RoundNum = 1 To NumRounds
    For Slotnum = 1 To F3BNumSlots(ThisTask)
      If alloc(ThisTask, RoundNum, Slotnum, 1) <> 0 Then
        ZeroSlot = ZeroSlot + 1
      End If
    Next Slotnum
  Next RoundNum
  
  If ZeroSlot = NumRounds Then
    F3BNumSlots(ThisTask) = ZeroSlot
  End If
End Sub
