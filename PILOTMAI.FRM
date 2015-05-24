VERSION 5.00
Begin VB.Form F3JPilotMaint 
   Caption         =   "Move Pilot"
   ClientHeight    =   5865
   ClientLeft      =   1620
   ClientTop       =   540
   ClientWidth     =   3585
   HelpContextID   =   410
   LinkTopic       =   "Form6"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5865
   ScaleWidth      =   3585
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   5400
      WhatsThisHelpID =   410
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Pilot From Slot"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      WhatsThisHelpID =   410
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Pilot to Slot"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   3600
      WhatsThisHelpID =   410
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pilots Not Allocated"
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      WhatsThisHelpID =   410
      Width           =   3255
      Begin VB.ListBox List2 
         Height          =   840
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pilots in Selected Slot"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   410
      Width           =   3255
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   3015
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
         WhatsThisHelpID =   410
         Width           =   255
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   495
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         WhatsThisHelpID =   410
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
End
Attribute VB_Name = "F3JPilotMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UpdateList1()
Dim RoundNum As Integer
Dim Slot As Integer
Dim P As Integer
Dim SearchStr As String
Dim Selected As Recordset
Dim Line As String
Dim SQLString As String
  List1.Clear
  RoundNum = Val(Text1.Text)
  Slot = Val(Text2.Text)
  P = 1
  SQLString = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & ") AND (Scores.Task = 'A'))ORDER BY Pilot.Last_Name;"
  Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If Selected.RecordCount <> 0 Then
    Do Until Selected.EOF
      Line = Format(Str(Selected!Pilot_ID), "!@@@@") & " " & Selected!First_Name & " " & Selected!Last_Name
      List1.AddItem Line
      Selected.MoveNext
    Loop
    Selected.Close
  Else
    List1.AddItem "No Pilots"
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
      Line = Format(Str(Selected!Pilot_ID), "!@@@@") & Selected!First_Name & " " & Selected!Last_Name
      List2.AddItem Line
      Selected.MoveNext
    Loop
  End If
  Selected.Close
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
    VScroll1.Enabled = False
    RemovePilot
  End If
End Sub

Private Sub Command3_Click()
  Unload Me
End Sub


Private Sub Form_Load()
 Text1.Text = 1
 Text2.Text = 1
 UpdateList1
 UpdateList2
 

End Sub

Private Sub VScroll1_Change()
  VScroll1.max = NumRounds
  VScroll1.Min = 1
  Text1.Text = VScroll1.Value
  UpdateList1
End Sub


Private Sub VScroll2_Change()
  VScroll2.max = NumSlots
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
  SQLString = "SELECT * From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Pilot_ID=" & Pilot & ") AND (Scores.Round= " & RoundNum & ") AND (Scores.Slot=" & Slot & "));"
  Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If Selected.RecordCount = 0 Then
    Exit Sub
  ElseIf Selected.RecordCount = 1 Then
    Selected.Delete
  Else
    Selected.MoveFirst
    Do Until Selected.EOF
      Selected.Delete
      Selected.MoveNext
    Loop
  End If
  UpdateList1
  UpdateList2

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
Dim SelSlot As Recordset
Dim FirstFreq
Dim CurrentPilotTeam
Dim Change As Boolean
Dim Response As Boolean
Dim Msg As String
Dim SearchStr As String
Dim T As Integer


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
  SQLString = "SELECT * From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Round= " & RoundNum & ") AND (Scores.Slot=" & Slot & "));"
  Set SelSlot = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If SelSlot.RecordCount <> 0 Then
    SelSlot.MoveFirst
    Do Until SelSlot.EOF
      SearchStr = "Pilot_ID =" + Str(SelSlot!Pilot_ID)
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
      
      ElseIf CurrentPilotTeam = Selected!Team_ID Then
        If Selected!Team_ID <> 0 Then 'If Team_ID > 0 then pilot is not in team
          Msg = "Team Conflict" & Chr(13) & "Allocate Anyway ? "
          If MsgBox(Msg, 19, "Allocate Program") <> vbYes Then
            Change = False
            Exit Do 'Pilot is in same team as others
          End If
        End If
      End If
      SelSlot.MoveNext
    Loop
    Selected.Close
    SelSlot.Close
    If Change = True Then 'if there are no conflicts then add to alloc
      If CurrentContestType <> "F3B" Then
        'set flag off so pilot does not appear in bottom window
        Set SelSlot = F3JDb.OpenRecordset("Scores", dbOpenTable)
        SelSlot.AddNew
        SelSlot!Comp_ID = CurrentContest
        SelSlot!Round = Val(RoundNum)
        SelSlot!Slot = Val(Slot)
        SelSlot!Pilot_ID = Val(Pilot)
        SelSlot!Task = "A"
        SelSlot.Update
        SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, List.Flyoff FROM list WHERE ((List.Pilot_ID=" & Pilot & ") AND (List.Comp_ID=" & Str(CurrentContest) & "));"
        Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
        Selected.Edit
        Selected!Flyoff = False
        Selected.Update
        Selected.Close
      Else
        'set flag off so pilot does not appear in bottom window
        Set SelSlot = F3JDb.OpenRecordset("Scores", dbOpenTable)
        For T = 1 To 3
          SelSlot.AddNew
          SelSlot!Comp_ID = CurrentContest
          SelSlot!Round = Val(RoundNum)
          SelSlot!Slot = Val(Slot)
          SelSlot!Pilot_ID = Val(Pilot)
          If T = 1 Then
            SelSlot!Task = "A"
          ElseIf T = 2 Then
            SelSlot!Task = "B"
          ElseIf T = 3 Then
            SelSlot!Task = "C"
          End If
          SelSlot.Update
        Next T
        SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, List.Flyoff FROM list WHERE ((List.Pilot_ID=" & Pilot & ") AND (List.Comp_ID=" & Str(CurrentContest) & "));"
        Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
        Selected.Edit
        Selected!Flyoff = False
        Selected.Update
        Selected.Close
      End If
    End If
  Rem SelSlot.Close
  End If
  UpdateList1
  UpdateList2
  End Sub



