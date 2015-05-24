VERSION 5.00
Begin VB.Form GroupMaint 
   Caption         =   "Flight Group Maintenance"
   ClientHeight    =   5865
   ClientLeft      =   1455
   ClientTop       =   600
   ClientWidth     =   5175
   HelpContextID   =   410
   LinkTopic       =   "Form6"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5865
   ScaleWidth      =   5175
   Begin VB.CommandButton Command5 
      Caption         =   "Reset Groups"
      Height          =   495
      Left            =   3480
      TabIndex        =   17
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Set Up Groups"
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   2640
      WhatsThisHelpID =   410
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Pilot From Group"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      WhatsThisHelpID =   410
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Pilot to Group"
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   3600
      WhatsThisHelpID =   410
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pilots Not Allocated"
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   4200
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
      Caption         =   "Pilots in Selected Group"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   410
      Width           =   3255
      Begin VB.VScrollBar VScroll3 
         Height          =   495
         Left            =   2760
         TabIndex        =   14
         Top             =   600
         WhatsThisHelpID =   410
         Width           =   255
      End
      Begin VB.TextBox Text3 
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
         Left            =   2280
         TabIndex        =   13
         Top             =   600
         Width           =   495
      End
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
         Left            =   360
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
         Left            =   1320
         TabIndex        =   3
         Top             =   600
         Width           =   495
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   495
         Left            =   840
         TabIndex        =   2
         Top             =   600
         WhatsThisHelpID =   410
         Width           =   255
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   495
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         WhatsThisHelpID =   410
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Group"
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
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         Width           =   735
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
         Left            =   240
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
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "GroupMaint"
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
Dim Group As Integer
  List1.Clear
  RoundNum = Val(Text1.Text)
  Slot = Val(Text2.Text)
  Group = Val(Text3.Text)
  P = 1
  SQLString = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot,Scores.FlightGroup FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & Str(RoundNum) & ") AND (Scores.Slot=" & Str(Slot) & ") AND (Scores.FlightGroup = " & Str(Group) & ") AND (Scores.Task = 'B'))ORDER BY Pilot.Last_Name;"
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
    VScroll1.Enabled = True
    VScroll2.Enabled = True
    AddPilot
  End If

End Sub

Private Sub Command2_Click()
  If List1.Text = "" Then
    MsgBox ("Nothing Selected")
  Else
    VScroll1.Enabled = False
    VScroll2.Enabled = False
    RemovePilot
  End If
End Sub

Private Sub Command3_Click()
  Unload Me
End Sub


Private Sub Command4_Click()
  SetUpGroups
End Sub

Private Sub Command5_Click()
  DeleteGroups
End Sub

Private Sub Form_Load()
 Text1.Text = 1
 Text2.Text = 1
 Text3.Text = 1
 If FlightGroups = True Then
   Command4.Visible = False
 End If
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
    Selected.Edit
    Selected!FlightGroup = 0
    Selected.Update
  Else
    Selected.MoveFirst
    Do Until Selected.EOF
      Selected.Edit
      Selected!FlightGroup = 0
      Selected.Update
      Selected.MoveNext
    Loop
  End If
  
  UpdateList1
  UpdateList2

End Sub

Private Sub AddPilot()
Dim RoundNum As Integer
Dim Slot As Integer
Dim Group As Integer
Dim P As Integer
Dim Pos As Integer
Dim Pilot As String
Dim num As Integer
Dim Count As Integer
Dim SQLString As String
Dim Selected As Recordset

  
  RoundNum = Val(Text1.Text)
  Slot = Val(Text2.Text)
  Group = Val(Text3.Text)
  Pilot = Trim(List2.Text)
  Pilot = Left(Pilot, 3)
  Pilot = Trim(Pilot)
  SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, List.Flyoff FROM list WHERE ((List.Pilot_ID=" & Pilot & ") AND (List.Comp_ID=" & Str(CurrentContest) & "));"
  
  Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  Selected.Edit
  Selected!Flyoff = False
  Selected.Update
  P = Val(Pilot)
  SQLString = "SELECT * From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Pilot_ID=" & Pilot & ") AND (Scores.Round= " & RoundNum & ") AND (Scores.Slot=" & Slot & "));"
  Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If Selected.RecordCount = 0 Then
    Exit Sub
  ElseIf Selected.RecordCount = 1 Then
    Selected.Edit
    Selected!FlightGroup = Group
    Selected.Update
  Else
    Selected.MoveFirst
    Do Until Selected.EOF
      Selected.Edit
      Selected!FlightGroup = Group
      Selected.Update
      Selected.MoveNext
    Loop
  End If
  
  UpdateList1
  UpdateList2
  End Sub



Private Sub VScroll3_Change()
  VScroll3.Min = 1
  Text3.Text = VScroll3.Value
  UpdateList1
End Sub



Public Sub SetUpGroups()
Dim RoundNum As Integer
Dim Slot As Integer
Dim Group As Integer
Dim NumberInGroups
Dim Pilot As String
Dim num As Integer
Dim Count As Integer
Dim SQLString As String
Dim Selected As Recordset
Dim ChangeContest As Recordset
  
  NumberInGroups = InputBox("How Many Pilots In Each Flight Group")
  If NumberInGroups <> "" Then
    For RoundNum = 1 To NumRounds
      For Slot = 1 To NumSlots
        SQLString = "SELECT * From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Round= " & Trim(Str(RoundNum)) & ") AND (Scores.Slot=" & Trim(Str(Slot)) & ") AND (Scores.Task = 'B'));"
        Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
        Selected.MoveFirst
        Count = 1
        Group = 1
        Do Until Selected.EOF
          Selected.Edit
          Selected!FlightGroup = Group
          Selected.Update
          Selected.MoveNext
          If Count <> NumberInGroups Then
            Count = Count + 1
          Else
            Count = 1
            Group = Group + 1
          End If
        Loop
      Next Slot
    Next RoundNum
    SQLString = "SELECT * From Competition WHERE (Competition.Comp_ID = " & Str(CurrentContest) & ");"
    Set ChangeContest = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    If ChangeContest.RecordCount <> 0 Then
      ChangeContest.Edit
      ChangeContest!FlightGroups = True
      ChangeContest.Update
      FlightGroups = True
    End If
    UpdateList1
    UpdateList2
    Command4.Visible = False
  End If
    
End Sub

Public Sub DeleteGroups()
Dim RoundNum As Integer
Dim Slot As Integer
Dim Group As Integer
Dim NumberInGroups
Dim Pilot As String
Dim num As Integer
Dim Count As Integer
Dim SQLString As String
Dim Selected As Recordset
Dim ChangeContest As Recordset

  
    For RoundNum = 1 To NumRounds
      For Slot = 1 To NumSlots
        SQLString = "SELECT * From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Round= " & Trim(Str(RoundNum)) & ") AND (Scores.Slot=" & Trim(Str(Slot)) & ") AND (Scores.Task = 'B'));"
        Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
        Selected.MoveFirst
        
        Do Until Selected.EOF
          Selected.Edit
          Selected!FlightGroup = 0
          Selected.Update
          Selected.MoveNext
        Loop
      Next Slot
    Next RoundNum
    SQLString = "SELECT * From Competition WHERE (Competition.Comp_ID = " & Str(CurrentContest) & ");"
    Set ChangeContest = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    If ChangeContest.RecordCount <> 0 Then
      ChangeContest.Edit
      ChangeContest!FlightGroups = False
      ChangeContest.Update
      FlightGroups = True
    End If
    Command4.Visible = True
    UpdateList1
    UpdateList2

 
    
End Sub
