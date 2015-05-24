VERSION 5.00
Begin VB.Form ChampMaint 
   Caption         =   "Championship Maintenance Form"
   ClientHeight    =   6240
   ClientLeft      =   525
   ClientTop       =   930
   ClientWidth     =   7020
   LinkTopic       =   "Form4"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6240
   ScaleWidth      =   7020
   Begin VB.Frame Frame3 
      Caption         =   "Select Championship"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton Command9 
         Caption         =   "Pilot Statistics"
         Height          =   375
         Left            =   4800
         TabIndex        =   14
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print Results"
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Caption         =   "View Results"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Delete "
         Height          =   375
         Left            =   4800
         TabIndex        =   11
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add New "
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Edit Details"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DONE"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<< REMOVE"
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD >>"
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contests in Selected Championship"
      Height          =   4095
      Left            =   4080
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
      Begin VB.ListBox List2 
         Height          =   3570
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "List of Contests Available"
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   2535
      Begin VB.ListBox List1 
         Height          =   3570
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "ChampMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub UpdateList1()
Dim TeamSelect As Recordset
Dim SQLString As String
Dim Line As String
 List1.Clear
 SQLString = "SELECT * FROM Competition WHERE ((Competition.ChampID = 0)) ORDER BY Competition.Comp_Name ASC;"
 Set TeamSelect = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
 If TeamSelect.RecordCount <> 0 Then
   TeamSelect.MoveFirst
 End If
 Do Until TeamSelect.EOF = True
   Line = Format(Str(TeamSelect!Comp_ID), "!@@@@") & " " & Format(TeamSelect!CompType, "!@@@@@") & " " & TeamSelect!Comp_Name
   If TeamSelect!Comp_ID = CurrentContest Then
     Line = Line & " **"
   End If
   List1.AddItem Line
   TeamSelect.MoveNext
 Loop
TeamSelect.Close
End Sub

Public Sub UpdateList2()
Dim TeamSelect As Recordset
Dim SQLString As String
Dim Line As String
 List2.Clear
 If CurrentChamp <> 0 Then
   SQLString = "SELECT * FROM Competition WHERE ((Competition.ChampID = " & Str(CurrentChamp) & ")) ORDER BY Comp_Name ASC;"
   Set TeamSelect = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
   If TeamSelect.RecordCount <> 0 Then
     TeamSelect.MoveFirst
   End If
   Do Until TeamSelect.EOF = True
     Line = Format(Str(TeamSelect!Comp_ID), "!@@@@") & " " & TeamSelect!Comp_Name
     If TeamSelect!Comp_ID = CurrentContest Then
       Line = Line & " **"
     End If
     List2.AddItem Line
     SelectedContestType = TeamSelect!CompType
     TeamSelect.MoveNext
   Loop
   
   TeamSelect.Close
 Else
   Line = "No Contests"
   List2.AddItem Line
 End If
End Sub

Private Sub Combo1_Click()
  Dim Selected As String
  
  Selected = Trim(Combo1.Text)
  Selected = Left(Selected, 3)
  CurrentChamp = Int(Trim(Selected))
  Command1.Enabled = True
  Command2.Enabled = True
  UpdateList2
End Sub


Private Sub Combo1_DropDown()
Dim TeamSelect As Recordset
Dim SQLString As String
Dim Line As String
 Combo1.Clear
 SQLString = "SELECT * FROM Championship;"
 Set TeamSelect = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
 If TeamSelect.RecordCount <> 0 Then
   TeamSelect.MoveFirst
 End If
 Do Until TeamSelect.EOF = True
   Line = Format(Str(TeamSelect!ChampID), "!@@@@") & " " & TeamSelect!ChampName
   Combo1.AddItem Line
   TeamSelect.MoveNext
 Loop
TeamSelect.Close
End Sub


Private Sub Command1_Click()
  AddContest
End Sub

Private Sub Command2_Click()
  RemoveContest
End Sub

Private Sub Command3_Click()
  Unload Me
End Sub

Private Sub Command4_Click()
  If CurrentChamp = 0 Then
    MsgBox ("No Championship Selected")
  Else
   FromChange = True
   EditChamp.Caption = "Edit Championship's Details"
   EditChamp.Show
  End If
End Sub

Private Sub Command5_Click()
  FromChange = False
  EditChamp.Caption = "Enter New Championship's Details"
  EditChamp.Show
End Sub


Private Sub Command7_Click()
  ViewResults
End Sub

Private Sub Command9_Click()
  If CurrentChamp <> 0 Then
    PilotStats.Show
  Else
    MsgBox ("No Championship Selected")
  End If
End Sub

Private Sub Form_Load()
  CurrentChamp = 0
  SelectedContestType = ""
  FromChange = False
  Command1.Enabled = False
  Command2.Enabled = False
  UpdateList1
  UpdateList2
End Sub


Private Sub Text1_Change()

End Sub



Public Sub AddContest()
Dim ChampSet As Recordset
Dim Selected As String
 
 Selected = Trim(List1.Text)
 Selected = Left(Selected, 3)
 Selected = Trim(Selected)
 
 If Selected = "" Then
   MsgBox ("Nothing Selected")
 Else
   SQLString = "SELECT * FROM Competition WHERE ((Competition.Comp_ID = " & Selected & "));"
   Set ChampSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
   If ChampSet.RecordCount > 0 Then
     If SelectedContestType = "" Then
       SelectedContestType = ChampSet!CompType
     End If
     If ChampSet!CompType = SelectedContestType Then
      ChampSet.Edit
      ChampSet!ChampID = CurrentChamp
      ChampSet.Update
      SelectedContestType = ChampSet!CompType
     Else
      MsgBox ("This Contest is not the same type as the others")
     End If
   End If
 End If
 UpdateList2
 UpdateList1
 ChampSet.Close
End Sub

Public Sub RemoveContest()
Dim ChampSet As Recordset
Dim Selected As String
 
 Selected = Trim(List2.Text)
 Selected = Left(Selected, 3)
 Selected = Trim(Selected)
 
 If Selected = "" Then
   MsgBox ("Nothing Selected")
 Else
   SQLString = "SELECT * FROM Competition WHERE ((Competition.Comp_ID = " & Selected & "));"
   Set ChampSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
   If ChampSet.RecordCount > 0 Then
     ChampSet.Edit
     ChampSet!ChampID = 0
     ChampSet.Update
   End If
 End If
 UpdateList2
 UpdateList1
 ChampSet.Close
End Sub

Public Sub DeleteChamp()

End Sub

Public Sub ViewResults()
Dim SQLString As String
Dim CompSet As Recordset
Dim Selected As Recordset
Dim I As Integer
Dim TotalChampScore() As Integer
Dim maxnum As Integer
Dim Line As String

  If CurrentChamp <> 0 Then
    SQLString = "SELECT * FROM Pilot ORDER BY Pilot.Pilot_ID DESC;"
    Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    Selected.MoveFirst
    maxnum = Selected!Pilot_ID
    ReDim TotalChampScore(maxnum)
    Currentpilot = 0
    SQLString = "SELECT DISTINCTROW Competition.ChampID, List.FlyoffScore, Competition.Completed, Pilot.Pilot_ID, List.FinalCompScore, Pilot.First_Name, Pilot.Last_Name FROM Pilot INNER JOIN ((Championship INNER JOIN Competition ON Championship.ChampID = Competition.ChampID) INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) ON Pilot.Pilot_ID = List.Pilot_ID Where ((Competition.ChampID = " & Trim(Str(CurrentChamp)) & ") And (Competition.Completed = True)) ORDER BY Pilot.Pilot_ID;"
    Set CompSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    If CompSet.RecordCount = 0 Then
      MsgBox ("There are no complete contests in the Championship")
      
    Else
      If Cummlative = True Then
        CompSet.MoveFirst
        Do Until CompSet.EOF
          TotalChampScore(CompSet!Pilot_ID) = TotalChampScore(CompSet!Pilot_ID) + CompSet!FinalCompScore
          CompSet.MoveNext
        Loop
      Else
        CompSet.MoveFirst
        Do Until CompSet.EOF
          TotalChampScore(CompSet!Pilot_ID) = TotalChampScore(CompSet!Pilot_ID) + Place(CompSet!FlyoffScore)
          CompSet.MoveNext
        Loop
      End If
      Selected.MoveFirst
      Do Until Selected.EOF
        Selected.Edit
        Selected!Res2 = TotalChampScore(Selected!Pilot_ID)
        Selected.Update
        Selected.MoveNext
      Loop
      Selected.Close
      SQLString = "SELECT * FROM Pilot WHERE ((Pilot.Res2 <> 0)) ORDER BY Pilot.Res2 DESC;"
      Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
      Selected.MoveFirst
      GeneralDisplay.Caption = "Championship Standings"
      I = 1
      Do Until Selected.EOF
        Line = Str(Format(I, "0#")) & "  " & Str(Format(Selected!Res2, "0#")) & "    " & Selected!First_Name & " " & Selected!Last_Name
        GeneralDisplay.List1.AddItem Line
        Selected.MoveNext
        I = I + 1
      Loop
      GeneralDisplay.Show
    End If
    Selected.MoveFirst
    ' Do housekeeping
    Do Until Selected.EOF
      If Selected!Res2 <> 0 Then
        Selected.Edit
        Selected!Res2 = 0
        Selected.Update
        Selected.MoveNext
      Else
        Selected.MoveNext
      End If
    Loop
    CompSet.Close
    Selected.Close
  Else
    MsgBox ("No Championship selected")
  End If
End Sub
