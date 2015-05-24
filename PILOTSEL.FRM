VERSION 5.00
Begin VB.Form PilotSelect 
   Caption         =   "Enter Pilots into Contest"
   ClientHeight    =   6030
   ClientLeft      =   435
   ClientTop       =   660
   ClientWidth     =   6720
   HelpContextID   =   500
   LinkTopic       =   "Form6"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6030
   ScaleWidth      =   6720
   Begin VB.CommandButton Command4 
      Caption         =   "Change Selected Pilots Details"
      Height          =   375
      HelpContextID   =   500
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      WhatsThisHelpID =   370
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New Pilot to Database"
      Height          =   375
      HelpContextID   =   500
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      WhatsThisHelpID =   370
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      Caption         =   "DONE"
      Height          =   375
      HelpContextID   =   500
      Left            =   3960
      TabIndex        =   6
      Top             =   5400
      WhatsThisHelpID =   370
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pilots in Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   960
      WhatsThisHelpID =   370
      Width           =   2415
      Begin VB.ListBox List1 
         Height          =   3765
         HelpContextID   =   500
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pilots in Selected Contest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   3960
      TabIndex        =   2
      Top             =   960
      WhatsThisHelpID =   370
      Width           =   2535
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.ListBox List2 
         Height          =   3570
         HelpContextID   =   500
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Number of Pilots"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   240
         WhatsThisHelpID =   370
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD >> "
      Height          =   495
      HelpContextID   =   500
      Left            =   2640
      TabIndex        =   1
      Top             =   2520
      WhatsThisHelpID =   370
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<< REMOVE"
      Height          =   495
      HelpContextID   =   500
      Left            =   2640
      TabIndex        =   0
      Top             =   3600
      WhatsThisHelpID =   370
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Select Pilots for Contest"
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
      Left            =   1080
      TabIndex        =   7
      Top             =   120
      WhatsThisHelpID =   370
      Width           =   4215
   End
End
Attribute VB_Name = "PilotSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub VScroll1_Change()

End Sub


Private Sub Command1_Click()
Dim PilotNum As String
Dim SQLString As String
Dim AddSelect As Recordset
Dim PilotList As Recordset

  PilotNum = Trim(List1.Text)
  PilotNum = Left(PilotNum, 3)
  PilotNum = Trim(PilotNum)
  If PilotNum = "" Then
    MsgBox ("Nothing Selected")
  Else
    SQLString = "SELECT * FROM List WHERE ((List.Pilot_ID =" & PilotNum & ") AND (List.Comp_ID =" & Str(CurrentContest) & "));"
    Set AddSelect = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    If AddSelect.RecordCount = 0 Then
      AddSelect.AddNew
      AddSelect!Comp_ID = CurrentContest
      AddSelect!Pilot_ID = PilotNum
      AddSelect!Team_ID = 0
      AddSelect!InTeam = False
      Randomize
      AddSelect!Rnd1 = Int((400 * Rnd) + 1)  ' Generate random value between 1 and 400.
      Randomize
      AddSelect!Rnd2 = Int((400 * Rnd) + 1)  ' Generate random value between 1 and 400.
      SQLString = "SELECT * FROM Pilot WHERE (Pilot_ID=" & PilotNum & ");"
      Set PilotList = F3JDb.OpenRecordset(SQLString, dbOpenSnapshot)
      AddSelect!FreqUsed = PilotList!Freq1
      AddSelect.Update
      'UpdateBox1
      UpdateBox2
      AddSelect.Close
    Else
      MsgBox ("Duplicate Entry - Check the pilots in the contest")
    End If
  End If
End Sub

Private Sub Command2_Click()
Dim PilotNum As String
Dim SQLString As String
Dim AddSelect As Recordset

  PilotNum = Trim(List2.Text)
  PilotNum = Left(PilotNum, 3)
  PilotNum = Trim(PilotNum)
  If PilotNum = "" Then
    MsgBox ("Nothing Selected")
  Else
    SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, List.Team_ID, List.InTeam From List WHERE ((List.Comp_ID=" & Str(CurrentContest) & ") AND (List.Pilot_ID=" & PilotNum & "));"
    Set AddSelect = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    AddSelect.Delete
    UpdateBox1
    UpdateBox2
    AddSelect.Close
  End If
End Sub

Private Sub Command3_Click()
  PilotEntry.Show
  UpdateBox1
End Sub



Public Sub UpdateBox1()
Dim TeamSelect As Recordset
Dim SQLString As String
Dim Line As String
 List1.Clear
 SQLString = "SELECT * FROM Pilot WHERE Current = True ORDER BY Pilot.Last_Name;"
 Set TeamSelect = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
 If TeamSelect.RecordCount <> 0 Then
   TeamSelect.MoveFirst
 End If
 Do Until TeamSelect.EOF = True
   Line = Format(Str(TeamSelect!Pilot_ID), "!@@@@") & "  " & TeamSelect!First_Name & " " & TeamSelect!Last_Name
   List1.AddItem Line
   TeamSelect.MoveNext
 Loop
TeamSelect.Close
End Sub

Private Sub UpdateBox2()
Dim TeamSelect As Recordset
Dim SQLString As String
Dim Line As String

 List2.Clear
 SQLString = "SELECT DISTINCTROW Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, List.InTeam, List.Team_ID, List.Comp_ID FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID WHERE ((List.Comp_ID=" & Str(CurrentContest) & ")) ORDER BY Pilot.Last_Name;"
 Set TeamSelect = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
 If TeamSelect.RecordCount <> 0 Then
   TeamSelect.MoveFirst
 End If
   
 Do Until TeamSelect.EOF = True
   Line = Format(Str(TeamSelect!Pilot_ID), "!@@@@") & "  " & TeamSelect!First_Name & " " & TeamSelect!Last_Name
   List2.AddItem Line
   TeamSelect.MoveNext
 Loop
 Text1.Text = Str(TeamSelect.RecordCount)
TeamSelect.Close
End Sub

Private Sub Command4_Click()
  Dim PilotNum As String
  
  PilotNum = Trim(List1.Text)
  PilotNum = Left(PilotNum, 3)
  PilotNum = Trim(PilotNum)
  
  If PilotNum = "" Then
    MsgBox ("Nothing Selected")
  Else
    EditPilotID = PilotNum
    FromChange = True
    PilotEntry.Show
  End If
End Sub

Private Sub Command6_Click()
  Unload Me
End Sub

Private Sub Form_Load()
 UpdateBox1
 UpdateBox2
End Sub


