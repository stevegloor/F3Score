VERSION 5.00
Begin VB.Form ViewF3JScores 
   ClientHeight    =   6375
   ClientLeft      =   480
   ClientTop       =   780
   ClientWidth     =   6330
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
   ScaleHeight     =   6375
   ScaleWidth      =   6330
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scores"
      Height          =   3975
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   6015
      Begin VB.ListBox List2 
         Height          =   3420
         Left            =   1920
         TabIndex        =   10
         Top             =   360
         Width           =   3975
      End
      Begin VB.ListBox List1 
         Height          =   3420
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   495
      Left            =   3720
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
      Left            =   3600
      TabIndex        =   0
      Top             =   5880
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
      Left            =   3240
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
      Caption         =   " View Form"
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
Attribute VB_Name = "ViewF3JScores"
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
Dim Sec As Integer
Dim Line As String

  List1.Clear
  List2.Clear
  SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.RawScore FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & ") AND (Scores.Slot=" & Slot & "))ORDER BY Scores.Score DESC;"
  Line = "    Pilot Name"
  List1.AddItem Line
  Line = "Mins" & "   " & "Secs" & "  " & "Land" & "    " & "F Pen" & "     " & "S Pen" & "     " & "Score"
  List2.AddItem Line
  
  Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
  If SelectedSlot.RecordCount <> 0 Then
    SelectedSlot.MoveFirst
    Do Until SelectedSlot.EOF
      If SelectedSlot!Flight_Time <> 0 Then
         Mins = SelectedSlot!Flight_Time \ 60
         Sec = SelectedSlot!Flight_Time Mod 60
      End If
      Line = SelectedSlot!First_Name & " " & SelectedSlot!Last_Name
      List1.AddItem Line
      Line = "  " & Format(Mins, "0#") & "      " & Format(Sec, "0#") & "     " & Format(SelectedSlot!LandBonus, "0##") & "      " & Format(SelectedSlot!Penalty, "0##") & "      " & Format(SelectedSlot!Score, "####.0")
      List2.AddItem Line
      SelectedSlot.MoveNext
      Mins = 0
      Sec = 0
    Loop
  Else
    MsgBox ("No Pilots in Slot")
  End If
End Sub
