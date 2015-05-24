VERSION 5.00
Begin VB.Form TeamSelect 
   Caption         =   "Team Select Form"
   ClientHeight    =   6030
   ClientLeft      =   270
   ClientTop       =   675
   ClientWidth     =   6375
   HelpContextID   =   360
   LinkTopic       =   "Form6"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6030
   ScaleWidth      =   6375
   Begin VB.CommandButton Command2 
      Caption         =   "<< REMOVE"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   3840
      WhatsThisHelpID =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD >>"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   2880
      WhatsThisHelpID =   360
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pilots in Team"
      Height          =   3135
      Left            =   3960
      TabIndex        =   7
      Top             =   1920
      WhatsThisHelpID =   360
      Width           =   2175
      Begin VB.ListBox List2 
         Height          =   2595
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   1440
      WhatsThisHelpID =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   1440
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pilots in Contest not in Teams"
      Height          =   3975
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      WhatsThisHelpID =   360
      Width           =   2415
      Begin VB.ListBox List1 
         Height          =   3375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "DONE"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   5280
      WhatsThisHelpID =   360
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Team"
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
      Left            =   4560
      TabIndex        =   6
      Top             =   1080
      WhatsThisHelpID =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   " Team Select Form"
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
      Left            =   840
      TabIndex        =   1
      Top             =   240
      WhatsThisHelpID =   360
      Width           =   4215
   End
End
Attribute VB_Name = "TeamSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click()

End Sub

Private Sub Command1_Click()
Dim PilotNum As String
Dim SQLString As String
Dim AddSelect As Recordset

  PilotNum = Trim(List1.Text)
  PilotNum = Left(PilotNum, 3)
  PilotNum = Trim(PilotNum)
  If PilotNum = "" Then
    MsgBox ("Nothing Selected")
  Else
    SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, List.Team_ID, List.InTeam From List WHERE ((List.Comp_ID=" & Str(CurrentContest) & ") AND (List.Pilot_ID=" & PilotNum & "));"
    Set AddSelect = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    AddSelect.Edit
    AddSelect!Team_ID = Text1.Text
    AddSelect!InTeam = True
    AddSelect.Update
    UpdateBox1
    UpdateBox2
    AddSelect.Close
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
    AddSelect.Edit
    AddSelect!Team_ID = 0
    AddSelect!InTeam = False
    AddSelect.Update
    UpdateBox1
    UpdateBox2
    AddSelect.Close
  End If
  
End Sub


Private Sub Command6_Click()
  Unload Me
End Sub


Public Sub Form_Load()
  Text1.Text = 1
  UpdateBox1
  UpdateBox2
End Sub



Public Sub UpdateBox1()
Dim TeamSelect As Recordset
Dim SQLString As String
Dim Line As String
 List1.Clear
 SQLString = "SELECT DISTINCTROW Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, List.InTeam, List.Team_ID, List.Comp_ID FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID WHERE ((List.InTeam=False) AND (List.Comp_ID=" & Str(CurrentContest) & ")) ORDER BY Pilot.Last_Name;"
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

Public Sub UpdateBox2()
Dim TeamMemSel As Recordset
Dim SQLString As String
Dim Line As String
Dim TeamNum As String
 List2.Clear
 TeamNum = Text1.Text
 SQLString = "SELECT DISTINCTROW Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, List.InTeam, List.Team_ID, List.Comp_ID FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID WHERE ((List.InTeam=True) AND (List.Team_ID=" & TeamNum & ") AND (List.Comp_ID=" & Str(CurrentContest) & "))ORDER BY Pilot.Last_Name;"
 Set TeamMemSel = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
 
 If TeamMemSel.RecordCount <> 0 Then
   TeamMemSel.MoveFirst
 End If
 
 Do Until TeamMemSel.EOF = True
   Line = Format(Str(TeamMemSel!Pilot_ID), "!@@@@") & " " & TeamMemSel!First_Name & " " & TeamMemSel!Last_Name
   List2.AddItem Line
   TeamMemSel.MoveNext
 Loop
 TeamMemSel.Close
End Sub

Private Sub VScroll1_Change()
 VScroll1.max = 30
 VScroll1.Min = 1
 Text1.Text = VScroll1.Value
 UpdateBox2
End Sub


