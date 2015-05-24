VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   Caption         =   "Report Selection"
   ClientHeight    =   5235
   ClientLeft      =   780
   ClientTop       =   975
   ClientWidth     =   4455
   HelpContextID   =   340
   LinkTopic       =   "Form3"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5235
   ScaleWidth      =   4455
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   3360
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Output Report as HTML"
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   4560
      WhatsThisHelpID =   340
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Report"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   480
      WhatsThisHelpID =   340
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Report Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   360
      WhatsThisHelpID =   340
      Width           =   2415
      Begin VB.OptionButton Option1 
         Caption         =   "Round/Slot Allocation "
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Contest Flight Groups"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   10
         Top             =   4320
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Full Contest Summary"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   9
         Top             =   3840
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Scores for a Round"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Teams in Contest"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Top             =   2880
         WhatsThisHelpID =   340
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All Pilots in Database"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   3360
         WhatsThisHelpID =   340
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pilots in Contest"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   2400
         WhatsThisHelpID =   340
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Contest Leader Board"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         WhatsThisHelpID =   340
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Round/Slot Allocation with Landing Spots"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   840
         WhatsThisHelpID =   340
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Selected As Recordset
Dim ReportType As Integer
Dim Pagelength As Integer
Dim OutputFile As String

Public Sub PrintContest2()
Dim SQLString3 As String
Dim SelectedSlot As Recordset
Dim RoundNum As Integer
Dim Slot As Integer
Dim Line As String
Dim tabset(7) As Integer
Dim Linecount As Integer
Dim Col As Integer
Dim P As Integer

tabset(0) = 1
tabset(1) = 30
tabset(2) = 60
tabset(3) = 90
tabset(4) = 120
tabset(5) = 150
tabset(6) = 180


 Col = 0

 Printer.FontSize = 16
 Printer.CurrentX = Printer.ScaleWidth / 2 - TextWidth("Round/Slot Allocation")
 Printer.Print "Round/Slot Allocation"
 Linecount = 4
 For RoundNum = 1 To NumRounds
   Printer.FontSize = 14
   Printer.CurrentY = TextHeight("sample") * Linecount
   Printer.Print "Round "; RoundNum
   Linecount = Linecount + 2
   For Slot = 1 To NumSlots
     Printer.FontSize = 12
     Printer.CurrentY = TextHeight("sample") * Linecount
     Printer.Print "Slot "; Slot
     Linecount = Linecount + 2
     Printer.FontSize = 10
     'SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & Str(Round) & ") AND (Scores.Slot=" & Str(Slot) & "))ORDER BY Pilot.Last_Name;"
     'SQLString3 = "SELECT DISTINCTROW List.Comp_ID, Pilot.First_Name, Pilot.Last_Name, List.Pilot_ID, Scores.Round, Scores.Slot, List.FreqUsed, List.Team_ID FROM Pilot INNER JOIN ((Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID)) ON (Pilot.Pilot_ID = List.Pilot_ID) AND (Pilot.Pilot_ID = Scores.Pilot_ID) Where ((List.Comp_ID = " & Str(CurrentContest) & ") And (Scores.Round = " & Str(Round) & ") And (Scores.Slot = " & Str(Slot) & ")) ORDER BY Pilot.Last_Name ASC;"
     SQLString3 = "SELECT DISTINCTROW List.Comp_ID, Pilot.First_Name, Pilot.Last_Name, Pilot.res1, List.Pilot_ID, Scores.Round, Scores.Slot, Scores.Task, List.FreqUsed, List.Team_ID FROM Pilot INNER JOIN ((Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID)) ON (Pilot.Pilot_ID = List.Pilot_ID) AND (Pilot.Pilot_ID = Scores.Pilot_ID) Where ((List.Comp_ID = " & Str(CurrentContest) & ") And (Scores.Round = " & Str(RoundNum) & ") And (Scores.Slot = " & Str(Slot) & ") AND (Scores.Task = '" & Currenttask & "')) ORDER BY Pilot.Last_Name ASC;"
     Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
     If SelectedSlot.RecordCount <> 0 Then
       SelectedSlot.MoveFirst
       Do Until SelectedSlot.EOF
           Printer.CurrentY = TextHeight("sample") * Linecount
           Printer.CurrentX = TextWidth("g") * tabset(Col)
           If SelectedSlot!Res1 <> "null" Then
             Printer.Print SelectedSlot!Res1; " "; Left(SelectedSlot!First_Name, 1); "."; SelectedSlot!Last_Name; " "; Format(Str(SelectedSlot!FreqUsed), "#.000"); " "
           Else
             Printer.Print "    "; Left(SelectedSlot!First_Name, 1); "."; SelectedSlot!Last_Name; " "; Format(Str(SelectedSlot!FreqUsed), "#.000"); " "
           End If
           SelectedSlot.MoveNext
           
           If Col = 3 Then
             If Linecount < Pagelength Then
               Linecount = Linecount + 2
             Else
              Printer.NewPage
              Linecount = 5
             End If
             Col = 0
           Else
             Col = Col + 1
           End If
       Loop
       P = 1
       Col = 0
       If Linecount < Pagelength Then
         Linecount = Linecount + 2
       Else
         Printer.NewPage
         Linecount = 5
       End If
     End If
   Next Slot
   P = 1
   Col = 0
   If Linecount < Pagelength Then
     Linecount = Linecount + 1
   Else
     Printer.NewPage
     Linecount = 5
   End If
 Next RoundNum

Printer.EndDoc

End Sub
Private Sub Check1_Click(Index As Integer)
  ReportType = Index
  
End Sub


Private Sub Command1_Click()
    
  Select Case ReportType
  
  Case 0
    If CurrentContestType = "Club" Then
      PrintPilotsContest
    ElseIf CurrentContestType = "F3B" Then
      Currenttask = InputBox("Enter the Task that you want printed" + Chr$(13) + "Enter A, B or C")
      If Currenttask = "A" Or Currenttask = "a" Then
        Currenttask = "A"
        NumSlots = F3BNumSlots(1)
      ElseIf Currenttask = "B" Or Currenttask = "b" Then
        Currenttask = "B"
        NumSlots = F3BNumSlots(2)
      Else
        Currenttask = "C"
        NumSlots = F3BNumSlots(3)
      End If
      PrintContest
    Else
      Currenttask = "A"
      PrintContest
    End If
  Case 1
    PrintLeaderBoard
  Case 2
    PrintPilotsContest
  Case 3
    PrintContestTeams
  Case 4
    PrintPilots
  Case 5
    PrintRound
  Case 6
    PrintSummary
  Case 7
    PrintFlightGroups
  Case 8
    If CurrentContestType = "Club" Then
      PrintPilotsContest
    ElseIf CurrentContestType = "F3B" Then
      Currenttask = InputBox("Enter the Task that you want printed" + Chr$(13) + "Enter A, B or C")
      If Currenttask = "A" Or Currenttask = "a" Then
        Currenttask = "A"
        NumSlots = F3BNumSlots(1)
      ElseIf Currenttask = "B" Or Currenttask = "b" Then
        Currenttask = "B"
        NumSlots = F3BNumSlots(2)
      Else
        Currenttask = "C"
        NumSlots = F3BNumSlots(3)
      End If
      PrintContest2
    Else
      Currenttask = "A"
      PrintContest2
    End If
   
  End Select
  

End Sub

Private Sub Command2_Click()

  
  Select Case ReportType
    Case 0
      MsgBox ("Not printable as HTML")
    Case 1
      CommonDialog2.Filter = "HTML Files (*.htm)|*.htm"
      CommonDialog2.FilterIndex = 2
      CommonDialog2.ShowSave
      OutputFile = CommonDialog2.FileName
      If OutputFile = "" Then
      MsgBox ("no file selected")
      Exit Sub
      End If
      PrintLeaderBoardHTML
    Case 2
      CommonDialog2.Filter = "HTML Files (*.htm)|*.htm"
      CommonDialog2.FilterIndex = 2
      CommonDialog2.ShowSave
      OutputFile = CommonDialog2.FileName
      If OutputFile = "" Then
        MsgBox ("no file selected")
        Exit Sub
      End If
      PrintPilotsContestHTML
    Case 3
      CommonDialog2.Filter = "HTML Files (*.htm)|*.htm"
      CommonDialog2.FilterIndex = 2
      CommonDialog2.ShowSave
      OutputFile = CommonDialog2.FileName
      If OutputFile = "" Then
        MsgBox ("no file selected")
        Exit Sub
      End If
      PrintContestTeamsHTML
    Case 4
      MsgBox ("Not printable as HTML")
    Case 5
      CommonDialog2.Filter = "HTML Files (*.htm)|*.htm"
      CommonDialog2.FilterIndex = 2
      CommonDialog2.ShowSave
      OutputFile = CommonDialog2.FileName
      If OutputFile = "" Then
        MsgBox ("no file selected")
        Exit Sub
      End If
      PrintRoundHTML
    Case 6
      CommonDialog2.Filter = "HTML Files (*.htm)|*.htm"
      CommonDialog2.FilterIndex = 2
      CommonDialog2.ShowSave
      OutputFile = CommonDialog2.FileName
      If OutputFile = "" Then
        MsgBox ("no file selected")
        Exit Sub
      End If
      PrintSummaryHTML
    Case 7
      MsgBox ("Not printable as HTML")
  End Select
  
End Sub

Private Sub Command3_Click()
  Unload Me
  
End Sub


Private Sub Form_Load()
  Pagelength = 75
End Sub

Private Sub Option1_Click(Index As Integer)
  ReportType = Index
  
End Sub



Public Sub PrintContest()
Dim SQLString3 As String
Dim SelectedSlot As Recordset
Dim RoundNum As Integer
Dim Slot As Integer
Dim NextLine As Integer
Dim Line As String
Dim tabset(15) As Integer
Dim Linecount As Integer
Dim Col As Integer
Dim P As Integer
Dim Spots() As Integer
Dim CurrSpot As Integer
Dim MaxSpot As Integer

tabset(0) = 1
tabset(1) = 15
tabset(2) = 30
tabset(3) = 45
tabset(4) = 60
tabset(5) = 75
tabset(6) = 90
tabset(7) = 105
tabset(8) = 120
tabset(9) = 135
tabset(10) = 150
tabset(11) = 165
tabset(12) = 180
  
 'On Error GoTo Errorhandler
 Col = 0
 ReDim Spots(15)
 Printer.FontSize = 16
 Printer.Orientation = 2
 Pagelength = 49
 Printer.CurrentX = Printer.ScaleWidth / 2 - TextWidth("Round/Slot Allocation")
 Printer.Print "Round/Slot Allocation"
 Linecount = 4
 For RoundNum = 1 To NumRounds
   Printer.FontSize = 14
   Printer.CurrentY = TextHeight("sample") * Linecount
   Printer.Print "Round "; RoundNum
   Linecount = Linecount + 2
   For Slot = 1 To NumSlots
     Printer.FontSize = 12
     Printer.CurrentY = TextHeight("sample") * Linecount
     Printer.Print "Slot "; Slot
     Linecount = Linecount + 2
     Printer.FontSize = 8
     'SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & Str(RoundNum) & ") AND (Scores.Slot=" & Str(Slot) & "))ORDER BY Pilot.Last_Name;"
     SQLString3 = "SELECT DISTINCTROW List.Comp_ID, Pilot.First_Name, Pilot.Last_Name, List.Pilot_ID, Scores.Round, Scores.Slot, Scores.Task, List.FreqUsed, List.Team_ID FROM Pilot INNER JOIN ((Competition INNER JOIN List ON Competition.Comp_ID = List.Comp_ID) INNER JOIN Scores ON (Competition.Comp_ID = Scores.Comp_ID) AND (List.Pilot_ID = Scores.Pilot_ID)) ON (Pilot.Pilot_ID = List.Pilot_ID) AND (Pilot.Pilot_ID = Scores.Pilot_ID) Where ((List.Comp_ID = " & Str(CurrentContest) & ") And (Scores.Round = " & Str(RoundNum) & ") And (Scores.Slot = " & Str(Slot) & ") AND (Scores.Task = '" & Currenttask & "')) ORDER BY List.Team_ID ASC;"
     'Debug.Print SQLString3
     Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
     If SelectedSlot.RecordCount <> 0 Then
       SelectedSlot.MoveFirst
       Do Until SelectedSlot.EOF
         If SelectedSlot.RecordCount > 13 Then
           MsgBox ("Use Print Pilots in Contest to print this task")
           Printer.KillDoc
           Exit Sub
         End If
         NextLine = Linecount + 1
         If SelectedSlot!Team_ID = 0 Then
           CurrSpot = 0
           Do Until Spots(CurrSpot) = 0 'Find a vacant spot
             CurrSpot = CurrSpot + 1
           Loop
         Else
           If SelectedSlot!Team_ID > 11 Then
             CurrSpot = SelectedSlot!Team_ID Mod 11
           Else
             CurrSpot = SelectedSlot!Team_ID
           End If
           If Spots(CurrSpot) = 1 Then 'if the team spot has already been allocated
             CurrSpot = 0              'then go to any slot
             Do Until Spots(CurrSpot) = 0 'Find a vacant spot
               CurrSpot = CurrSpot + 1
             Loop
           End If
         End If
         Printer.CurrentY = TextHeight("sample") * Linecount
         Printer.CurrentX = TextWidth("g") * tabset(CurrSpot)
         Printer.Print "  Spot " & Str(CurrSpot + 1)
         'Debug.Print CurrSpot
         Printer.CurrentY = TextHeight("sample") * NextLine
         Printer.CurrentX = TextWidth("g") * tabset(CurrSpot)
         Printer.Print Left(Trim(SelectedSlot!First_Name), 1); "."; Left(SelectedSlot!Last_Name, 10)
         'Debug.Print Left(Trim(SelectedSlot!First_Name), 1); "."; Left(SelectedSlot!Last_Name, 10)
         SelectedSlot.MoveNext
         If CurrSpot > MaxSpot Then
           MaxSpot = CurrSpot
         End If
         Spots(CurrSpot) = 1
         Col = Col + 1
       Loop
       For CurrSpot = 0 To MaxSpot 'Fill in the blank spots
         If Spots(CurrSpot) = 0 Then
           Printer.CurrentY = TextHeight("sample") * Linecount
           Printer.CurrentX = TextWidth("g") * tabset(CurrSpot)
           Printer.Print "  Spot " & Str(CurrSpot + 1)
         End If
       Next CurrSpot
       P = 1
       Col = 0
       If Linecount < Pagelength Then
         Linecount = Linecount + 3
       Else
         Printer.NewPage
         Linecount = 5
       End If
       ReDim Spots(15)
       MaxSpot = 0
     End If
   Next Slot
   P = 1
   Col = 0
   If Linecount < Pagelength Then
      Linecount = Linecount + 2
   Else
      Printer.NewPage
      Linecount = 5
   End If
   Printer.NewPage
   Printer.EndDoc
   Printer.FontSize = 16
   Printer.Orientation = 2
   Pagelength = 49
   Linecount = 5
 Next RoundNum

Printer.EndDoc
'SelectedSlot.Close
Exit Sub
Errorhandler:
  MsgBox ("You have not entered any pilots into your contest or you have not allocated the rounds")
End Sub

Public Sub PrintLeaderBoard()
Dim Listset As Recordset
Dim SQLString1 As String
Dim rounds As Integer
Dim sco() As Single
Dim Count As Integer
Dim NumPilots As Integer
Dim tabset(20) As Integer
Dim Linecount As Integer
Dim Col As Integer
Dim P As Integer
Dim I As Integer

  tabset(0) = 1
  tabset(1) = 8
  tabset(2) = 45
  tabset(3) = 55
  tabset(4) = 65
  tabset(5) = 75
  tabset(6) = 85
  tabset(7) = 95
  tabset(8) = 105
  tabset(9) = 115
  tabset(10) = 125
  tabset(11) = 135
  tabset(12) = 145
  tabset(13) = 155
  tabset(14) = 165
  tabset(15) = 175
  tabset(16) = 185
  tabset(17) = 195
  tabset(18) = 205
  
  rounds = NumRounds
  Call ScoreAddUp(sco(), NumPilots)
  Count = 1
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Pilot.Res1 FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & ")) ORDER BY List.TempScore DESC;"
  Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Listset.MoveFirst
  If rounds > 7 Then
    Printer.Orientation = 2
    Pagelength = 52
  Else
    Printer.Orientation = 1
    Pagelength = 75
  End If
  Printer.FontSize = 16
  Printer.CurrentX = Printer.ScaleWidth / 2 - TextWidth("Leader Board")
  Printer.Print "Leader Board"
  Linecount = 4
  Printer.FontSize = 12
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.Print "Place"
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.CurrentX = TextWidth("g") * (tabset(1) + 7)
  Printer.Print "Name"
  Printer.CurrentX = TextWidth("g") * tabset(2)
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.Print "Total"
  For I = 3 To NumRounds + 2
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.CurrentX = TextWidth("g") * tabset(I)
    Printer.Print "R"; (I - 2)
  Next I
  Linecount = Linecount + 1
  Printer.CurrentY = TextHeight("sample") * Linecount
  'Print out the information
  Linecount = Linecount + 1
  Count = 1
  Do Until Listset.EOF
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.CurrentX = TextWidth("g") * tabset(0)
    Printer.Print Count
    Printer.CurrentX = TextWidth("g") * tabset(1)
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.Print Trim(Listset!First_Name) & " " & Trim(Listset!Last_Name) & " " & Trim(Listset!Res1)
    Printer.CurrentX = TextWidth("g") * tabset(2)
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.Print Format(sco(Listset!Pilot_ID, 0), "#.0")
    Col = 3
    For I = 1 To rounds
      Printer.CurrentX = TextWidth("g") * tabset(Col)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print Format(sco(Listset!Pilot_ID, I), "#.0")
      Col = Col + 1
    Next I
    Listset.MoveNext
    Count = Count + 1
    If Linecount < Pagelength Then
      Linecount = Linecount + 2
    Else
      Printer.NewPage
      Linecount = 2
      Printer.FontSize = 12
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print "Place"
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.CurrentX = TextWidth("g") * (tabset(1) + 7)
      Printer.Print "Name"
      Printer.CurrentX = TextWidth("g") * tabset(2)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print "Total"
      For I = 3 To NumRounds + 2
        Printer.CurrentY = TextHeight("sample") * Linecount
        Printer.CurrentX = TextWidth("g") * tabset(I)
        Printer.Print "R"; (I - 2)
      Next I
      Linecount = Linecount + 2
    
    End If
  Loop
  
  
  
 SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, List.TempScore From List WHERE (List.Comp_ID=" & Str(CurrentContest) & ") ORDER BY List.Pilot_ID;"
 Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
 Listset.MoveFirst
 Do Until Listset.EOF
   Listset.Edit
   Listset!TempScore = 0
   Listset.Update
   Listset.MoveNext
 Loop
 Printer.EndDoc
 Listset.Close
End Sub

Public Sub PrintPilotsContest()
Dim PilotSet As Recordset
Dim SQLString1 As String
Dim Tab1 As Integer
Dim tabset(10) As Integer
Dim NumPilots As Integer
Dim Linecount As Integer
Dim Col As Integer
Dim P As Integer
Dim Count As Integer

On Error GoTo Errorhandler
  
  tabset(0) = 1
  tabset(1) = 15
  tabset(2) = 35
  tabset(3) = 45
  tabset(4) = 55
  tabset(5) = 65
  tabset(6) = 75
  tabset(7) = 85
  tabset(8) = 95
  tabset(9) = 105


  
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID,List.FreqUsed, List.Team_ID, Pilot.First_Name, Pilot.Last_Name, Pilot.Freq1, Pilot.Freq2, Pilot.Number, Pilot.Res1 FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & ")) ORDER BY Pilot.Last_Name ASC;"
  Set PilotSet = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  PilotSet.MoveFirst
  Printer.FontSize = 16
  Printer.CurrentX = Printer.ScaleWidth / 2 - TextWidth("Leader Board")
  Printer.Print "List of Pilots in Contest"
  Linecount = 4
  Printer.FontSize = 14
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.Print "Number"
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.CurrentX = TextWidth("g") * (tabset(1) + 8)
  Printer.Print "Name"
  Printer.CurrentX = TextWidth("g") * tabset(6)
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.Print "Freq Used"
  Printer.FontSize = 12
  Printer.Print ""
  Linecount = Linecount + 1
  Printer.CurrentY = TextHeight("sample") * Linecount
  
  'Print out the information
  Linecount = Linecount + 1
  Count = 1
  Do Until PilotSet.EOF
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.CurrentX = TextWidth("g") * tabset(0)
    Printer.Print PilotSet!Pilot_ID
    Printer.CurrentX = TextWidth("g") * tabset(1)
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.Print Trim(PilotSet!First_Name) & " " & Trim(PilotSet!Last_Name) & " " & Trim(PilotSet!Res1)
    Printer.CurrentX = TextWidth("g") * tabset(6)
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.Print Format(PilotSet!FreqUsed, "##.000")
 
    PilotSet.MoveNext
    Count = Count + 1
    If Linecount < Pagelength Then
      Linecount = Linecount + 2
    Else
      Printer.NewPage
      Linecount = 2
    End If
  Loop
  Printer.EndDoc
  PilotSet.Close
  Exit Sub
  
Errorhandler:
  MsgBox ("You have not entered any pilots into your contest or you have not allocated the rounds")
End Sub


Public Sub PrintPilots()
Dim PilotSet As Recordset
Dim SQLString1 As String
Dim Tab1 As Integer
Dim tabset(10) As Integer
Dim NumPilots As Integer
Dim Linecount As Integer
Dim Col As Integer
Dim P As Integer
Dim Count As Integer

  tabset(0) = 1
  tabset(1) = 15
  tabset(2) = 35
  tabset(3) = 45
  tabset(4) = 55
  tabset(5) = 65
  tabset(6) = 75
  tabset(7) = 85
  tabset(8) = 95
  tabset(9) = 105


  
  SQLString1 = "SELECT * FROM Pilot ORDER BY Pilot.Last_Name ASC;"
  Set PilotSet = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  PilotSet.MoveFirst
  Printer.FontSize = 16
  Printer.CurrentX = Printer.ScaleWidth / 2 - TextWidth("List of Pilots in Database")
  Printer.Print "List of Pilots in DataBase"
  Linecount = 4
  Printer.FontSize = 12
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.Print "Number"
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.CurrentX = TextWidth("g") * (tabset(1) + 8)
  Printer.Print "Name"
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.CurrentX = TextWidth("g") * (tabset(3) + 4)
  Printer.Print "Address"
  Printer.CurrentX = TextWidth("g") * tabset(6)
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.Print "Club"
  Printer.CurrentX = TextWidth("g") * tabset(7)
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.Print "Freq1"
  Printer.CurrentX = TextWidth("g") * tabset(8)
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.Print "Freq2"
  Printer.CurrentX = TextWidth("g") * tabset(9)
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.Print "Curr."
  
  Printer.FontSize = 10
  Printer.Print ""
  Linecount = Linecount + 1
  Printer.CurrentY = TextHeight("sample") * Linecount
  
  'Print out the information
  Linecount = Linecount + 1
  Count = 1
  Do Until PilotSet.EOF
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.CurrentX = TextWidth("g") * tabset(0)
    Printer.Print PilotSet!Number
    Printer.CurrentX = TextWidth("g") * tabset(1)
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.Print Trim(PilotSet!First_Name) & " " & Trim(PilotSet!Last_Name)
    If PilotSet!Address1 <> "none" Then
      Printer.CurrentX = TextWidth("g") * tabset(3)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print Trim(PilotSet!Address1) & "  " & Trim(PilotSet!Address2)
    End If
    If PilotSet!Club <> "none" Then
      Printer.CurrentX = TextWidth("g") * tabset(6)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print PilotSet!Club
    End If
    Printer.CurrentX = TextWidth("g") * tabset(7)
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.Print PilotSet!Freq1
    Printer.CurrentX = TextWidth("g") * tabset(8)
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.Print PilotSet!Freq2
    Printer.CurrentX = TextWidth("g") * tabset(9)
    Printer.CurrentY = TextHeight("sample") * Linecount
    If PilotSet!current = True Then
      Printer.Print "Yes"
    Else
      Printer.Print "No"
    End If
    PilotSet.MoveNext
    Count = Count + 1
    If Linecount < Pagelength Then
      Linecount = Linecount + 2
    Else
      Printer.NewPage
      Linecount = 2
    End If
  Loop
  Printer.EndDoc
  PilotSet.Close
End Sub


Public Sub PrintContestTeams()
Dim PilotSet As Recordset
Dim Listset As Recordset
Dim SQLString1 As String
Dim Tab1 As Integer
Dim tabset(10) As Integer
Dim NumPilots As Integer
Dim Linecount As Integer
Dim Col As Integer
Dim P As Integer
Dim Count As Integer
Dim CurrentTeam As Integer
Dim TeamScore As Integer
Dim sco() As Single
On Error GoTo Errorhandler:
  tabset(0) = 1
  tabset(1) = 15
  tabset(2) = 35
  tabset(3) = 45
  tabset(4) = 55
  tabset(5) = 65
  tabset(6) = 75
  tabset(7) = 85
  tabset(8) = 95
  tabset(9) = 105

  TeamScore = 0
  CurrentTeam = 0
  Call ScoreAddUp(sco(), NumPilots)
  
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID,List.FreqUsed, List.Team_ID, List.InTeam, List.FinalCompScore, List.TempScore, Pilot.First_Name, Pilot.Last_Name, Pilot.Freq1, Pilot.Freq2, Pilot.Club, Pilot.Number, Pilot.Res1 FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & ")) ORDER BY List.Team_ID ASC;"
  Set PilotSet = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  PilotSet.MoveFirst
  Printer.FontSize = 16
  Printer.CurrentX = Printer.ScaleWidth / 2 - TextWidth("Leader Board")
  Printer.Print "List of Teams in Contest"
  Linecount = 4
  Printer.FontSize = 14
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.Print "Freq"
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.CurrentX = TextWidth("g") * (tabset(1) + 8)
  Printer.Print "Name"
  Printer.CurrentX = TextWidth("g") * tabset(4)
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.Print "Score"
  Printer.FontSize = 12
  Printer.Print ""
  Linecount = Linecount + 1
  Printer.CurrentY = TextHeight("sample") * Linecount
  
  'Print out the information
  Linecount = Linecount + 1
  Count = 1
  If PilotSet!InTeam = False Then
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.CurrentX = TextWidth("g") * tabset(0)
    Printer.Print "Pilots not in Teams"
    If Linecount < Pagelength Then
      Linecount = Linecount + 2
    Else
      Printer.NewPage
      Linecount = 2
    End If
    Printer.FontSize = 10
  
    Do Until PilotSet!InTeam = True
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.CurrentX = TextWidth("g") * tabset(0)
      Printer.Print PilotSet!FreqUsed
      Printer.CurrentX = TextWidth("g") * tabset(1)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print Trim(PilotSet!First_Name) & " " & Trim(PilotSet!Last_Name) & " " & Trim(PilotSet!Res1)
      Printer.CurrentX = TextWidth("g") * tabset(4)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print PilotSet!TempScore
      PilotSet.MoveNext
      If PilotSet.EOF = True Then
        Exit Do
      End If
      Count = Count + 1
      If Linecount < Pagelength Then
        Linecount = Linecount + 2
      Else
        Printer.NewPage
        Linecount = 2
      End If
    Loop
  End If
  Do Until PilotSet.EOF
    If CurrentTeam <> PilotSet!Team_ID Then
      CurrentTeam = PilotSet!Team_ID
      If CurrentTeam > 1 Then
        Printer.CurrentY = TextHeight("sample") * Linecount
        Printer.CurrentX = TextWidth("g") * tabset(1)
        Printer.Print "Total"
        Printer.CurrentY = TextHeight("sample") * Linecount
        Printer.CurrentX = TextWidth("g") * tabset(4)
        Printer.Print TeamScore
        TeamScore = 0
        Linecount = Linecount + 2
      End If
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.CurrentX = TextWidth("g") * tabset(0)
      Printer.FontSize = 12
      Printer.Print "Team "; CurrentTeam
      Printer.FontSize = 10
      If Linecount < Pagelength Then
        Linecount = Linecount + 2
      Else
        Printer.NewPage
        Linecount = 2
      End If
    End If
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.CurrentX = TextWidth("g") * tabset(0)
    Printer.Print PilotSet!FreqUsed
    Printer.CurrentX = TextWidth("g") * tabset(1)
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.Print Trim(PilotSet!First_Name) & " " & Trim(PilotSet!Last_Name) & " " & Trim(PilotSet!Res1)
    Printer.CurrentX = TextWidth("g") * tabset(4)
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.Print PilotSet!TempScore
    TeamScore = TeamScore + PilotSet!TempScore
    PilotSet.MoveNext
    Count = Count + 1
    If Linecount < Pagelength Then
      Linecount = Linecount + 2
    Else
      Printer.NewPage
      Linecount = 2
    End If
  Loop
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.CurrentX = TextWidth("g") * tabset(1)
  Printer.Print "Total"
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.CurrentX = TextWidth("g") * tabset(4)
  Printer.Print TeamScore
  Printer.EndDoc
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, List.TempScore,List.FinalCompScore, list.FlyoffScore From List WHERE (List.Comp_ID=" & Str(CurrentContest) & ") ORDER BY List.Pilot_ID;"
  Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Listset.MoveFirst
  Do Until Listset.EOF
   Listset.Edit
   Listset!TempScore = 0
   Listset.Update
   Listset.MoveNext
 Loop
PilotSet.Close
Listset.Close
Exit Sub
Errorhandler:
  MsgBox ("No Teams Entered")

End Sub

Public Sub PrintSummary()
Dim Listset As Recordset
Dim Scoreset As Recordset
Dim SQLString1 As String
Dim rounds As Integer
Dim I As Integer
Dim Place As Integer
Dim Tab1 As Integer
Dim Count As Integer
Dim sco() As Single
Dim maxnum As Integer
Dim tabset(100) As Integer
Dim NumPilots As Integer
Dim MinScore As Integer
Dim Linecount As Integer
Dim Col As Integer
Dim P As Integer
Dim Mins As Integer
Dim Secs As Single
Dim J As Integer
Dim Pagewidth As Integer
Dim LanPageLength As Integer
Dim Startround As Integer
Dim Endround As Integer
LanPageLength = 50

  tabset(0) = 1
  tabset(1) = 8
  tabset(2) = 15
  tabset(3) = 20
  tabset(4) = 25
  tabset(5) = 30
  tabset(6) = 35
  tabset(7) = 40
  tabset(8) = 47
  tabset(9) = 50
  tabset(10) = 55
  'Landscape Pagewidth = 170
  If NumRounds < 6 Then
    J = 60
    For I = 11 To 45
      tabset(I) = J
      J = J + 4
    Next I
    Pagewidth = 45
    Startround = 1
    Endround = NumRounds
  Else
    Startround = InputBox("Which round to start Printing from ")
    Endround = InputBox("Which round to finish Printing < 6 more than start")
    If Endround > NumRounds Then
      MsgBox ("End round is greater than the number of rounds")
      Do While Endround > NumRounds
        Endround = InputBox("Which round to finish Printing < 6 more than start")
      Loop
    End If
    If Endround - Startround > 6 Then
      MsgBox ("A Maximum or 6 rounds can be printed at a time")
      Do While Endround - Startround > 6
        Startround = InputBox("Which round to start Printing from ")
        Endround = InputBox("Which round to finish Printing < 6 more than start")
      Loop
    End If
    J = 60
    For I = 11 To 56
      tabset(I) = J
      J = J + 3
    Next I
    Pagewidth = 55
  End If
  rounds = NumRounds
  Printer.Orientation = 2
  Call ScoreAddUp(sco(), NumPilots)
  Call SummHeader(Linecount, tabset(), Pagewidth, Startround, Endround)
  
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Pilot.Res1 FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & ")) ORDER BY List.TempScore DESC;"
  Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Listset.MoveFirst
  
  Linecount = Linecount + 1
  Count = Startround
  Do Until Listset.EOF
    SQLString1 = "SELECT * From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Pilot_ID=" & Str(Listset!Pilot_ID) & ")) ORDER BY Scores.Round,Scores.Task ASC;"
    Set Scoreset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
    Linecount = Linecount + 1
    Printer.FontSize = 10
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.CurrentX = TextWidth("g") * tabset(0)
    Printer.Print Count
    Printer.CurrentX = TextWidth("g") * tabset(1)
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.Print Trim(Listset!First_Name) & " " & Trim(Listset!Last_Name) & " " & Trim(Listset!Res1)
    Printer.CurrentX = TextWidth("g") * tabset(6)
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.Print sco(Listset!Pilot_ID, 0)
    Col = 10
    If CurrentContestType = "F3B" Then
      Printer.FontSize = 10
      
      Printer.CurrentX = TextWidth("g") * tabset(8)
      Printer.CurrentY = TextHeight("sample") * (Linecount)
      Printer.Print "Score"
      Printer.CurrentX = TextWidth("g") * tabset(8)
      Printer.CurrentY = TextHeight("sample") * (Linecount + 1)
      Printer.Print "Time"
      Printer.CurrentX = TextWidth("g") * tabset(8)
      Printer.CurrentY = TextHeight("sample") * (Linecount + 2)
      Printer.Print "Total"
      I = 1
      Scoreset.MoveFirst
      Do While I <> Startround
       Scoreset.MoveNext
       I = I + 1
      Loop
      Do Until I > Endround
        Printer.FontSize = 8
        Printer.CurrentX = TextWidth("g") * tabset(Col)
        Printer.CurrentY = TextHeight("sample") * Linecount
        Printer.Print Scoreset!Score
        
        If Scoreset!task = "A" Then
          Mins = Scoreset!Flight_Time \ 60
          Secs = Scoreset!Flight_Time Mod 60
          Printer.CurrentX = TextWidth("g") * tabset(Col)
          Printer.CurrentY = TextHeight("sample") * (Linecount + 1)
          Printer.Print Trim(Str(Mins)); ":"; Trim(Str(Secs))
        ElseIf Scoreset!task = "B" Then
          Printer.CurrentX = TextWidth("g") * tabset(Col)
          Printer.CurrentY = TextHeight("sample") * (Linecount + 1)
          Printer.Print Trim(Scoreset!Flight_Time)
          Printer.CurrentX = TextWidth("g") * tabset(Col)
          Printer.CurrentY = TextHeight("sample") * (Linecount + 2)
          Printer.FontSize = 10
          Printer.Print Trim(Str(sco(Listset!Pilot_ID, I)))
          Printer.FontSize = 8
          'I = I + 1
        ElseIf Scoreset!task = "C" Then
          Mins = Scoreset!Flight_Time \ 100
          Secs = Scoreset!Flight_Time Mod 100
          Printer.CurrentX = TextWidth("g") * tabset(Col)
          Printer.CurrentY = TextHeight("sample") * (Linecount + 1)
          Printer.Print Trim(Str(Mins)); "."; Trim(Str(Secs))
          I = I + 1
        End If
        Col = Col + 2
        'If Col > Pagewidth Then
        '  Linecount = Linecount + 4
        '  Col = 10
        '  Printer.FontSize = 9
        '  Printer.CurrentX = TextWidth("g") * tabset(8)
        '  Printer.CurrentY = TextHeight("sample") * (Linecount)
        '  Printer.Print "Score"
        '  Printer.CurrentX = TextWidth("g") * tabset(8)
        '  Printer.CurrentY = TextHeight("sample") * (Linecount + 1)
        '  Printer.Print "Time"
        '  Printer.CurrentX = TextWidth("g") * tabset(8)
        '  Printer.CurrentY = TextHeight("sample") * (Linecount + 2)
        '  Printer.Print "Total"
        '  Printer.FontSize = 8
        'End If
        
        'If Scoreset.EOF <> True Then
          Scoreset.MoveNext
        'End If
      Loop
    
    
    Else
     'F3j printing part
      Printer.FontSize = 10
      
      Printer.CurrentX = TextWidth("g") * tabset(8)
      Printer.CurrentY = TextHeight("sample") * (Linecount)
      Printer.Print "Flight"
      Printer.CurrentX = TextWidth("g") * tabset(8)
      Printer.CurrentY = TextHeight("sample") * (Linecount + 1)
      Printer.Print "Score"
      
      I = 1
      Scoreset.MoveFirst
      Do While I <> Startround
       Scoreset.MoveNext
       I = I + 1
      Loop
      Do Until I = Endround + 1
        Printer.FontSize = 8
        Mins = Scoreset!Flight_Time \ 60
        'Secs = Scoreset!Flight_Time Mod 60
        Secs = Scoreset!Flight_Time - ((Scoreset!Flight_Time \ 60) * 60)
        Printer.CurrentX = TextWidth("g") * tabset(Col)
        Printer.CurrentY = TextHeight("sample") * Linecount
        Printer.Print Mins; ":"; Format(Secs, "0#.0")
        Printer.CurrentX = TextWidth("g") * tabset(Col + 2)
        Printer.CurrentY = TextHeight("sample") * Linecount
        Printer.Print " "; Scoreset!LandBonus
        Printer.CurrentX = TextWidth("g") * tabset(Col + 4)
        Printer.CurrentY = TextHeight("sample") * Linecount
        Printer.Print Scoreset!Penalty + Scoreset!Res2
        
        Printer.CurrentX = TextWidth("g") * tabset(Col + 2)
        Printer.CurrentY = TextHeight("sample") * (Linecount + 1)
        Printer.Print Scoreset!Score
        Col = Col + 6
        If Col > Pagewidth Then
          Linecount = Linecount + 4
          Col = 10
          Printer.FontSize = 10
          Printer.CurrentX = TextWidth("g") * tabset(8)
          Printer.CurrentY = TextHeight("sample") * (Linecount)
          Printer.Print "Flight"
          Printer.CurrentX = TextWidth("g") * tabset(8)
          Printer.CurrentY = TextHeight("sample") * (Linecount + 1)
          Printer.Print "Score"
        End If
        I = I + 1
        If Scoreset.EOF <> True Then
          Scoreset.MoveNext
        End If
      Loop
    
    End If
    Listset.MoveNext
    Count = Count + 1
    If Linecount < LanPageLength Then
      Linecount = Linecount + 3
    Else
      Call SummFooter(Linecount, tabset(), Startround, Endround)
      Printer.NewPage
      Linecount = 1
      Call SummHeader(Linecount, tabset(), Pagewidth, Startround, Endround)
      Linecount = Linecount + 1
    End If
  Loop
  
  
  
 SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, List.TempScore From List WHERE (List.Comp_ID=" & Str(CurrentContest) & ") ORDER BY List.Pilot_ID;"
 Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
 Listset.MoveFirst
 Do Until Listset.EOF
   Listset.Edit
   Listset!TempScore = 0
   Listset.Update
   Listset.MoveNext
 Loop
 Call SummFooter(Linecount, tabset(), Startround, Endround)
 Printer.EndDoc
End Sub

Public Sub PrintRound()
Dim PilotSet As Recordset
Dim SQLString1 As String
Dim Tab1 As Integer
Dim tabset(100) As Integer
Dim NumPilots As Integer
Dim Linecount As Integer
Dim Col As Integer
Dim P As Integer
Dim Count As Integer
Dim RoundNum
Dim Mins As Integer
Dim Secs As Single
Dim Currentslot As Integer
Dim Colwidth As Integer
Dim Pagewidth As Integer
Dim J As Integer
Dim I As Integer
Dim RoundScore As Single
Dim ActualPenalty As Single

  Pagelength = 68
  tabset(0) = 1
  tabset(1) = 8
  tabset(2) = 15
  tabset(3) = 20
  tabset(4) = 25
  tabset(5) = 30
  tabset(6) = 35
  tabset(7) = 40
  tabset(8) = 45
  tabset(9) = 50
  tabset(10) = 55
  
  
  'Landscape Pagewidth = 170
  J = 60
  For I = 11 To 45
    tabset(I) = J
    J = J + 5
  Next I
  Pagewidth = 45
  'Else
  '  J = 60
  '  For I = 11 To 56
  '    tabset(I) = J
  '    J = J + 3
  '  Next I
  '  Pagewidth = 55
  'End If
  RoundNum = InputBox("Enter the Round Number that you want printed", 0)
  If RoundNum > NumRounds Then
    MsgBox ("Round Number is greater than the number of Rounds")
    Do While RoundNum > NumRounds Or RoundNum <> 0
      RoundNum = InputBox("Enter the Round Number that you want printed", 0)
    Loop
  ElseIf RoundNum = 0 Then
    Exit Sub
  End If
  If CurrentContestType = "F3B" Then
    SQLString1 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Pilot.Res1 ,Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.RawScore, Scores.Task, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & "))ORDER BY Scores.Slot ASC,Pilot.Pilot_ID,Scores.Task ASC;"
  Else
    SQLString1 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.First_Name, Pilot.Pilot_ID, Pilot.Last_Name, Pilot.Res1 ,Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.RawScore, Scores.Task, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & "))ORDER BY Scores.Slot ASC,Pilot.Pilot_ID,Scores.Score DESC;"
  End If
  Set PilotSet = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Linecount = Linecount + 1
  PilotSet.MoveFirst
  
  Printer.FontSize = 16
  Printer.CurrentX = Printer.ScaleWidth / 2 - TextWidth("Summary of Scores for Round  ")
  Printer.Print "Summary of Scores for Round " & Str(RoundNum)
  Linecount = 4
  Printer.FontSize = 14
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.Print "Slot"
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.CurrentX = TextWidth("g") * (tabset(1) + 7)
  Printer.Print "Name"
  Printer.CurrentX = TextWidth("g") * tabset(8)
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.Print "Score"
  'Skipping a column here
  Colwidth = 10
  Linecount = Linecount + 2
    
  If CurrentContestType = "F3B" Then
    Printer.FontSize = 10
    Printer.CurrentY = TextHeight("sample") * (Linecount)
    Printer.CurrentX = TextWidth("g") * tabset(Colwidth)
    Printer.Print " "; "Time"
    Printer.CurrentY = TextHeight("sample") * (Linecount)
    Printer.CurrentX = TextWidth("g") * tabset(Colwidth + 2)
    Printer.Print " "; "Land"
    Printer.CurrentY = TextHeight("sample") * (Linecount)
    Printer.CurrentX = TextWidth("g") * tabset(Colwidth + 4)
    Printer.Print " "; "Pen"
    Printer.CurrentY = TextHeight("sample") * (Linecount)
    Printer.CurrentX = TextWidth("g") * tabset(Colwidth + 6)
    Printer.Print " "; "Score"
    Linecount = Linecount + 2
  Else
    Printer.FontSize = 10
    Printer.CurrentY = TextHeight("sample") * (Linecount)
    Printer.CurrentX = TextWidth("g") * tabset(Colwidth)
    Printer.Print " "; "Time"
    Printer.CurrentY = TextHeight("sample") * (Linecount)
    Printer.CurrentX = TextWidth("g") * tabset(Colwidth + 2)
    Printer.Print " "; "Land"
    Printer.CurrentY = TextHeight("sample") * (Linecount)
    Printer.CurrentX = TextWidth("g") * tabset(Colwidth + 4)
    Printer.Print " "; "F Pen"
    Printer.CurrentY = TextHeight("sample") * (Linecount)
    Printer.CurrentX = TextWidth("g") * tabset(Colwidth + 6)
    Printer.Print " "; "S Pen"
    Linecount = Linecount + 1
  End If
  If CurrentContestType = "F3B" Then
    PilotSet.MoveFirst
    Do Until PilotSet.EOF
      RoundScore = 0
      Printer.FontSize = 10
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.CurrentX = TextWidth("g") * tabset(0)
      Printer.Print PilotSet!Slot
      Printer.CurrentX = TextWidth("g") * tabset(1)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print Trim(PilotSet!First_Name) & " " & Trim(PilotSet!Last_Name) & " " & Trim(PilotSet!Res1)
      Col = 10
      Printer.CurrentX = TextWidth("g") * tabset(9)
      Printer.CurrentY = TextHeight("sample") * (Linecount)
      Printer.Print "A"
      Printer.CurrentX = TextWidth("g") * tabset(Col + 6)
      Printer.CurrentY = TextHeight("sample") * (Linecount)
      Printer.Print PilotSet!Score
      Mins = PilotSet!Flight_Time \ 60
      Secs = PilotSet!Flight_Time Mod 60
      Printer.CurrentX = TextWidth("g") * tabset(Col)
      Printer.CurrentY = TextHeight("sample") * (Linecount)
      Printer.Print Trim(Str(Mins)); ":"; Trim(Str(Secs))
      Printer.CurrentX = TextWidth("g") * tabset(Col + 2)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print PilotSet!LandBonus
      Printer.CurrentX = TextWidth("g") * tabset(Col + 4)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print PilotSet!Res2
      RoundScore = RoundScore + PilotSet!Score
      Linecount = Linecount + 1
      PilotSet.MoveNext
      Printer.CurrentX = TextWidth("g") * tabset(9)
      Printer.CurrentY = TextHeight("sample") * (Linecount)
      Printer.Print "B"
      Printer.CurrentX = TextWidth("g") * tabset(Col + 6)
      Printer.CurrentY = TextHeight("sample") * (Linecount)
      Printer.Print PilotSet!Score
      Printer.CurrentX = TextWidth("g") * tabset(Col)
      Printer.CurrentY = TextHeight("sample") * (Linecount)
      Printer.Print PilotSet!Flight_Time
      Printer.CurrentX = TextWidth("g") * tabset(Col + 4)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print PilotSet!Res2
      RoundScore = RoundScore + PilotSet!Score
      Linecount = Linecount + 1
      PilotSet.MoveNext
      Printer.CurrentX = TextWidth("g") * tabset(9)
      Printer.CurrentY = TextHeight("sample") * (Linecount)
      Printer.Print "C"
      Printer.CurrentX = TextWidth("g") * tabset(Col + 6)
      Printer.CurrentY = TextHeight("sample") * (Linecount)
      Printer.Print PilotSet!Score
      Mins = PilotSet!Flight_Time \ 100
      Secs = PilotSet!Flight_Time Mod 100
      Printer.CurrentX = TextWidth("g") * tabset(Col)
      Printer.CurrentY = TextHeight("sample") * (Linecount)
      Printer.Print Trim(Str(Mins)); ":"; Trim(Str(Secs))
      Printer.CurrentX = TextWidth("g") * tabset(Col + 4)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print PilotSet!Penalty
      RoundScore = RoundScore + PilotSet!Score
      Printer.CurrentX = TextWidth("g") * tabset(8)
      Printer.CurrentY = TextHeight("sample") * (Linecount - 2)
      Printer.Print RoundScore
      Linecount = Linecount + 2
      If Linecount > Pagelength Then
        Printer.NewPage
        Linecount = 2
        Printer.FontSize = 14
        Printer.CurrentY = TextHeight("sample") * Linecount
        Printer.Print "Slot"
        Printer.CurrentY = TextHeight("sample") * Linecount
        Printer.CurrentX = TextWidth("g") * (tabset(1) + 7)
        Printer.Print "Name"
        Printer.CurrentX = TextWidth("g") * tabset(6)
        Printer.CurrentY = TextHeight("sample") * Linecount
        Printer.Print "Score"
        'Skipping a column here
        Colwidth = 10
        Linecount = Linecount + 2
        Printer.FontSize = 10
        Printer.CurrentY = TextHeight("sample") * (Linecount)
        Printer.CurrentX = TextWidth("g") * tabset(Colwidth)
        Printer.Print " "; "Time"
        Printer.CurrentY = TextHeight("sample") * (Linecount)
        Printer.CurrentX = TextWidth("g") * tabset(Colwidth + 2)
        Printer.Print " "; "Land"
        Printer.CurrentY = TextHeight("sample") * (Linecount)
        Printer.CurrentX = TextWidth("g") * tabset(Colwidth + 4)
        Printer.Print " "; "Pen"
        Printer.CurrentY = TextHeight("sample") * (Linecount)
        Printer.CurrentX = TextWidth("g") * tabset(Colwidth + 6)
        Printer.Print " "; "Score"
        Linecount = Linecount + 2
      
      End If
      If PilotSet.EOF <> True Then
        PilotSet.MoveNext
      End If
    Loop
  ElseIf CurrentContestType = "AustOpen" Then
    'AustOpen printing part
    PilotSet.MoveFirst
    Do Until PilotSet.EOF
      Printer.FontSize = 10
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.CurrentX = TextWidth("g") * tabset(0)
      Printer.Print PilotSet!Slot
      Printer.CurrentX = TextWidth("g") * tabset(1)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print Trim(PilotSet!First_Name) & " " & Trim(PilotSet!Last_Name) & " " & Trim(PilotSet!Res1)
      Printer.CurrentX = TextWidth("g") * tabset(8)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print PilotSet!Score
      Col = 10
      Mins = PilotSet!Flight_Time \ 60
      Secs = PilotSet!Flight_Time Mod 60
      Printer.CurrentX = TextWidth("g") * tabset(Col)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print Trim(Str(Mins)); ":"; Trim(Str(Secs))
      Printer.CurrentX = TextWidth("g") * tabset(Col + 2)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print PilotSet!LandBonus
      Printer.CurrentX = TextWidth("g") * tabset(Col + 5)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print PilotSet!Res2
      Linecount = Linecount + 2
      If PilotSet.EOF <> True Then
        PilotSet.MoveNext
      End If
    Loop
  Else
    'F3J printing part
    PilotSet.MoveFirst
    Do Until PilotSet.EOF
      Printer.FontSize = 10
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.CurrentX = TextWidth("g") * tabset(0)
      Printer.Print PilotSet!Slot
      Printer.CurrentX = TextWidth("g") * tabset(1)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print Trim(PilotSet!First_Name) & " " & Trim(PilotSet!Last_Name) & " " & Trim(PilotSet!Res1)
      Printer.CurrentX = TextWidth("g") * tabset(8)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print PilotSet!Score
      Col = 10
      Mins = PilotSet!Flight_Time \ 60
      Secs = PilotSet!Flight_Time - ((PilotSet!Flight_Time \ 60) * 60)
      Printer.CurrentX = TextWidth("g") * tabset(Col)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print Trim(Str(Mins)); ":"; Format(Trim(Str(Secs)), "0#.0")
      Printer.CurrentX = TextWidth("g") * tabset(Col + 2)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print PilotSet!LandBonus
      Printer.CurrentX = TextWidth("g") * tabset(Col + 4)
      Printer.CurrentY = TextHeight("sample") * Linecount
      If PilotSet!Penalty > 0 Then
        ActualPenalty = PilotSet!Penalty - (PilotSet!Flight_Time - (max * 60))
        Printer.Print ActualPenalty
      Else
        Printer.Print PilotSet!Penalty
      End If
      Printer.CurrentX = TextWidth("g") * tabset(Col + 6)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print PilotSet!Res2
      Linecount = Linecount + 2
      If PilotSet.EOF <> True Then
        PilotSet.MoveNext
      End If
    Loop
  End If
  Printer.EndDoc
End Sub

Public Sub SummHeaderHTML(Colwidth)
Dim I As Integer
Dim OutString As String
  
  Print #1, "<TR>"
  OutString = "<TH Width =" & """2%""" & ">Place</TH>"
  Print #1, OutString
  OutString = "<TH Width =" & """18%""" & ">Name</TH>"
  Print #1, OutString
  OutString = "<TH Width =" & """5%""" & ">Total</TH>"
  Print #1, OutString
  OutString = "<TH Width =" & """5%""" & "></TH>"
  Print #1, OutString
  
    
  
  
  For I = 1 To NumRounds
    OutString = "<TH Width =""" & Trim(Str(Colwidth)) & "%"">" & "<BIG>" & "R" & Str(I) & "</BIG>" & "</TH>"
    Print #1, OutString
    OutString = "<TH Width =""" & Trim(Str(Colwidth)) & "%"">" & "</TH>"
    Print #1, OutString
    OutString = "<TH Width =""" & Trim(Str(Colwidth)) & "%"">" & "</TH>"
    Print #1, OutString
    
  Next I
  Print #1, "</TR>"
  
  
  If CurrentContestType = "F3B" Then
      'skip over the first 4 columns
      Skip4cols "Task"
      For I = 1 To NumRounds
        OutString = "<TH Width =""" & Trim(Str(Colwidth)) & "%"">A" & "</TH>"
        Print #1, OutString
        OutString = "<TH Width =""" & Trim(Str(Colwidth)) & "%"">B" & "</TH>"
        Print #1, OutString
        OutString = "<TH Width =""" & Trim(Str(Colwidth)) & "%"">C" & "</TH>"
        Print #1, OutString
      Next I
      Print #1, "</TR>"
      
    Else
     'skip over the first 4 columns
      Skip4cols " "
      For I = 1 To NumRounds
        OutString = "<TH Width =""" & Trim(Str(Colwidth)) & "%"">Time" & "</TH>"
        Print #1, OutString
        OutString = "<TH Width =""" & Trim(Str(Colwidth)) & "%"">Land" & "</TH>"
        Print #1, OutString
        OutString = "<TH Width =""" & Trim(Str(Colwidth)) & "%"">Pen" & "</TH>"
        Print #1, OutString
      Next I
      Print #1, "</TR>"
    
    End If
  
 
  
End Sub

Public Sub SummFooter(Linecount As Integer, tabset() As Integer, Startround As Integer, Endround As Integer)

Dim AX, AY, BX, BY
Dim Colwidth
Dim I As Integer

Colwidth = 10

  For I = Startround To Endround
    DrawWidth = 2
    Printer.CurrentY = TextHeight("sample") * 4
    Printer.CurrentX = TextWidth("g") * tabset(Colwidth)
    BX = TextWidth("g") * tabset(Colwidth)
    BY = TextHeight("sample") * Linecount
    Printer.Line -(BX, BY)
    Colwidth = Colwidth + 6
  Next I
  
  Printer.CurrentY = TextHeight("sample") * 4
  Printer.CurrentX = TextWidth("g") * tabset(Colwidth)
  BX = TextWidth("g") * tabset(Colwidth)
  BY = TextHeight("sample") * Linecount
  Printer.Line -(BX, BY)

End Sub

Public Sub PrintFlightGroups()
Dim PilotSet As Recordset
Dim Selected As Recordset
Dim SQLString1 As String
Dim NumPilots As Integer
Dim CurrentGroup As Integer
Dim Line As String
Dim RoundNum
Dim Slot
Dim Search As String
Dim tabset(13) As Integer
Dim Linecount As Integer
Dim Col As Integer
Dim Pos As Integer

  tabset(0) = 1
  tabset(1) = 25
  tabset(2) = 35
  tabset(3) = 45
  tabset(4) = 55
  tabset(5) = 65
  tabset(6) = 75
  tabset(7) = 85
  tabset(8) = 95
  tabset(9) = 105


 Col = 0
 ReDim Spots(15)
 If FlightGroups = False Then
   MsgBox ("Task B Flight groups not set up")
   Exit Sub
 Else
   CurrentGroup = 0
    'SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & "));"
    SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID,List.FreqUsed, List.Team_ID, List.InTeam, Pilot.First_Name, Pilot.Last_Name, Pilot.Freq1, Pilot.Freq2, Pilot.Club, Pilot.Number FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & "));"
    Set PilotSet = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
   PilotSet.MoveFirst
   
   Printer.FontSize = 16
   Printer.Orientation = 1
   Pagelength = 75
   Printer.CurrentX = Printer.ScaleWidth / 2 - TextWidth("Round/Slot Allocation")
   Printer.Print "Summary of Flight Groups"
   Linecount = 4
   For RoundNum = 1 To NumRounds
   
     Printer.FontSize = 14
     Printer.CurrentY = TextHeight("sample") * Linecount
     Printer.Print "Round "; RoundNum
     If Linecount < Pagelength Then
        Linecount = Linecount + 2
     Else
        Printer.NewPage
        Linecount = 2
     End If
     For Slot = 1 To NumSlots
       Printer.FontSize = 12
       Printer.CurrentY = TextHeight("sample") * Linecount
       Printer.Print "Slot "; Slot
       Linecount = Linecount + 1
       Printer.FontSize = 10
       CurrentGroup = 0
       SQLString1 = "SELECT * From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Round= " & Trim(Str(RoundNum)) & ") AND (Scores.Slot=" & Trim(Str(Slot)) & ") AND (Scores.Task = 'B'))ORDER By Scores.FlightGroup;"
       Set Selected = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
       Selected.MoveFirst
       Do Until Selected.EOF
         If CurrentGroup <> Selected!FlightGroup Then
           Pos = 1
           If Linecount < Pagelength Then
             Linecount = Linecount + 2
           Else
             Printer.NewPage
             Linecount = 1
           End If
           CurrentGroup = Selected!FlightGroup
           Printer.CurrentY = TextHeight("sample") * Linecount
           Printer.CurrentX = TextWidth("g") * tabset(0)
           Printer.Print "Flight Group" & CurrentGroup
          End If
          Search = "Pilot_ID = " & Str(Selected!Pilot_ID)
          PilotSet.FindFirst Search
          Printer.CurrentX = TextWidth("g") * tabset(Pos)
          Printer.CurrentY = TextHeight("sample") * Linecount
          Printer.Print Trim(Left(PilotSet!First_Name, 1)) & ". " & Trim(PilotSet!Last_Name)
          PilotSet.MoveFirst
          Selected.MoveNext
          Pos = Pos + 2
        Loop
      If Linecount < Pagelength Then
        Linecount = Linecount + 2
      Else
        Printer.NewPage
        Linecount = 2
      End If
      Next Slot
    Next RoundNum
  
   Printer.EndDoc
  
   PilotSet.Close
   Selected.Close
 End If


End Sub


Public Sub PrintContestTeamsHTML()
Dim PilotSet As Recordset
Dim Listset As Recordset
Dim SQLString1 As String
Dim Tab1 As Integer
Dim tabset(10) As Integer
Dim NumPilots As Integer
Dim Linecount As Integer
Dim Col As Integer
Dim P As Integer
Dim Count As Integer
Dim CurrentTeam As Integer
Dim TeamScore As Integer
Dim sco() As Single
Dim OutString As String

  
  tabset(0) = 1
  tabset(1) = 15
  tabset(2) = 35
  tabset(3) = 45
  tabset(4) = 55
  tabset(5) = 65
  tabset(6) = 75
  tabset(7) = 85
  tabset(8) = 95
  tabset(9) = 105
   
  Open OutputFile For Output As #1
  
  TeamScore = 0
  CurrentTeam = 0
  Call ScoreAddUp(sco(), NumPilots)
  
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID,List.FreqUsed, List.Team_ID, List.InTeam, List.FinalCompScore, Pilot.First_Name, Pilot.Last_Name, Pilot.Freq1, Pilot.Freq2, Pilot.Club, Pilot.Number, List.TempScore FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & ")) ORDER BY List.Team_ID ASC;"
  Set PilotSet = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  
  Print #1, "<HTML>"
  Print #1, "<HEAD>"
  Print #1, "<TITLE>"
  Print #1, "List of Teams in Contest"
  Print #1, "</TITLE>"
  Print #1, "</HEAD>"
  Print #1, "<BODY>"
  Print #1, "<P ALIGN=Center><BIG><BIG>List of Teams in Contest<BIG></BIG>"
  Print #1, "<p>"
  Print #1, "<TABLE>"
  OutString = "<TABLE BORDER WIDTH=" & """100%""" & ">"
  Print #1, OutString
  Print #1, "<TR>"
  Print #1, "<TH>Freq</TH>"
  Print #1, "<TH>Name</TH>"
  Print #1, "<TH>Score</TH>"
  Print #1, "</TR>"
  
  
  'Print out the information
  
  Count = 1
  
  If PilotSet!InTeam = False Then
    Print #1, "<TR>"
    Print #1, "<TD COLSPAN = 3>Pilots not in Teams</TD>"
    Print #1, "</TR>"
    Do Until PilotSet!InTeam = True
      Print #1, "<TR>"
      OutString = "<TD>" & Str(PilotSet!FreqUsed) & "</TD>"
      Print #1, OutString
      OutString = "<TD>" & Trim(PilotSet!First_Name) & " " & Trim(PilotSet!Last_Name) & " " & Trim(PilotSet!Res1) & "</TD>"
      Print #1, OutString
      OutString = "<TD>" & Str(PilotSet!TempScore) & "</TD>"
      Print #1, OutString
      Print #1, "</TR>"
      PilotSet.MoveNext
      If PilotSet.EOF = True Then
        Exit Do
      End If
    Loop
  End If
  Do Until PilotSet.EOF
    If CurrentTeam <> PilotSet!Team_ID Then
      CurrentTeam = PilotSet!Team_ID
      If CurrentTeam > 1 Then
        Print #1, "<TR>"
        Print #1, "<TD COLSPAN = 2>Total</TD>"
        OutString = "<TD>" & Str(TeamScore) & "</TD>"
        Print #1, OutString
        TeamScore = 0
        Print #1, "</TR>"
      End If
      Print #1, "<TR>"
      OutString = "<TD COLSPAN = 3>Team  " & Str(CurrentTeam) & "</TD>"
      Print #1, OutString
      Print #1, "</TR>"
    End If
    Print #1, "<TR>"
      OutString = "<TD>" & Str(PilotSet!FreqUsed) & "</TD>"
      Print #1, OutString
      OutString = "<TD>" & Trim(PilotSet!First_Name) & " " & Trim(PilotSet!Last_Name) & " " & Trim(PilotSet!Res1) & "</TD>"
      Print #1, OutString
      OutString = "<TD>" & Str(PilotSet!TempScore) & "</TD>"
      Print #1, OutString
      Print #1, "</TR>"
    TeamScore = TeamScore + PilotSet!TempScore
    PilotSet.MoveNext
    
  Loop
  Print #1, "<TR>"
  Print #1, "<TD COLSPAN = 2>Total</TD>"
  OutString = "<TD>" & Str(TeamScore) & "</TD>"
  Print #1, OutString
  TeamScore = 0
  Print #1, "</TR>"
  
  'end the HTML document
  Print #1, "</TABLE>"
  Print #1, "</BODY>"
  Print #1, "</HTML>"
  Close #1
  MsgBox ("Report printed")
  'do house keeping
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, List.TempScore,List.FinalCompScore, list.FlyoffScore From List WHERE (List.Comp_ID=" & Str(CurrentContest) & ") ORDER BY List.Pilot_ID;"
  Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Listset.MoveFirst
  Do Until Listset.EOF
   Listset.Edit
   Listset!TempScore = 0
   Listset.Update
   Listset.MoveNext
  Loop
PilotSet.Close
Listset.Close
End Sub

Public Sub PrintLeaderBoardHTML()
Dim Listset As Recordset
Dim Scoreset As Recordset
Dim SQLString1 As String
Dim rounds As Integer
Dim sco() As Single
Dim Count As Integer
Dim NumPilots As Integer
Dim tabset(10) As Integer
Dim Linecount As Integer
Dim Col As Integer
Dim P As Integer
Dim I As Integer
Dim Colwidth As Integer
Dim OutString As String

  
  

  rounds = NumRounds
  Call ScoreAddUp(sco(), NumPilots)
  Count = 1
  Open OutputFile For Output As #1
  Print #1, "<HTML>"
  Print #1, "<HEAD>"
  Print #1, "<TITLE>"
  Print #1, "Full Summary of Contest"
  Print #1, "</TITLE>"
  Print #1, "</HEAD>"
  Print #1, "<BODY>"
  Print #1, "<P ALIGN=Center><BIG><BIG>Leader Board<BIG></BIG>"
  Print #1, "<p>"
  Print #1, "<TABLE>"
  If NumRounds < 5 Then
    Colwidth = 50 / NumRounds
    OutString = "<TABLE BORDER WIDTH=" & """80%""" & ">"
    Print #1, OutString
  Else
    Colwidth = 40 / NumRounds
    OutString = "<TABLE BORDER WIDTH=" & """100%""" & ">"
    Print #1, OutString
  End If
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Pilot.Res1 FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & ")) ORDER BY List.TempScore DESC;"
  Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Listset.MoveFirst
  
  
  Print #1, "<TR>"
  OutString = "<TH Width =" & """2%""" & ">Place</TH>"
  Print #1, OutString
  OutString = "<TH Width =" & """18%""" & ">Name</TH>"
  Print #1, OutString
  OutString = "<TH Width =" & """5%""" & ">Total</TH>"
  Print #1, OutString
  
  For I = 1 To NumRounds
    OutString = "<TH Width =""" & Trim(Str(Colwidth)) & "%"">" & "Round" & Str(I) & "</TH>"
    Print #1, OutString
  Next I
  Print #1, "</TR>"
  
  
  Count = 1
  Do Until Listset.EOF
    SQLString1 = "SELECT * From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Pilot_ID=" & Str(Listset!Pilot_ID) & ")) ORDER BY Scores.Round,Scores.Task ASC;"
    Set Scoreset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
    Print #1, "<TR>"
    OutString = "<TD Width =" & """2%""" & ">" & Str(Count) & "</TD>"
    Print #1, OutString
    OutString = "<TD Width =" & """18%""" & ">" & Trim(Listset!First_Name) & " " & Trim(Listset!Last_Name) & " " & Trim(Listset!Res1) & "</TD>"
    Print #1, OutString
    OutString = "<TD Width =" & """5%""" & ">" & Str(sco(Listset!Pilot_ID, 0)) & "</TD>"
    Print #1, OutString
    
    For I = 1 To rounds
      OutString = "<TD Width =""" & Trim(Str(Colwidth)) & "%"">" & Str(sco(Listset!Pilot_ID, I)) & "</TD>"
      Print #1, OutString
    Next I
    Listset.MoveNext
    Count = Count + 1
  Loop
  Print #1, "</TABLE>"
  Print #1, "</BODY>"
  Print #1, "</HTML>"
  Close #1
  MsgBox ("Report printed")
  
  
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, List.TempScore From List WHERE (List.Comp_ID=" & Str(CurrentContest) & ") ORDER BY List.Pilot_ID;"
  Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Listset.MoveFirst
  Do Until Listset.EOF
    Listset.Edit
    Listset!TempScore = 0
    Listset.Update
    Listset.MoveNext
  Loop
Listset.Close
End Sub

Public Sub PrintPilotsContestHTML()
Dim PilotSet As Recordset
Dim SQLString1 As String
Dim Tab1 As Integer
Dim NumPilots As Integer
Dim Linecount As Integer
Dim Col As Integer
Dim P As Integer
Dim Count As Integer
Dim OutString As String
  
  Open OutputFile For Output As #1
  
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID,List.FreqUsed, List.Team_ID, Pilot.First_Name, Pilot.Last_Name, Pilot.Freq1, Pilot.Freq2, Pilot.Club, Pilot.Number, Pilot.Res1 FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & ")) ORDER BY Pilot.Last_Name ASC;"
  Set PilotSet = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Print #1, "<HTML>"
  Print #1, "<HEAD>"
  Print #1, "<TITLE>"
  Print #1, "List of Pilots in Contest"
  Print #1, "</TITLE>"
  Print #1, "</HEAD>"
  Print #1, "<BODY>"
  Print #1, "<P ALIGN=Center><BIG><BIG>List of Pilots in Contest<BIG></BIG>"
  Print #1, "<p>"
  Print #1, "<TABLE>"
  OutString = "<TABLE BORDER WIDTH=" & """100%""" & ">"
  Print #1, OutString
  Print #1, "<TR>"
  
  Print #1, "<TH>Name</TH>"
  Print #1, "<TH>Club</TH>"
  Print #1, "<TH>Freq Used</TH>"
  Print #1, "</TR>"
  
  Do Until PilotSet.EOF
    Print #1, "<TR>"
    
    OutString = "<TD>" & Trim(PilotSet!First_Name) & " " & Trim(PilotSet!Last_Name) & " " & Trim(PilotSet!Res1) & "</TD>"
    Print #1, OutString
    OutString = "<TD>" & PilotSet!Club & "</TD>"
    Print #1, OutString
    OutString = "<TD>" & Str(Format(PilotSet!FreqUsed, "##.000")) & "</TD>"
    Print #1, OutString
    Print #1, "</TR>"
    PilotSet.MoveNext
  Loop
  Print #1, "</TABLE>"
  Print #1, "</BODY>"
  Print #1, "</HTML>"
  Close #1
  MsgBox ("Report printed")
  PilotSet.Close

End Sub

Public Sub PrintRoundHTML()
Dim PilotSet As Recordset
Dim Listset As Recordset
Dim SQLString1 As String
Dim Tab1 As Integer
Dim OutString As String
Dim NumPilots As Integer
Dim Linecount As Integer
Dim Col As Integer
Dim P As Integer
Dim Count As Integer
Dim RoundNum As Integer
Dim Mins As Integer
Dim Secs As Single
Dim Currentslot As Integer
Dim Colwidth As Integer
Dim Pagewidth As Integer
Dim J As Integer
Dim I As Integer
Dim RoundScore As Integer
Dim SearchStr As String


  
  RoundNum = InputBox("Enter the Round Number that you want printed")
    If RoundNum > NumRounds Then
      MsgBox ("Round Number is greater than the number of Rounds")
      Exit Sub
    End If
 
  If CurrentContestType = "F3B" Then
    SQLString1 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Pilot.Res1 ,Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.RawScore, Scores.Task, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & "))ORDER BY Scores.Slot ASC,Pilot.Pilot_ID,Scores.Task ASC;"
  ElseIf CurrentContestType = "AustOpen" Then
    SQLString1 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.First_Name, Pilot.Pilot_ID, Pilot.Last_Name, Pilot.Res1 ,Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.RawScore, Scores.Task, Scores.Res2 FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & "))ORDER BY Scores.Slot ASC,Pilot.Pilot_ID,Scores.Score DESC;"
  Else
    SQLString1 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.First_Name, Pilot.Pilot_ID, Pilot.Last_Name, Pilot.Res1 ,Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.RawScore, Scores.Task FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Round=" & RoundNum & "))ORDER BY Scores.Slot ASC,Pilot.Pilot_ID,Scores.Score DESC;"
  End If
  Set PilotSet = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Linecount = Linecount + 1
  Open OutputFile For Output As #1
  Print #1, "<HTML>"
  Print #1, "<HEAD>"
  Print #1, "<TITLE>"
  Print #1, "Scores for Round"; " "; Str(RoundNum)
  Print #1, "</TITLE>"
  Print #1, "</HEAD>"
  Print #1, "<BODY>"
  Print #1, "<P ALIGN=Center><BIG><BIG>Scores for Round"; " "; Str(RoundNum); "<BIG></BIG>"
  Print #1, "<p>"
  Print #1, "<TABLE>"
  OutString = "<TABLE BORDER WIDTH=" & """80%""" & ">"
  Print #1, OutString
  Print #1, "<TR>"
  OutString = "<TH Width =" & """2%""" & ">Slot</TH>"
  Print #1, OutString
  OutString = "<TH Width =" & """18%""" & ">Name</TH>"
  Print #1, OutString
  OutString = "<TH Width =" & """10%""" & ">Score</TH>"
  Print #1, OutString
  
 
    
  If CurrentContestType = "F3B" Then
    OutString = "<TD Width =" & """5%""" & "></TD>"
    Print #1, OutString
    OutString = "<TH Width =" & """10%""" & ">Score</TH>"
    Print #1, OutString
    OutString = "<TH Width =" & """10%""" & ">Time</TH>"
    Print #1, OutString
    Print #1, "</TR>"
  Else
    OutString = "<TH Width =" & """10%""" & ">Time</TH>"
    Print #1, OutString
    OutString = "<TH Width =" & """10%""" & ">Land</TH>"
    Print #1, OutString
    OutString = "<TH Width =" & """10%""" & ">Pen</TH>"
    Print #1, OutString
    Print #1, "</TR>"
  End If
  
  If CurrentContestType = "F3B" Then
    PilotSet.MoveFirst
    Do Until PilotSet.EOF
      Print #1, "<TR ALIGN=Center>"
      OutString = "<TD Width =" & """2%""" & ">" & Str(PilotSet!Slot) & "</TD>"
      Print #1, OutString
      OutString = "<TD ALIGN=Left Width =" & """18%""" & ">" & Trim(PilotSet!First_Name) & " " & Trim(PilotSet!Last_Name) & " " & Trim(PilotSet!Res1) & "</TD>"
      Print #1, OutString
      OutString = "<TD Width =" & """10%"">" & Str(Listset!TempScore) & "</TD>"
      Print #1, OutString
      OutString = "<TD Width =" & """5%""" & ">A</TD>"
      Print #1, OutString
      OutString = "<TD Width =" & """10%"">" & Str(PilotSet!Score) & "</TD>"
      Print #1, OutString
      Mins = PilotSet!Flight_Time \ 60
      Secs = PilotSet!Flight_Time Mod 60
      OutString = "<TD Width =" & """10%"">" & Str(Mins) & ":" & Str(Secs) & "</TD>"
      Print #1, OutString
      Print #1, "</TR>"
      'Skip the first 4 Columns
      Print #1, "<TR ALIGN=Center>"
      OutString = "<TD Width =" & """2%""" & "></TD>"
      Print #1, OutString
      OutString = "<TD Width =" & """18%""" & "></TD>"
      Print #1, OutString
      OutString = "<TD Width =" & """10%""" & "></TD>"
      Print #1, OutString
      OutString = "<TD Width =" & """5%""" & ">B</TD>"
      Print #1, OutString
      OutString = "<TD Width =" & """10%"">" & Str(PilotSet!Score) & "</TD>"
      Print #1, OutString
      OutString = "<TD Width =" & """10%"">" & Str(PilotSet!Flight_Time) & "</TD>"
      Print #1, OutString
      Print #1, "</TR>"
      PilotSet.FindNext SearchStr
      'Skip the first 4 Columns
      Print #1, "<TR ALIGN=Center>"
      OutString = "<TD Width =" & """2%""" & "></TD>"
      Print #1, OutString
      OutString = "<TD Width =" & """18%""" & "></TD>"
      Print #1, OutString
      OutString = "<TD Width =" & """10%"">" & "</TD>"
      Print #1, OutString
      OutString = "<TD Width =" & """5%""" & ">C</TD>"
      Print #1, OutString
      OutString = "<TD Width =" & """10%"">" & Str(PilotSet!Score) & "</TD>"
      Print #1, OutString
      Mins = PilotSet!Flight_Time \ 100
      Secs = PilotSet!Flight_Time Mod 100
      OutString = "<TD Width =" & """10%"">" & Str(Mins) & ":" & Str(Secs) & "</TD>"
      Print #1, OutString
      Print #1, "</TR>"
      If PilotSet.EOF <> True Then
        PilotSet.MoveNext
      End If
    Loop
  Else
    'F3j printing part
    PilotSet.MoveFirst
    Do Until PilotSet.EOF
      Print #1, "<TR ALIGN=Center>"
      OutString = "<TD Width =" & """2%""" & ">" & Str(PilotSet!Slot) & "</TD>"
      Print #1, OutString
      OutString = "<TD ALIGN=Left Width =" & """18%""" & ">" & Trim(PilotSet!First_Name) & " " & Trim(PilotSet!Last_Name) & " " & Trim(PilotSet!Res1) & "</TD>"
      Print #1, OutString
      OutString = "<TD Width =" & """10%"">" & Str(PilotSet!Score) & "</TD>"
      Print #1, OutString
      Mins = PilotSet!Flight_Time \ 60
      Secs = PilotSet!Flight_Time - ((PilotSet!Flight_Time \ 60) * 60)
      OutString = "<TD Width =" & """10%"">" & Str(Mins) & ":" & Str(Secs) & "</TD>"
      Print #1, OutString
      OutString = "<TD Width =" & """10%"">" & Str(PilotSet!LandBonus) & "</TD>"
      Print #1, OutString
      OutString = "<TD Width =" & """10%"">" & Str(PilotSet!Penalty) & "</TD>"
      Print #1, OutString
      If PilotSet.EOF <> True Then
        PilotSet.MoveNext
      End If
      Print #1, "</TR>"
    Loop
  End If
  Print #1, "</TABLE>"
  Print #1, "</BODY>"
  Print #1, "</HTML>"
  'do housekeeping
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, List.TempScore From List WHERE (List.Comp_ID=" & Str(CurrentContest) & ") ORDER BY List.Pilot_ID;"
  Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Listset.MoveFirst
  Do Until Listset.EOF
    Listset.Edit
    Listset!TempScore = 0
    Listset.Update
    Listset.MoveNext
  Loop
  
  PilotSet.Close
  Close #1
  MsgBox ("Report Printed")
End Sub

Public Sub PrintSummaryHTML()
Dim Listset As Recordset
Dim Scoreset As Recordset
Dim SQLString1 As String
Dim rounds As Integer
Dim I As Integer
Dim Place As Integer
Dim Tab1 As Integer
Dim Count As Integer
Dim sco() As Single
Dim maxnum As Integer
Dim tabset(60) As Integer
Dim NumPilots As Integer
Dim MinScore As Single
Dim Linecount As Integer
Dim Col As Integer
Dim P As Integer
Dim Mins As Integer
Dim Secs As Single
Dim J As Integer
Dim Pagewidth As Integer
Dim LanPageLength As Integer
Dim OutString As String
Dim Colwidth As String
Dim Coltext As String

  
  Open OutputFile For Output As #1
  LanPageLength = 50

  Print #1, "<HTML>"
  Print #1, "<HEAD>"
  Print #1, "<TITLE>"
  Print #1, "Full Summary of Contest"
  Print #1, "</TITLE>"
  Print #1, "</HEAD>"
  Print #1, "<BODY>"
  Print #1, "<P ALIGN=Center><BIG><BIG>Full Summary of Contest<BIG></BIG>"
  Print #1, "<p>"
  Print #1, "<TABLE>"
  If NumRounds < 5 Then
    Colwidth = 70 / (NumRounds * 3)
    OutString = "<TABLE BORDER WIDTH=" & """100%""" & ">"
    Print #1, OutString
  Else
    Colwidth = 60 / (NumRounds * 3)
    OutString = "<TABLE BORDER WIDTH=" & """130%""" & ">"
    Print #1, OutString
  End If
  rounds = NumRounds
  Call ScoreAddUp(sco(), NumPilots)
  SummHeaderHTML Colwidth
  Count = 1
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Pilot.Res1 FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID = " & Str(CurrentContest) & ")) ORDER BY List.TempScore DESC;"
  Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Listset.MoveFirst
  Do Until Listset.EOF
    SQLString1 = "SELECT * From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Pilot_ID=" & Str(Listset!Pilot_ID) & ")) ORDER BY Scores.Round,Scores.Task ASC;"
    Set Scoreset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
    Print #1, "<TR ALIGN=Center>"
    OutString = "<TD Width =" & """2%""" & ">" & Str(Count) & "</TD>"
    Print #1, OutString
    OutString = "<TD ALIGN=Left Width =" & """18%""" & ">" & Trim(Listset!First_Name) & " " & Trim(Listset!Last_Name) & " " & Trim(Listset!Res1) & "</TD>"
    Print #1, OutString
    OutString = "<TD Width =" & """5%""" & ">" & Str(sco(Listset!Pilot_ID, 0)) & "</TD>"
    Print #1, OutString
    If CurrentContestType = "F3B" Then
      OutString = "<TD Width =" & """5%""" & ">Score</TD>"
      Print #1, OutString
      Scoreset.MoveFirst
      
      Do Until Scoreset.EOF
        'print all the scores in row 3
        OutString = "<TD Width =""" & Trim(Str(Colwidth)) & "%"">" & "<SMALL>" & Str(Scoreset!Score) & "</SMALL>" & "</TD>"
        Print #1, OutString
        Scoreset.MoveNext
      Loop
        'print the second line of scores
        
      Scoreset.MoveFirst
      'skip over the first 4 columns
      Skip4cols "Time"
      Do Until Scoreset.EOF
        'Print the times across the row
         If Scoreset!task = "A" Then
          Mins = Scoreset!Flight_Time \ 60
          Secs = Scoreset!Flight_Time - ((Scoreset!Flight_Time \ 60) * 60)
          OutString = "<TD Width =""" & Trim(Str(Colwidth)) & "%"">" & "<SMALL>" & Str(Mins) & ":" & Str(Secs) & "</SMALL>" & "</TD>"
          Print #1, OutString
          
        ElseIf Scoreset!task = "B" Then
          OutString = "<TD Width =""" & Trim(Str(Colwidth)) & "%"">" & "<SMALL>" & Str(Scoreset!Flight_Time) & "</SMALL>" & "</TD>"
          Print #1, OutString
          
        ElseIf Scoreset!task = "C" Then
          Mins = Scoreset!Flight_Time \ 100
          Secs = Scoreset!Flight_Time Mod 100
          OutString = "<TD Width =""" & Trim(Str(Colwidth)) & "%"">" & "<SMALL>" & Str(Mins) & ":" & Str(Secs) & "</SMALL>" & "</TD>"
          Print #1, OutString
        
        End If
        Scoreset.MoveNext
      Loop
      'print out Total scores across the row
      'skip over the first 4 columns
      Skip4cols "Total"
      For I = 1 To NumRounds
        OutString = "<TD Width =""" & Trim(Str(Colwidth)) & "%"">" & "</TD>"
        Print #1, OutString
        OutString = "<TD Width =""" & Trim(Str(Colwidth)) & "%"">" & "<SMALL>" & Str(sco(Listset!Pilot_ID, I)) & "</SMALL>" & "</TD>"
        Print #1, OutString
        OutString = "<TD Width =""" & Trim(Str(Colwidth)) & "%"">" & "</TD>"
        Print #1, OutString
      Next I
    
    Else
     'F3j printing part
      OutString = "<TD Width =" & """8%""" & ">Flight</TD>"
      Print #1, OutString
      Scoreset.MoveFirst
      
      Do Until Scoreset.EOF
        Mins = Scoreset!Flight_Time \ 60
        Secs = Scoreset!Flight_Time - ((Scoreset!Flight_Time \ 60) * 60)
        OutString = "<TD Width =""" & Trim(Str(Colwidth)) & "%"">" & "<SMALL>" & Str(Mins) & ":" & Str(Secs) & "</SMALL>" & "</TD>"
        Print #1, OutString
        OutString = "<TD Width =""" & Trim(Str(Colwidth)) & "%"">" & "<SMALL>" & Str(Scoreset!LandBonus) & "</SMALL>" & "</TD>"
        Print #1, OutString
        OutString = "<TD Width =""" & Trim(Str(Colwidth)) & "%"">" & "<SMALL>" & Str(Scoreset!Penalty) & "</SMALL>" & "</TD>"
        Print #1, OutString
        Scoreset.MoveNext
      Loop
      
      'skip over the first 4 columns
      Skip4cols "Score"
      For I = 1 To NumRounds
        OutString = "<TD Width =""" & Trim(Str(Colwidth)) & "%"">" & "</TD>"
        Print #1, OutString
        OutString = "<TD Width =""" & Trim(Str(Colwidth)) & "%"">" & "<SMALL>" & Str(sco(Listset!Pilot_ID, I)) & "</SMALL>" & "</TD>"
        Print #1, OutString
        OutString = "<TD Width =""" & Trim(Str(Colwidth)) & "%"">" & "</TD>"
        Print #1, OutString
      Next I
    End If
   
    
    Count = Count + 1
    Listset.MoveNext
  Loop
  'Do housekeeping
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, List.TempScore From List WHERE (List.Comp_ID=" & Str(CurrentContest) & ") ORDER BY List.Pilot_ID;"
  Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Listset.MoveFirst
  Do Until Listset.EOF
    Listset.Edit
    Listset!TempScore = 0
    Listset.Update
    Listset.MoveNext
  Loop
  
  Print #1, "</TABLE>"
  Print #1, "</BODY>"
  Print #1, "</HTML>"
  Close #1
  MsgBox ("Report printed")

End Sub



Public Sub SummHeader(ByRef Linecount As Integer, tabset() As Integer, Pagewidth As Integer, Startround As Integer, Endround As Integer)
Dim Colwidth As Integer
Dim Thisline As Integer

Dim I As Integer

  Thisline = 0
  Printer.FontSize = 16
  Printer.CurrentX = Printer.ScaleWidth / 2 - TextWidth("Leader Board")
  Printer.Print "Contest Summary"
  Linecount = 4
  Printer.FontSize = 12
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.Print "Place"
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.CurrentX = TextWidth("g") * (tabset(1) + 7)
  Printer.Print "Name"
  Printer.CurrentX = TextWidth("g") * tabset(6)
  Printer.CurrentY = TextHeight("sample") * Linecount
  Printer.Print "Total"
  'Skipping a column here
  Colwidth = 10
  For I = Startround To Endround
    
    If CurrentContestType = "F3B" Then
      Printer.FontSize = 10
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.CurrentX = TextWidth("g") * tabset(Colwidth + 2)
      Printer.Print "Round"; I
      Printer.CurrentY = TextHeight("sample") * (Linecount + 1)
      Printer.CurrentX = TextWidth("g") * tabset(Colwidth)
      Printer.Print "  "; "A"
      Printer.CurrentY = TextHeight("sample") * (Linecount + 1)
      Printer.CurrentX = TextWidth("g") * tabset(Colwidth + 2)
      Printer.Print "  "; "B"
      Printer.CurrentY = TextHeight("sample") * (Linecount + 1)
      Printer.CurrentX = TextWidth("g") * tabset(Colwidth + 4)
      Printer.Print "  "; "C"
      Printer.CurrentX = TextWidth("g") * tabset(8)
      Printer.CurrentY = TextHeight("sample") * (Linecount + 1)
      Printer.Print "Task"
      Colwidth = Colwidth + 6
      If Colwidth > Pagewidth Then
        Thisline = Linecount
        Linecount = Linecount + 4
        Colwidth = 10
      End If
    Else
      Printer.FontSize = 10
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.CurrentX = TextWidth("g") * tabset(Colwidth + 2)
      Printer.Print "Round"; I
      Printer.CurrentY = TextHeight("sample") * (Linecount + 1)
      Printer.CurrentX = TextWidth("g") * tabset(Colwidth)
      Printer.Print " "; "Time"
      Printer.CurrentY = TextHeight("sample") * (Linecount + 1)
      Printer.CurrentX = TextWidth("g") * tabset(Colwidth + 2)
      Printer.Print " "; "Land"
      Printer.CurrentY = TextHeight("sample") * (Linecount + 1)
      Printer.CurrentX = TextWidth("g") * tabset(Colwidth + 4)
      Printer.Print " "; "Pen"
      Colwidth = Colwidth + 6
      If Colwidth > Pagewidth Then
        Thisline = Linecount
        Linecount = Linecount + 4
        Colwidth = 10
      End If
    End If
  Next I
  
  If Thisline > 0 Then 'Restore the line counter to the original
    Linecount = Thisline 'line a second line of score is needed
  End If
End Sub

Public Sub Skip4cols(Col4text)
 Dim OutString As String
 
 'skip over the first 4 columns
  Print #1, "<TR ALIGN=Center>"
  OutString = "<TD Width =" & """2%""" & "></TD>"
  Print #1, OutString
  OutString = "<TD Width =" & """18%""" & "></TD>"
  Print #1, OutString
  OutString = "<TD Width =" & """5%""" & "></TD>"
  Print #1, OutString
  OutString = "<TD Width =" & """5%""" & ">" & Col4text & "</TD>"
  Print #1, OutString
End Sub



