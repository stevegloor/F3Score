VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form EnterClubScores 
   Caption         =   "Club Score Input Form"
   ClientHeight    =   4545
   ClientLeft      =   915
   ClientTop       =   945
   ClientWidth     =   6690
   LinkTopic       =   "Form4"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4545
   ScaleWidth      =   6690
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Score Pilot"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Flight Times for Selected Pilot"
      Height          =   2655
      Left            =   240
      TabIndex        =   1
      Top             =   720
      WhatsThisHelpID =   310
      Width           =   6015
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1815
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3201
         _Version        =   393216
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Total Score"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Time Remaining"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Total Time"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   3480
      Width           =   855
   End
End
Attribute VB_Name = "EnterClubScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelectedSlot As Recordset
Private Sub Combo1_Click()

Dim NumPilots As Integer
Dim NumScores As Integer
Dim SQLString3 As String
Dim Mins As Integer
Dim Sec As Single
Dim Count As Integer
Dim Col As Integer
Dim Row As Integer
Dim Pilot As String
Dim TotalScore As Integer
Dim TimeRemain As Integer

  If Combo1.Text <> " - Select Pilot -" Then
    Pilot = Trim(Combo1.Text)
    Pilot = Left(Pilot, 3)
    Pilot = Trim(Pilot)
    SQLString3 = "SELECT DISTINCTROW Scores.Comp_ID, Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, Scores.Round, Scores.Slot, Scores.Flight_Time, Scores.Score, Scores.LandBonus, Scores.Penalty, Scores.RawScore FROM Pilot INNER JOIN Scores ON Pilot.Pilot_ID = Scores.Pilot_ID WHERE ((Scores.Comp_ID=" & Str(CurrentContest) & ") AND (Scores.Pilot_ID =" & Pilot & ")) ORDER BY Scores.Round;"
    Set SelectedSlot = F3JDb.OpenRecordset(SQLString3, dbOpenDynaset)
    If SelectedSlot.RecordCount <> 0 Then
      NumPilots = SelectedSlot.RecordCount
      SetGrid NumPilots
      SelectedSlot.MoveFirst
      FillGrid SelectedSlot
      'If there are any scores then add them up
      SelectedSlot.MoveFirst
      If SelectedSlot!Score <> 0 Then
        TotalScore = 0
        TimeRemain = 0
        Do Until SelectedSlot.EOF
          TotalScore = TotalScore + SelectedSlot!Score
          SelectedSlot.MoveNext
        Loop
        Text4.Text = TotalScore
        TimeRemain = (ClubTime * 60) - TotalScore
        Mins = TotalScore \ 60
        Secs = TotalScore Mod 60
        Text1.Text = Str(Mins) & ":" & Str(Secs)
        Mins = TimeRemain \ 60
        Secs = TimeRemain Mod 60
        Text2.Text = Str(Mins) & ":" & Str(Secs)
      End If
    Else
      MsgBox ("No Pilots in Slot")
      Unload Me
      Exit Sub
    End If
  End If

End Sub


Private Sub Combo1_DropDown()
Dim TeamSelect As Recordset
Dim SQLString As String
Dim Line As String

 Combo1.Clear
 SQLString = "SELECT DISTINCTROW Pilot.Pilot_ID, Pilot.First_Name, Pilot.Last_Name, List.InTeam, List.Team_ID, List.Comp_ID FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID WHERE ((List.Comp_ID=" & Str(CurrentContest) & ")) ORDER BY Pilot.Last_Name;"
 Set TeamSelect = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
 If TeamSelect.RecordCount <> 0 Then
   TeamSelect.MoveFirst
   
 End If
   
 Do Until TeamSelect.EOF = True
   Line = Format(Str(TeamSelect!Pilot_ID), "!@@@@") & "  " & TeamSelect!First_Name & " " & TeamSelect!Last_Name
   Combo1.AddItem Line
   TeamSelect.MoveNext
 Loop
 TeamSelect.Close
End Sub


Private Sub Command1_Click()
  ClubScore SelectedSlot
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Command3_Click()
  PrintForm
End Sub

Private Sub Form_Load()

Dim NumPilots As Integer

  NumPilots = 5
  SetGrid NumPilots
  Combo1.Text = " - Select Pilot -"

End Sub


Private Sub msflexgrid1_SelChange()

Dim Pcol As Integer
Dim Prow As Integer

If Updated = True Then
  If Text3.Text <> "" Then
    Pcol = MSFlexGrid1.Col
    Prow = MSFlexGrid1.Row
    MSFlexGrid1.Row = Srow
    MSFlexGrid1.Col = Scol
    If MSFlexGrid1.Col = 2 Then
      If Val(Text3.Text) > 60 Then
        MsgBox ("Entry must be less that 60")
        Text3.Text = ""
        MSFlexGrid1.Row = Prow
        MSFlexGrid1.Col = Pcol
        Exit Sub
      End If
    End If
    If MSFlexGrid1.Col = 3 Then
      If Val(Text3.Text) > 100 Then
        MsgBox ("You cannot enter more than 100 points")
        Text3.Text = ""
        MSFlexGrid1.Row = Prow
        MSFlexGrid1.Col = Pcol
      End If
    End If
    MSFlexGrid1.Text = Text3.Text
    Text3.Text = ""
    MSFlexGrid1.Row = Prow
    MSFlexGrid1.Col = Pcol
  End If
End If
End Sub


Private Sub msflexgrid1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Text3.Text <> "" Then
      MSFlexGrid1.Text = Text3.Text
    End If
  ElseIf KeyAscii = 27 Then
    Text3.Text = ""
  Else
    Text3.Text = Text3.Text & Chr(KeyAscii)
    Srow = MSFlexGrid1.Row
    Scol = MSFlexGrid1.Col
  End If
End Sub
Private Sub msflexgrid1_DblClick()
  MSFlexGrid1.Col = 6
  If MSFlexGrid1.Text = "" Then
    MSFlexGrid1.Text = "*"
    For R = 1 To (MSFlexGrid1.Rows - 1)
      MSFlexGrid1.Row = R
      MSFlexGrid1.Text = ""
    Next R
    If MsgBox("Do you want all the scores cleared", 19) = vbYes Then
      ClearScores
    End If
      
  Else
    MSFlexGrid1.Text = ""
  End If
End Sub
Sub SetGrid(NumPilots As Integer)
  Dim Col2Fill
  Dim Row2Fill
  Dim NL As String
  Dim GridHgt
  Dim GridWid
  Dim I
  Updated = False
  
  GridHgt = MSFlexGrid1.Height
  GridWid = MSFlexGrid1.Width
  
  MSFlexGrid1.Cols = 7
  MSFlexGrid1.Colwidth(0) = GridWid / 3
  
  GridWid = GridWid - MSFlexGrid1.Colwidth(0) - 60
  MSFlexGrid1.Colwidth(6) = (GridWid / 6)
  
  MSFlexGrid1.Colwidth(2) = (GridWid / 6)
  MSFlexGrid1.Colwidth(3) = (GridWid / 6)
  MSFlexGrid1.Colwidth(4) = (GridWid / 6)
  MSFlexGrid1.Colwidth(5) = (GridWid / 6)
  MSFlexGrid1.Colwidth(6) = (GridWid / 6)
      
    
  MSFlexGrid1.Rows = (NumPilots + 1)
  If NumPilots > 7 Then
    MSFlexGrid1.ScrollBars = 2
  Else
    MSFlexGrid1.ScrollBars = 0
  End If
  MSFlexGrid1.RowHeight(0) = GridHgt / 7
  For I = 1 To NumPilots
    MSFlexGrid1.RowHeight(I) = GridHgt / 9
  Next I
 ' Clear the contents of the grid.
  Cleargrid
  'Initialize the line break variable.
  NL = Chr(10) & Chr(13)
  ' First row = current row.
  MSFlexGrid1.Row = 0
  MSFlexGrid1.Col = 0
  MSFlexGrid1.FixedAlignment(0) = 0
  MSFlexGrid1.Text = "Flight No"
  
  MSFlexGrid1.Row = 0
  'Put column headings in the first row.
  MSFlexGrid1.Col = 1
  MSFlexGrid1.FixedAlignment(1) = 2
  MSFlexGrid1.Text = "Mins"
  MSFlexGrid1.Col = 2
  MSFlexGrid1.FixedAlignment(2) = 2
  MSFlexGrid1.Text = "Sec"
  MSFlexGrid1.Col = 3
  MSFlexGrid1.FixedAlignment(3) = 2
  MSFlexGrid1.Text = "Land"
  MSFlexGrid1.Col = 4
  MSFlexGrid1.FixedAlignment(4) = 2
  MSFlexGrid1.Text = "Pen"
  MSFlexGrid1.Col = 5
  MSFlexGrid1.FixedAlignment(5) = 2
  MSFlexGrid1.Text = "Score"
  MSFlexGrid1.Col = 6
  MSFlexGrid1.FixedAlignment(6) = 0
  
  
  For I = 1 To 5
    MSFlexGrid1.ColAlignment(I) = 2
  Next I
  
  Updated = True
End Sub

Public Sub FillGrid(SelectedSlot)
Dim Row As Integer
Dim Mins As Integer
Dim Secs As Single
Dim Name As String

Updated = False
SelectedSlot.MoveFirst
Row = 1
Do Until SelectedSlot.EOF
  MSFlexGrid1.Row = Row
  Mins = SelectedSlot!Flight_Time \ 60
  Sec = SelectedSlot!Flight_Time - ((SelectedSlot!Flight_Time \ 60) * 60)
  MSFlexGrid1.Col = 1
  MSFlexGrid1.Text = Trim(Str(Mins))
  MSFlexGrid1.Col = 2
  MSFlexGrid1.Text = Trim(Str(Sec))
  MSFlexGrid1.Col = 3
  MSFlexGrid1.Text = Trim(Str(SelectedSlot!LandBonus))
  MSFlexGrid1.Col = 4
  MSFlexGrid1.Text = Trim(Str(SelectedSlot!Penalty))
  MSFlexGrid1.Col = 0
  MSFlexGrid1.Text = Trim(Str(SelectedSlot!Round))
  If SelectedSlot!Score > 0 Then
    MSFlexGrid1.Col = 5
    MSFlexGrid1.Text = Trim(Str(SelectedSlot!Score))
  Else
    MSFlexGrid1.Col = 5
    MSFlexGrid1.Text = "0"
  End If
  Row = Row + 1
  SelectedSlot.MoveNext
Loop
Updated = True
End Sub

Public Sub ClubScore(ByVal SelectedSlot)

Dim String1 As String
Dim Count As Integer
Dim Penalty(20) As Integer
Dim LandBonus As Boolean
Dim ManualPenalty
Dim Row As Integer
Dim FlightTime(20) As Integer
Dim Column As Integer
Dim MaxScore As Single
Dim ValLandBonus(20)
Dim Rawscore(20) As Integer
Dim Tagged As Boolean
Dim Refly As Boolean
Dim Mins As Integer
Dim Secs As Single
Dim Pilots As Integer
Dim TotalScore As Integer
Dim TimeRemain As Integer

  
  Pilots = NumPilots
  Row = 0
  MSFlexGrid1.Row = Row
  NumPilots = SelectedSlot.RecordCount
  Do Until Row = NumRounds
    Row = Row + 1

    MSFlexGrid1.Row = Row
    LandBonus = True
    
    MSFlexGrid1.Col = 1
    FlightTime(Row) = Val(MSFlexGrid1.Text) * 60
    MSFlexGrid1.Col = 2
    FlightTime(Row) = FlightTime(Row) + Val(Trim(MSFlexGrid1.Text))
   
    If FlightTime(Row) > (max * 60) Then
      LandBonus = False
      If Trim(MSFlexGrid1.Text) <> "0" Then
        Penalty(Row) = Penalty(Row) + Val(MSFlexGrid1.Text)
      End If
      Penalty(Row) = (FlightTime(Row) - (max * 60)) * 2
    End If
    MSFlexGrid1.Col = 3
    If Val(MSFlexGrid1.Text) < 30 Then
     If LandBonus = True Then
       MSFlexGrid1.Col = 3
       Select Case Val(MSFlexGrid1.Text)
       Case 1
         ValLandBonus(Row) = 100
       Case 2
         ValLandBonus(Row) = 95
       Case 3
         ValLandBonus(Row) = 90
       Case 4
         ValLandBonus(Row) = 85
       Case 5
         ValLandBonus(Row) = 80
       Case 6
         ValLandBonus(Row) = 75
       Case 7
         ValLandBonus(Row) = 70
       Case 8
         ValLandBonus(Row) = 65
       Case 9
         ValLandBonus(Row) = 60
       Case 10
         ValLandBonus(Row) = 55
       Case 11
         ValLandBonus(Row) = 50
       Case 12
         ValLandBonus(Row) = 45
       Case 13
         ValLandBonus(Row) = 40
       Case 14
         ValLandBonus(Row) = 34
       Case 15
         ValLandBonus(Row) = 30
       Case Else
         ValLandBonus(Row) = 0
       End Select
       
     Else
       ValLandBonus(Row) = 0
        
     End If
   Else
      MSFlexGrid1.Col = 3
      ValLandBonus(Row) = Val(MSFlexGrid1.Text)
   End If
   Rawscore(Row) = FlightTime(Row) + ValLandBonus(Row) - Penalty(Row) - ManualPenalty
  Loop
   
  SelectedSlot.MoveFirst
  Row = 1
  TotalScore = 0
  TimeRemain = 0
  Do Until SelectedSlot.EOF
    SelectedSlot.Edit
    SelectedSlot!Score = CInt(Rawscore(Row))
    SelectedSlot!Rawscore = CInt(Rawscore(Row))
    SelectedSlot!Flight_Time = FlightTime(Row)
    SelectedSlot!LandBonus = ValLandBonus(Row)
    SelectedSlot!Penalty = Penalty(Row)
    TotalScore = TotalScore + Rawscore(Row)
    If TotalScore > (ClubTime * 60) Then
      SelectedSlot!Score = (TotalScore - (ClubTime * 60)) * 2
    End If
    SelectedSlot.Update
    Row = Row + 1
    SelectedSlot.MoveNext
  Loop
  FillGrid SelectedSlot
  Text4.Text = TotalScore
  TimeRemain = (ClubTime * 60) - TotalScore
  Mins = TotalScore \ 60
  Secs = TotalScore Mod 60
  Text1.Text = Str(Mins) & ":" & Str(Secs)
  Mins = TimeRemain \ 60
  Secs = TimeRemain Mod 60
  Text2.Text = Str(Mins) & ":" & Str(Secs)
  
  
  
  MsgBox ("Pilot Scored")
  
End Sub

Public Sub ClearScores()
   ' Declare local variables.
    Dim Count As Integer
    ' Clear text from the first column.
    
   
    ' Clear text from the nonfixed cells.
    'MSFlexGrid1.SelStartCol = 1
    'MSFlexGrid1.SelStartRow = 1
    'MSFlexGrid1.SelEndCol = MSFlexGrid1.Cols - 2
    'MSFlexGrid1.SelEndRow = MSFlexGrid1.Rows - 1
    'MSFlexGrid1.FillStyle = 1
    'MSFlexGrid1.Text = "0"
    'MSFlexGrid1.FillStyle = 0
    'MSFlexGrid1.SelEndCol = 1
    'MSFlexGrid1.SelEndRow = 1
End Sub
Sub Cleargrid()
   ' Declare local variables.
    Dim Count As Integer
    ' Clear text from the first column.
    MSFlexGrid1.FixedCols = 1
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = " "
    MSFlexGrid1.Col = 0
    For Count = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = Count
        MSFlexGrid1.Text = " "
    Next Count
    ' Clear text from the first row.
    MSFlexGrid1.Row = 0
    For Count = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Col = Count
        MSFlexGrid1.Text = " "
    Next Count
    ' Clear text from the nonfixed cells.
    'MSFlexGrid1.SelStartCol = 1
    'MSFlexGrid1.SelStartRow = 1
    'MSFlexGrid1.SelEndCol = MSFlexGrid1.Cols - 1
    'MSFlexGrid1.SelEndRow = MSFlexGrid1.Rows - 1
    'MSFlexGrid1.FillStyle = 1
    'MSFlexGrid1.Text = ""
    'MSFlexGrid1.FillStyle = 0
    'MSFlexGrid1.SelEndCol = 1
    'MSFlexGrid1.SelEndRow = 1
End Sub
