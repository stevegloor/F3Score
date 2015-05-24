Attribute VB_Name = "Module1"
Option Explicit
Public F3JDb As Database
Public CurrentContest As Integer
Public blLocal As Boolean 'variable to check for local database
Public strDatabaseName As String 'to contain the database location
Public NumberPilots As Integer
Public DatabaseType As String
Public CurrentChamp As Integer
Public Currentpilot As String
Public Currenttask As String
Public ThisTask As Integer
Public Completed As Boolean
Public FromChange As Boolean
Public FromView As Boolean
Public FromEdit As Boolean
Public TxT As Integer
Public Cummlative As Boolean
Public NumRounds As Integer
Public CurrentContestType As String
Public CurrentRound As Integer
Public NumRoundsDone As Integer
Public NumSlots As Integer
Public F3BNumSlots() As Integer
Public datafile As String
Public maxtimes As Integer
Public MostUsedFreq As Integer
Public Finish As Boolean
Public CompRoundsDone As Boolean
Public max As Integer
Public EditPilotID As String
Public EditContestNumber As String
Public CongestFreq As String
Public alloc() As Integer
'Public allocB() As Integer
'Public allocC() As Integer
Public allocated As Boolean
Public RoundsCorrect As Integer
Public Scol As Integer
Public Srow As Integer
Public Updated As Boolean
Public Country As Integer
Public maxF3J As Integer
Public maxF3B As Integer
Public max811 As Integer
Public maxF3JFO As Integer
Public maxClub As Integer
Public ClubTime As Integer
Public ClubFlights As Integer
Public Lookup As Boolean
Public FlightGroups As Boolean
Public DisplayCountry As Boolean
Public TenK As Boolean
Public Place() As Integer
Public SelectedContestType As String
Public DiscardRound As Integer
Public LeaderboardDone As Boolean
Public ViewScores As Boolean
Public throw() As Single
Function ReadIniFile(ByVal IniFile, ByVal Section, ByVal Key) As String

End Function

Sub updateprogress(pb As Control, ByVal percent)
Dim num$        'use percent
If Not pb.AutoRedraw Then      'picture in memory ?
    pb.AutoRedraw = -1          'no, make one
    End If
    pb.Cls                      'clear picture in memory
    pb.ScaleWidth = 100         'new sclaemodus
    pb.DrawMode = 10            'not XOR Pen Modus
    num$ = Format$(percent, "###") + "%"
    pb.CurrentX = 50 - pb.TextWidth(num$) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(num$)) / 2
    pb.Print num$               'print percent
    pb.Line (0, 0)-(percent, pb.ScaleHeight), , BF
    pb.Refresh          'show differents

End Sub

Public Sub ScoreAddUp(ByRef sco() As Single, ByRef NumPilots)

Dim Listset As Recordset
Dim Scoreset As Recordset
Dim SQLString1 As String
Dim SQLString2 As String
Dim rounds As Integer
Dim I As Integer
Dim T As Integer
Dim J As Integer
Dim Count As Integer
Dim maxnum As Integer
Dim MinScore As Single
Dim Minscore2 As Single
Dim AMinscore As Single
Dim BMinscore As Single
Dim CMinscore As Single
Dim ThisTask As String
Dim Tempmin As Single
Dim ZeroNumber As Integer

'On Error GoTo errhandler
  ZeroNumber = 0
  rounds = NumRounds
  SQLString1 = "SELECT DISTINCTROW List.Comp_ID, List.Pilot_ID, List.TempScore From List WHERE (List.Comp_ID=" & Str(CurrentContest) & ") ORDER BY List.Pilot_ID ASC;"
  Set Listset = F3JDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Listset.MoveFirst
  NumPilots = Listset.RecordCount
  'Find the largest Pilot ID number
  Do Until Listset.EOF
    If Listset!Pilot_ID > maxnum Then
      maxnum = Listset!Pilot_ID
    End If
    Listset.MoveNext
  Loop
  Listset.MoveFirst
  ReDim sco(maxnum, rounds)
  ReDim throw(maxnum, 2)
  If CurrentContestType = "F3J" Then
    Do Until Listset.EOF
      MinScore = 1000
      SQLString2 = "SELECT DISTINCTROW Scores.Comp_ID, Scores.Pilot_ID, Scores.Round, Scores.Score, Scores.Task, Scores.Res2 From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Pilot_ID=" & Str(Listset!Pilot_ID) & ") AND (Scores.Task='A')) ORDER BY Scores.Round ASC;"
      Set Scoreset = F3JDb.OpenRecordset(SQLString2, dbOpenDynaset)
      Scoreset.MoveFirst
      I = 1
      J = 0
      Do
      'While J <> NumRounds
        If Scoreset!Score < MinScore Then
          MinScore = Scoreset!Score
        End If
        sco(Listset!Pilot_ID, I) = Scoreset!Score
        sco(Listset!Pilot_ID, 0) = sco(Listset!Pilot_ID, 0) + Scoreset!Score - Scoreset!Res2
        sco(Listset!Pilot_ID, 0) = Round(sco(Listset!Pilot_ID, 0), 1)
        I = I + 1
        J = J + 1
        Scoreset.MoveNext
      Loop Until Scoreset.EOF
      
      'Wend
      'Discard Lowest score if there are more that 5 rounds flown
      If J >= DiscardRound Then
        sco(Listset!Pilot_ID, 0) = sco(Listset!Pilot_ID, 0) - MinScore
        sco(Listset!Pilot_ID, 0) = Round(sco(Listset!Pilot_ID, 0), 1)
        throw(Listset!Pilot_ID, 0) = MinScore
      End If
      Listset.Edit
      Listset!TempScore = sco(Listset!Pilot_ID, 0)
      Listset.Update
      Listset.MoveNext
    Loop
  ElseIf CurrentContestType = "F3JFO" Then
    Do Until Listset.EOF
      MinScore = 1000
      SQLString2 = "SELECT DISTINCTROW Scores.Comp_ID, Scores.Pilot_ID, Scores.Round, Scores.Score, Scores.Task, Scores.Res2 From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Pilot_ID=" & Str(Listset!Pilot_ID) & ") AND (Scores.Task='A')) ORDER BY Scores.Round ASC;"
      Set Scoreset = F3JDb.OpenRecordset(SQLString2, dbOpenDynaset)
      Scoreset.MoveFirst
      I = 1
      J = 0
      Do
      'While J <> NumRounds
        If Scoreset!Score < MinScore Then
          MinScore = Scoreset!Score
        End If
        sco(Listset!Pilot_ID, I) = Scoreset!Score
        sco(Listset!Pilot_ID, 0) = sco(Listset!Pilot_ID, 0) + Scoreset!Score - Scoreset!Res2
        sco(Listset!Pilot_ID, 0) = Round(sco(Listset!Pilot_ID, 0), 1)
        I = I + 1
        J = J + 1
        Scoreset.MoveNext
      Loop Until Scoreset.EOF
      'Wend
      'Discard Lowest score if there are more that 5 rounds flown
      If J >= (DiscardRound - 2) Then
        sco(Listset!Pilot_ID, 0) = sco(Listset!Pilot_ID, 0) - MinScore
        sco(Listset!Pilot_ID, 0) = Round(sco(Listset!Pilot_ID, 0), 1)
        throw(Listset!Pilot_ID, 0) = MinScore
      End If
      Listset.Edit
      Listset!TempScore = sco(Listset!Pilot_ID, 0)
      Listset.Update
      Listset.MoveNext
    Loop
  ElseIf CurrentContestType = "F3B" Then
    Do Until Listset.EOF
      AMinscore = 1000
      BMinscore = 1000
      CMinscore = 1000
      For T = 1 To 3
        If T = 1 Then
            ThisTask = "A"
          ElseIf T = 2 Then
            ThisTask = "B"
          ElseIf T = 3 Then
            ThisTask = "C"
        End If
        SQLString2 = "SELECT DISTINCTROW Scores.Comp_ID, Scores.Pilot_ID, Scores.Round, Scores.Score, Scores.Task, Scores.Res2 From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Pilot_ID=" & Str(Listset!Pilot_ID) & ") AND (Scores.Task='" & ThisTask & "'))ORDER BY Scores.Round ASC;"
        Set Scoreset = F3JDb.OpenRecordset(SQLString2, dbOpenDynaset)
        Scoreset.MoveFirst
        I = 1
        J = 0
        Do
          If T = 1 Then
            If Scoreset!Score < AMinscore Then
              AMinscore = Scoreset!Score
            End If
          ElseIf T = 2 Then
            If Scoreset!Score < BMinscore Then
              BMinscore = Scoreset!Score
            End If
          ElseIf T = 3 Then
            If Scoreset!Score < CMinscore Then
              CMinscore = Scoreset!Score
            End If
          End If
        
          sco(Listset!Pilot_ID, I) = sco(Listset!Pilot_ID, I) + Scoreset!Score
          sco(Listset!Pilot_ID, 0) = sco(Listset!Pilot_ID, 0) + Scoreset!Score - Scoreset!Res2
          sco(Listset!Pilot_ID, 0) = Round(sco(Listset!Pilot_ID, 0), 0)
          Scoreset.MoveNext
          I = I + 1
          J = J + 1
        Loop Until Scoreset.EOF
      Next T
      'Discard Lowest scores if there are more that 5 rounds flown
      If J >= DiscardRound Then
        sco(Listset!Pilot_ID, 0) = sco(Listset!Pilot_ID, 0) - (AMinscore + BMinscore + CMinscore)
        sco(Listset!Pilot_ID, 0) = Round(sco(Listset!Pilot_ID, 0), 0)
        throw(Listset!Pilot_ID, 0) = AMinscore
        throw(Listset!Pilot_ID, 1) = BMinscore
        throw(Listset!Pilot_ID, 2) = CMinscore
      End If
      Listset.Edit
      Listset!TempScore = sco(Listset!Pilot_ID, 0)
      Listset.Update
      Listset.MoveNext
      
    Loop
  ElseIf CurrentContestType = "AustOpen" Then
    Do Until Listset.EOF
      MinScore = 1000
      Minscore2 = 1000
      Tempmin = 1000
      ZeroNumber = 0
      SQLString2 = "SELECT DISTINCTROW Scores.Comp_ID, Scores.Pilot_ID, Scores.Round, Scores.Score, Scores.Task, Scores.Res2 From Scores WHERE ((Scores.Comp_ID= " & Str(CurrentContest) & ") AND (Scores.Pilot_ID=" & Str(Listset!Pilot_ID) & ") AND (Scores.Task='A')) ORDER BY Scores.Round ASC;"
      Set Scoreset = F3JDb.OpenRecordset(SQLString2, dbOpenDynaset)
      Scoreset.MoveFirst
      I = 1
      J = 0
      Do
      'While J <> NumRounds
        If Scoreset!Score = 0 Then
          ZeroNumber = ZeroNumber + 1
        End If
        If Scoreset!Score < MinScore Then
          Tempmin = MinScore
          MinScore = Scoreset!Score
        ElseIf Scoreset!Score < Minscore2 Then
          Tempmin = Scoreset!Score
        End If
        
        If Tempmin < Minscore2 Then
          If Tempmin <> Minscore2 Then
            Minscore2 = Tempmin
          End If
        End If
        If ZeroNumber > 1 Then
          Minscore2 = 0
        End If
        sco(Listset!Pilot_ID, I) = Scoreset!Score
        sco(Listset!Pilot_ID, 0) = sco(Listset!Pilot_ID, 0) + Scoreset!Score - Scoreset!Res2
        sco(Listset!Pilot_ID, 0) = Round(sco(Listset!Pilot_ID, 0), 1)
        I = I + 1
        J = J + 1
        Scoreset.MoveNext
      'Wend
      Loop Until Scoreset.EOF
      'Discard Lowest score if there are more that 5 rounds flown
      If J >= DiscardRound * 2 Then
        sco(Listset!Pilot_ID, 0) = sco(Listset!Pilot_ID, 0) - (MinScore + Minscore2)
        sco(Listset!Pilot_ID, 0) = Round(sco(Listset!Pilot_ID, 0), 1)
        throw(Listset!Pilot_ID, 0) = MinScore
        throw(Listset!Pilot_ID, 1) = Minscore2
      ElseIf J >= DiscardRound Then
        sco(Listset!Pilot_ID, 0) = sco(Listset!Pilot_ID, 0) - MinScore
        sco(Listset!Pilot_ID, 0) = Round(sco(Listset!Pilot_ID, 0), 1)
        throw(Listset!Pilot_ID, 0) = MinScore
      End If
      Listset.Edit
      Listset!TempScore = sco(Listset!Pilot_ID, 0)
      Listset.Update
      Listset.MoveNext
    Loop
  
  End If
  Exit Sub
errhandler:
  MsgBox ("There is a problem with this database")
End Sub
