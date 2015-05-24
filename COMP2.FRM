VERSION 5.00
Begin VB.Form ContestEnter 
   BackColor       =   &H00C0C0C0&
   Caption         =   "ContestEntry"
   ClientHeight    =   4245
   ClientLeft      =   615
   ClientTop       =   840
   ClientWidth     =   6465
   FillColor       =   &H00C0C000&
   HelpContextID   =   280
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4245
   ScaleWidth      =   6465
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contest Details"
      Height          =   1935
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   6015
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Top             =   1560
         WhatsThisHelpID =   280
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   1200
         WhatsThisHelpID =   280
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   840
         WhatsThisHelpID =   280
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1920
         TabIndex        =   11
         Top             =   360
         WhatsThisHelpID =   280
         Width           =   2895
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Number of Rounds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         WhatsThisHelpID =   280
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   1200
         WhatsThisHelpID =   280
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   840
         WhatsThisHelpID =   280
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contest Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         WhatsThisHelpID =   280
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Type Of Contest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      WhatsThisHelpID =   280
      Width           =   6015
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "F3B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         WhatsThisHelpID =   280
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "AustOpen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   7
         Top             =   360
         WhatsThisHelpID =   280
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Club"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         WhatsThisHelpID =   280
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "F3JFO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         WhatsThisHelpID =   280
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "F3J"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         WhatsThisHelpID =   280
         Width           =   855
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   3720
      WhatsThisHelpID =   280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3720
      WhatsThisHelpID =   280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contest Entry Form"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   280
      Width           =   5535
   End
End
Attribute VB_Name = "ContestEnter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MyTable As Recordset
Public CurrentRoundNumber As Integer
Dim SelectedSet As Recordset
  
Private Sub Command1_Click()
  Dim I As Integer
  Dim Message As String
  Message = "Contest Updated OK"
  On Error GoTo errhandler
  
  If Text1.Text = "" Then
    MsgBox ("There must be an entry for the Contest Name")
  ElseIf Text2.Text = "" Then
    MsgBox ("You have not entered a start date for the contest")
  ElseIf Text3.Text = "" Then
    MsgBox ("You have not entered an end date for the contest")
  Else
    If FromChange = True Then
      SelectedSet.Edit
      If Val(Text4.Text) > CurrentRoundNumber And CompRoundsDone = True Then
        MsgBox (" You cannot increase the number of rounds without first" & Chr(13) & "deleting all the scores and re-allocating")
        Message = "Did Not Update Database with round change"
      Else
        If Text4.Text <> "" Then
            SelectedSet![NumRounds] = Text4.Text
            'Remove scores from database after round number is changed
            SQLString = "SELECT * FROM Scores WHERE ((Scores.Comp_ID = " & Str(CurrentContest) & ") AND (Scores.Round > " & Text4.Text & ")) ORDER BY Pilot_ID;"
            Set DeleteSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
            If DeleteSet.RecordCount = 0 Then
                MsgBox ("There are no scores to delete")
            Else
                MsgString = "This will remove scores from Round " & Text4.Text & " onward - Are you sure?"
                If MsgBox(MsgString, 1) = vbOK Then
                  DeleteSet.MoveFirst
                  Do Until DeleteSet.EOF
                  DeleteSet.Delete
                  DeleteSet.MoveNext
                  Loop
                  DeleteSet.Close
                End If
            End If
        End If
      End If
      If Trim(MyTable![Comp_Name]) <> Trim(Text1.Text) Then
        SelectedSet![Comp_Name] = Text1.Text
      End If
      SelectedSet![CompStartDate] = Text2.Text
      SelectedSet![CompEndDate] = Text3.Text
      For I = 0 To 4
        If Option1(I) = True Then
          SelectedSet!CompType = Option1(I).Caption
        End If
      Next I
      CurrentContestType = SelectedSet!CompType
      Form1.Text1.Text = SelectedSet!Comp_Name
      If CurrentContestType = "F3J" Then
        max = maxF3J
        Form1.F3B.Visible = False

      ElseIf CurrentContestType = "F3JFO" Then
        max = maxF3JFO
        Form1.F3B.Visible = False

      ElseIf CurrentContestType = "AustOpen" Then
        max = max811
        Form1.F3B.Visible = False

      ElseIf CurrentContestType = "F3B" Then
        max = maxF3B
        Form1.F3B.Visible = True
    
      ElseIf CurrentContestType = "Club" Then
        max = maxClub
        Form1.F3B.Visible = False
      
      End If
      NumRounds = SelectedSet!NumRounds
      SelectedSet.Update
      Form1.Text1.Text = SelectedSet!Comp_Name
      Form1.Text2.Text = SelectedSet!NumRounds
      MsgBox (Message)
      
    Else
      MyTable.AddNew
      MyTable![Comp_Name] = Text1.Text
      MyTable![CompStartDate] = Text2.Text
      MyTable![CompEndDate] = Text3.Text
      If Text4.Text <> "" Then
        MyTable![NumRounds] = Text4.Text
      End If
      For I = 0 To 4
        If Option1(I) = True Then
           MyTable!CompType = Option1(I).Caption
        End If
      Next I
      MyTable!RoundsDone = False
      MyTable.Update
      MyTable.Index = "Comp_Name"
      MyTable.Seek "=", Text1.Text
      CurrentContest = MyTable!Comp_ID
      CurrentContestType = MyTable!CompType
      Form1.Text1.Text = MyTable!Comp_Name
      If CurrentContestType = "F3J" Then
        max = maxF3J
        Form1.F3B.Visible = False

      ElseIf CurrentContestType = "F3JFO" Then
        max = maxF3JFO
        Form1.F3B.Visible = False

      ElseIf CurrentContestType = "AustOpen" Then
        max = max811
        Form1.F3B.Visible = False

      ElseIf CurrentContestType = "F3B" Then
        max = maxF3B
        Form1.F3B.Visible = True
    
      ElseIf CurrentContestType = "Club" Then
        max = maxClub
        Form1.F3B.Visible = False
      
      End If
      NumRounds = MyTable!NumRounds
      NumRoundsDone = MyTable!RoundsScored
      CompRoundsDone = MyTable!RoundsDone
      FlightGroups = MyTable!FlightGroups
      'MyTable!Res1 = Str(0)
      'MyTable!Res2 = 0
      MsgBox ("The Contest is the Current Contest")
    
    End If
    If MyTable.BOF = True Then
      MyTable.MoveFirst
    ElseIf MyTable.EOF = True Then
      MyTable.MoveLast
    End If
    Unload Me
  End If
  
  Exit Sub
 
errhandler:
  MsgBox ("There was an error with one of datafields")
  Unload Me
End Sub

  

Private Sub Command3_Click()
  FromChange = False
  Unload Me
End Sub

Private Sub Form_Load()
  
  Dim SQLString As String
  
  Set MyTable = F3JDb.OpenRecordset("Competition", dbOpenTable)
  
  MyTable.Index = "Comp_ID"
  Command1.Caption = "Save"
  Command3.Caption = "Cancel"
  Option1(0) = True
  If FromChange = True Then
    'SQL the contest and display
    SQLString = "SELECT * FROM Competition WHERE Comp_ID =" & EditContestNumber
    Set SelectedSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
    Text1.Text = SelectedSet!Comp_Name
    Text2.Text = SelectedSet![CompStartDate]
    Text3.Text = SelectedSet![CompEndDate]
    Text4.Text = SelectedSet!NumRounds
    CurrentRoundNumber = SelectedSet!NumRounds
    If SelectedSet![CompType] = "F3J" Then
      Option1(0).Value = True
    ElseIf SelectedSet![CompType] = "F3B" Then
      Option1(4).Value = True
    ElseIf SelectedSet![CompType] = "F3JFO" Then
      Option1(1).Value = True
    ElseIf SelectedSet![CompType] = "Club" Then
      Option1(2).Value = True
      Text4.Text = ClubFlights
    ElseIf SelectedSet![CompType] = "AustOpen" Then
     Option1(3).Value = True
    End If
   End If
End Sub

Private Sub Label6_Click()
End Sub


Private Sub Option1_Click(Index As Integer)
  If Option1(1).Value = True Then
    'Text4.Text = 2
  ElseIf Option1(2).Value = True Then
    Text4.Text = ClubFlights
  
  End If
  
End Sub


