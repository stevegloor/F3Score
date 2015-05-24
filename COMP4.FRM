VERSION 5.00
Begin VB.Form ContestSelect 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Contest Select"
   ClientHeight    =   4395
   ClientLeft      =   465
   ClientTop       =   945
   ClientWidth     =   6465
   FillColor       =   &H00C0C0C0&
   HelpContextID   =   300
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4395
   ScaleWidth      =   6465
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3600
      TabIndex        =   20
      Text            =   "Text5"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Text            =   "Text4"
      Top             =   2280
      Width           =   375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contest Details"
      Height          =   1575
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   5895
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
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "End Date"
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
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   1080
         WhatsThisHelpID =   300
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Start Date"
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
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   720
         WhatsThisHelpID =   300
         Width           =   1215
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
         Left            =   240
         TabIndex        =   15
         Top             =   240
         WhatsThisHelpID =   300
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
      TabIndex        =   5
      Top             =   2760
      WhatsThisHelpID =   300
      Width           =   5895
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
         TabIndex        =   10
         Top             =   480
         WhatsThisHelpID =   300
         Width           =   855
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
         TabIndex        =   9
         Top             =   480
         WhatsThisHelpID =   300
         Width           =   855
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
         TabIndex        =   8
         Top             =   480
         WhatsThisHelpID =   300
         Width           =   1215
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
         Left            =   3480
         TabIndex        =   7
         Top             =   480
         WhatsThisHelpID =   300
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aust. Open"
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
         Height          =   495
         Index           =   3
         Left            =   4680
         TabIndex        =   6
         Top             =   360
         WhatsThisHelpID =   300
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   3840
      WhatsThisHelpID =   300
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      WhatsThisHelpID =   300
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3840
      WhatsThisHelpID =   300
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   3840
      WhatsThisHelpID =   300
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rounds set up in database"
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Number of Rounds"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contest Select Form"
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
      Left            =   600
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   300
      Width           =   5535
   End
End
Attribute VB_Name = "ContestSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyTable As Recordset
  
Private Sub Command1_Click()
  MyTable.MoveNext
  If MyTable.EOF = True Then
    MsgBox ("This is the last record")
    MyTable.MoveLast
    
  End If
  Text1.Text = MyTable!Comp_Name
  Text2.Text = MyTable![CompStartDate]
  Text3.Text = MyTable![CompEndDate]
  If MyTable!roundsdone = True Then
    Text5.Text = "YES"
  Else
    Text5.Text = "NO"
  End If
  Text4.Text = MyTable!NumRounds
  If MyTable![CompType] = "F3J" Then
     Option1(0).Value = True
  ElseIf MyTable![CompType] = "F3JFO" Then
     Option1(1).Value = True
  ElseIf MyTable![CompType] = "Club" Then
     Option1(2).Value = True
  ElseIf MyTable![CompType] = "AustOpen" Then
     Option1(3).Value = True
  ElseIf MyTable![CompType] = "F3B" Then
     Option1(4).Value = True
  End If
  
  End Sub

  

Private Sub Command2_Click()
  MyTable.MovePrevious
  If MyTable.BOF = True Then
    MsgBox ("This is the First Record")
    MyTable.MoveFirst
    
  End If
  Text1.Text = MyTable!Comp_Name
  Text2.Text = MyTable![CompStartDate]
  Text3.Text = MyTable![CompEndDate]
  If MyTable!roundsdone = True Then
    Text5.Text = "YES"
  Else
    Text5.Text = "NO"
  End If
  Text4.Text = MyTable!NumRounds
  If MyTable![CompType] = "F3J" Then
     Option1(0).Value = True
  ElseIf MyTable![CompType] = "F3JFO" Then
     Option1(1).Value = True
  ElseIf MyTable![CompType] = "Club" Then
     Option1(2).Value = True
  ElseIf MyTable![CompType] = "AustOpen" Then
     Option1(3).Value = True
  ElseIf MyTable![CompType] = "F3B" Then
     Option1(4).Value = True
  End If
End Sub


Private Sub Command3_Click()
  ReDim F3BNumSlots(3)
  CurrentContest = MyTable!Comp_ID
  NumRounds = MyTable!NumRounds
  CurrentContestType = MyTable!CompType
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
    'Form1.Comp.MovePilot.Visible = False
  
  ElseIf CurrentContestType = "Club" Then
    max = maxClub
    Form1.F3B.Visible = False
  
  End If
  
  CompRoundsDone = MyTable!roundsdone
  Form1.Text1.Text = MyTable!Comp_Name
  NumRoundsDone = MyTable!RoundsScored
  Completed = MyTable!Completed
  FlightGroups = MyTable!FlightGroups
  If MyTable!Completed = True Then
    Form1.Text3.Text = "YES"
  Else
    Form1.Text3.Text = "NO"
  End If
  Form1.Text2.Text = Str(NumRounds)
  If CurrentContestType = "F3B" Then
    F3BNumSlots(1) = MyTable!NumSlots
    If MyTable!Res1 <> "" Then
      F3BNumSlots(2) = Val(MyTable!Res1)
    End If
    F3BNumSlots(3) = MyTable!Res2
    NumSlots = MyTable!NumSlots
  Else
    NumSlots = MyTable!NumSlots
    MyTable.Edit
  End If
  MsgBox ("Contest Selected")
  Form1.Cleargrid
  MyTable.Close
  Unload Me
 
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Command4_Click()
  MyTable.Close
  Unload Me
End Sub

Private Sub Form_Load()
  Set MyTable = F3JDb.OpenRecordset("Competition", dbOpenTable)
  
  MyTable.Index = "Comp_ID"
  Command1.Caption = "Next >>"
  Command2.Caption = "<< Prev"
  Command3.Caption = "Select Contest"
  Command4.Caption = "Cancel"
  
  If MyTable.BOF = True Then
     MyTable.MoveFirst
  ElseIf MyTable.EOF = True Then
     MyTable.MoveLast
  End If
  Text1.Text = MyTable!Comp_Name
  Text2.Text = MyTable![CompStartDate]
  Text3.Text = MyTable![CompEndDate]
  If MyTable!roundsdone = True Then
    Text5.Text = "YES"
  Else
    Text5.Text = "NO"
  End If
  Text4.Text = MyTable!NumRounds
  If MyTable![CompType] = "F3J" Then
     Option1(0).Value = True
  ElseIf MyTable![CompType] = "F3JFO" Then
     Option1(1).Value = True
  ElseIf MyTable![CompType] = "CLUB" Then
     Option1(2).Value = True
  ElseIf MyTable![CompType] = "AustOpen" Then
     Option1(3).Value = True
  ElseIf MyTable![CompType] = "F3B" Then
     Option1(4).Value = True
  End If
  
End Sub


