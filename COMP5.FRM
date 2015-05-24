VERSION 5.00
Begin VB.Form PilotAdd 
   Caption         =   "Form5"
   ClientHeight    =   5055
   ClientLeft      =   540
   ClientTop       =   1440
   ClientWidth     =   6390
   HelpContextID   =   490
   LinkTopic       =   "Form5"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5055
   ScaleWidth      =   6390
   Begin VB.TextBox Text7 
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
      Left            =   5400
      TabIndex        =   19
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   3360
      TabIndex        =   18
      Top             =   4200
      WhatsThisHelpID =   320
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Next >>"
      Height          =   495
      Left            =   2400
      TabIndex        =   17
      Top             =   4200
      WhatsThisHelpID =   320
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<< Prev"
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   4200
      WhatsThisHelpID =   320
      Width           =   855
   End
   Begin VB.TextBox Text6 
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
      Left            =   2160
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   1080
      TabIndex        =   14
      Top             =   4200
      WhatsThisHelpID =   320
      Width           =   1335
   End
   Begin VB.TextBox Text5 
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
      Left            =   3960
      TabIndex        =   13
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox text4 
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
      Left            =   1440
      TabIndex        =   10
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   4200
      WhatsThisHelpID =   320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   4800
      TabIndex        =   7
      Top             =   1080
      WhatsThisHelpID =   320
      Width           =   975
   End
   Begin VB.TextBox Text2 
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
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
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
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label8 
      Caption         =   "Records"
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   1920
      WhatsThisHelpID =   320
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Association No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2400
      WhatsThisHelpID =   320
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Freq 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   3360
      WhatsThisHelpID =   320
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Freq 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   3360
      WhatsThisHelpID =   320
      Width           =   855
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   1920
      WhatsThisHelpID =   320
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   1440
      WhatsThisHelpID =   320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "First Name"
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
      Left            =   720
      TabIndex        =   5
      Top             =   960
      WhatsThisHelpID =   320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Competitor Search Form"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   320
      Width           =   4575
   End
End
Attribute VB_Name = "PilotAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MyTable As Recordset
Dim Selected As Recordset
Private Sub Command1_Click()
  Dim SQLString As String
  
  If Text1.Text = "" Then
     SQLString = "SELECT * FROM Pilot Where Pilot.Last_Name like '" & Text2.Text & "' ORDER BY Pilot.First_Name ASC;"
  ElseIf Text2.Text = "" Then
     SQLString = "SELECT * FROM Pilot Where Pilot.First_Name like '" & Text1.Text & "' ORDER BY Pilot.Last_Name ASC;"
  ElseIf (Text1.Text = "") And (Text2.Text = "") And (Text6.Text <> "") Then
     SQLString = "SELECT * FROM Pilot Where (Pilot.number = '" & Text6.Text & "' ORDER BY Pilot.Last_Name ASC;"
  ElseIf (Text1.Text = "") And (Text2.Text = "") And (Text3.Text <> "") Then
     SQLString = "SELECT * FROM Pilot Where (Pilot.Club = '" & Text3.Text & "' ORDER BY Pilot.Last_Name ASC;"
  Else
    SQLString = "SELECT * FROM Pilot Where (Pilot.First_Name = '" & Text1.Text & "' AND Pilot.Last_Name = '" & Text2.Text & "') ORDER BY Pilot.First_Name ASC;"
  End If
  Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  Text7.Text = Selected.RecordCount
  If Selected.RecordCount = 0 Then
    MsgBox ("Not in Database")
    PilotEntry.Text2.Text = Text1.Text
    PilotEntry.Text3.Text = Text2.Text
    PilotEntry.Show
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
  ElseIf Selected.RecordCount = 1 Then
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command2.Enabled = False
    Command5.Enabled = False
    Text4.Visible = True
    Text5.Visible = True
    Label6.Visible = True
    Label5.Visible = True
    Text1.Text = Selected!First_Name
    Text2.Text = Selected!Last_Name
    Text3.Text = Selected!Club
    Text4.Text = Selected!Freq1
    Text5.Text = Selected!Freq2
    Text6.Text = Selected!Number
    EditPilotID = Selected!Pilot_ID
  ElseIf Selected.RecordCount > 1 Then
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command2.Enabled = True
    Command5.Enabled = True
    Text4.Visible = True
    Text5.Visible = True
    Label6.Visible = True
    Label5.Visible = True
    Selected.MoveFirst
    Text1.Text = Selected!First_Name
    Text2.Text = Selected!Last_Name
    Text3.Text = Selected!Club
    Text4.Text = Selected!Freq1
    Text5.Text = Selected!Freq2
    Text6.Text = Selected!Number
    EditPilotID = Selected!Pilot_ID
  End If
  

End Sub

Private Sub Command2_Click()
  Selected.MovePrevious
  If Selected.BOF = True Then
    MsgBox ("This is the First record")
    Selected.MoveFirst
  End If
  Text1.Text = Selected!First_Name
  Text2.Text = Selected!Last_Name
  Text3.Text = Selected!Club
  Text4.Text = Selected!Freq1
  Text5.Text = Selected!Freq2
  Text6.Text = Selected!Number
  EditPilotID = Selected!Pilot_ID
  
End Sub
Private Sub Command3_Click()
  Unload Me
End Sub

Private Sub Command4_Click()
  FromChange = True
  
  PilotEntry.Show
  
End Sub

Private Sub Command5_Click()
  Selected.MoveNext
  If Selected.EOF = True Then
    MsgBox ("This is the Last Record")
    Selected.MoveLast
  End If
  Text1.Text = Selected!First_Name
  Text2.Text = Selected!Last_Name
  Text3.Text = Selected!Club
  Text4.Text = Selected!Freq1
  Text5.Text = Selected!Freq2
  Text6.Text = Selected!Number
  EditPilotID = Selected!Pilot_ID

End Sub

Private Sub Command6_Click()
  Text1.Text = ""
  Text2.Text = ""
  Text3.Text = ""
  Text4.Text = ""
  Text5.Text = ""
  Text6.Text = ""
  Text7.Text = ""
  Text4.Visible = False
  Text5.Visible = False
  Label6.Visible = False
  Label5.Visible = False
  Command4.Enabled = False
  Command2.Enabled = False
  Command5.Enabled = False
  EditPilotID = ""
End Sub

Private Sub Form_Load()
  
  Command1.Caption = "Search"
  Command6.Caption = "Clear"
  Command3.Caption = "DONE"
  Command4.Caption = "Change"
  Text4.Visible = False
  Text5.Visible = False
  Label6.Visible = False
  Label5.Visible = False
  Text6.Enabled = True
  Command3.Enabled = True
  Command4.Enabled = False
  Command2.Enabled = False
  Command5.Enabled = False
  FromEdit = False
End Sub




