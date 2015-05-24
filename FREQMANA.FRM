VERSION 5.00
Begin VB.Form Freq 
   Caption         =   "Frequencies"
   ClientHeight    =   5070
   ClientLeft      =   825
   ClientTop       =   1245
   ClientWidth     =   4740
   HelpContextID   =   380
   LinkTopic       =   "Form6"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5070
   ScaleWidth      =   4740
   Begin VB.CommandButton Command3 
      Caption         =   "Lookup Frequencies"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   2400
      WhatsThisHelpID =   380
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DONE"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   4560
      WhatsThisHelpID =   380
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find Pilots on Selected Frequency"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1800
      WhatsThisHelpID =   380
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
      Left            =   3480
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Usage"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   960
      WhatsThisHelpID =   380
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Number of Pilots in Contest"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1200
      WhatsThisHelpID =   380
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Freq."
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   960
      WhatsThisHelpID =   380
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Frequency Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   240
      WhatsThisHelpID =   380
      Width           =   3495
   End
End
Attribute VB_Name = "Freq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  CongestFreq = Left(List1.Text, 6)
  If CongestFreq <> "" Then
    Form2.Show
  Else
    MsgBox ("Nothing Selected")
  End If
End Sub


Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Command3_Click()
  FreqDisp.Show
End Sub

Private Sub Form_Load()
  DoFreq
End Sub



Public Sub DoFreq()
  Dim Selected As Recordset
  Dim NumFreq As Integer
  Dim SQLString As String
  Dim CheckFreq
  Dim Freq() As Double
  Dim I As Integer
  Dim Line As String
  Dim MaxFreqUse As Integer
  
  SQLString = "SELECT * from List Where List.Comp_ID =" & Str(CurrentContest) & " ORDER BY FreqUsed ASC;"
  Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If Selected.RecordCount <> 0 Then
  List1.Clear
  
    'Find the number of different frequencys
    Selected.MoveFirst
    CheckFreq = Selected!FreqUsed
    Selected.MoveNext
    NumFreq = 1
    Do Until Selected.EOF
      If Selected!FreqUsed <> CheckFreq Then
         NumFreq = NumFreq + 1
         CheckFreq = Selected!FreqUsed
      End If
      Selected.MoveNext
    Loop
    ReDim Freq(NumFreq, 2)
    MaxFreqUse = 1
    Selected.MoveFirst
    I = 1
    Freq(I, 1) = Selected!FreqUsed
    Do Until Selected.EOF
      If Selected!FreqUsed = Freq(I, 1) Then
         Freq(I, 2) = Freq(I, 2) + 1
         If Freq(I, 2) > MaxFreqUse Then
           MaxFreqUse = Freq(I, 2)
         End If
      Else
         I = I + 1
         Freq(I, 1) = Selected!FreqUsed
         Freq(I, 2) = Freq(I, 2) + 1
         If Freq(I, 2) > MaxFreqUse Then
           MaxFreqUse = Freq(I, 2)
         End If
      End If
      Selected.MoveNext
    Loop
    
    For I = 1 To NumFreq
      Line = Format(Trim(Str(Freq(I, 1))), "##.000") & "      " & Str(Freq(I, 2))
      List1.AddItem Line
      'Line = Str(Freq(I, 2))
      'List1.AddItem Line
    Next I
    Text1.Text = Str(Selected.RecordCount)
    MostUsedFreq = MaxFreqUse
  Else
    MsgBox ("This contest has no pilots entered")
  End If
End Sub

Private Sub List2_Click()

End Sub


