VERSION 5.00
Begin VB.Form EditChamp 
   Caption         =   "#"
   ClientHeight    =   2835
   ClientLeft      =   825
   ClientTop       =   1365
   ClientWidth     =   6885
   LinkTopic       =   "Form4"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2835
   ScaleWidth      =   6885
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "#"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "End Date"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Championship Name"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "EditChamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ChampSet As Recordset

Private Sub Command1_Click()
  EnterDetails
End Sub

Private Sub Command2_Click()
  ChampSet.Close
  Unload Me
End Sub

Private Sub Form_Load()
  
  If FromChange = True Then
    SQLString = "SELECT * FROM Championship WHERE ((Championship.ChampID = " & Str(CurrentChamp) & "));"
    Command1.Caption = "Change"
  Else
    SQLString = "SELECT * FROM Championship;"
    Command1.Caption = "Add"
  End If
  Set ChampSet = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If FromChange = True Then
    Text1.Text = ChampSet!ChampName
    Text2.Text = ChampSet!StartDate
    Text3.Text = ChampSet!EndDate
  End If
End Sub

Public Sub EnterDetails()
  If FromChange = True Then
    ChampSet.Edit
  Else
    ChampSet.AddNew
  End If
  ChampSet!ChampName = Text1.Text
  ChampSet!StartDate = Text2.Text
  ChampSet!EndDate = Text3.Text
  ChampSet.Update
  MsgBox ("Details Entered")
  UpdateBox
  Unload Me
End Sub

Public Sub UpdateBox()
Dim TeamSelect As Recordset
Dim SQLString As String
Dim Line As String
 ChampMaint.Combo1.Clear
 SQLString = "SELECT * FROM Championship;"
 Set TeamSelect = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
 If TeamSelect.RecordCount <> 0 Then
   TeamSelect.MoveFirst
 End If
 Do Until TeamSelect.EOF = True
   Line = Format(Str(TeamSelect!ChampID), "!@@@@") & " " & TeamSelect!ChampName
   ChampMaint.Combo1.AddItem Line
   TeamSelect.MoveNext
 Loop
TeamSelect.Close
End Sub
