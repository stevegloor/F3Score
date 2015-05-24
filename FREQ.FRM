VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Frequency Change Form"
   ClientHeight    =   4290
   ClientLeft      =   435
   ClientTop       =   885
   ClientWidth     =   6720
   HelpContextID   =   330
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4290
   ScaleWidth      =   6720
   Begin VB.CommandButton Command5 
      Caption         =   "Done"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   3720
      WhatsThisHelpID =   330
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Change Pilot to Alternate Frequency"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   3000
      WhatsThisHelpID =   330
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Next >>"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   3000
      WhatsThisHelpID =   330
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<< Prev"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   3000
      WhatsThisHelpID =   330
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Line Line5 
      X1              =   960
      X2              =   480
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line4 
      X1              =   480
      X2              =   480
      Y1              =   1320
      Y2              =   3480
   End
   Begin VB.Line Line3 
      X1              =   480
      X2              =   6000
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line2 
      X1              =   6000
      X2              =   6000
      Y1              =   1320
      Y2              =   3480
   End
   Begin VB.Line Line1 
      X1              =   3240
      X2              =   6000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label4 
      Caption         =   $"FREQ.frx":0000
      Height          =   975
      Left            =   2640
      TabIndex        =   11
      Top             =   1920
      WhatsThisHelpID =   330
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Alternate Frequency"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   2040
      WhatsThisHelpID =   330
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Users of Selected Frequency"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1200
      WhatsThisHelpID =   330
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Congested Frequency"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      WhatsThisHelpID =   330
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Selected As Recordset
Public Sub Command1_Click()
  
End Sub


Private Sub Command2_Click()
  Selected.MovePrevious
  If Selected.BOF = True Then
    MsgBox ("This is the first record")
    Selected.MoveFirst
  End If
  
  Text4.Text = Selected!Freq2
  Text2.Text = Selected!First_Name
  Text3.Text = Selected!Last_Name
End Sub

Private Sub Command3_Click()
  Selected.MoveNext
  If Selected.EOF = True Then
    MsgBox ("This is the last record")
    Selected.MoveLast
  End If
  
  Text4.Text = Selected!Freq2
  Text2.Text = Selected!First_Name
  Text3.Text = Selected!Last_Name
End Sub

Private Sub Command4_Click()
  Selected.Edit
  Selected!FreqUsed = Text4.Text
  If Selected!Freq2 = "" Then
    Selected!Freq2 = Text4.Text
  End If
  Selected.Update
  
End Sub

Private Sub Command5_Click()
  Freq.DoFreq
  Unload Me
End Sub

Private Sub Command6_Click()
    Text1.Text = ""
    Text4.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    
End Sub

Public Sub Form_Load()
Dim SQLString As String
  
  SQLString = "SELECT DISTINCTROW List.Pilot_ID, List.Comp_ID, List.FreqUsed, List.UsedAlt, Pilot.Freq1, Pilot.Freq2, Pilot.First_Name, Pilot.Last_Name FROM Pilot INNER JOIN List ON Pilot.Pilot_ID = List.Pilot_ID Where ((List.Comp_ID =" & Str(CurrentContest) & " AND List.FreqUsed = " & CongestFreq & ")) ORDER BY Pilot.Freq1 ASC;"
  Set Selected = F3JDb.OpenRecordset(SQLString, dbOpenDynaset)
  If Selected.RecordCount <> 0 Then
    Text1.Text = CongestFreq
    Text4.Text = Selected!Freq2
    Text2.Text = Selected!First_Name
    Text3.Text = Selected!Last_Name
  Else
    MsgBox ("No one is on this frequency")
  End If
End Sub


