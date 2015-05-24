VERSION 5.00
Begin VB.Form PilotEntry 
   BackColor       =   &H00C0C000&
   Caption         =   "Pilot Details Entry Form"
   ClientHeight    =   5640
   ClientLeft      =   405
   ClientTop       =   795
   ClientWidth     =   7500
   HelpContextID   =   290
   LinkTopic       =   "Form3"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5640
   ScaleWidth      =   7500
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      Caption         =   "Country - Optional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   31
      Top             =   2520
      Width           =   5535
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Use Pilots Country's Frequncies"
         Height          =   615
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Abbreviation"
         Height          =   255
         Left            =   2280
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Country Name"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Additional Details - Optional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   360
      TabIndex        =   26
      Top             =   3600
      Width           =   6855
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Top             =   360
         Width           =   4455
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2160
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C000&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   30
         Top             =   360
         WhatsThisHelpID =   290
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C000&
         Caption         =   "Club"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   1080
         WhatsThisHelpID =   290
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "Suburb/Postcode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   720
         WhatsThisHelpID =   290
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C000&
         Caption         =   "Association Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         WhatsThisHelpID =   290
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Pilot Details - Mandatory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   360
      TabIndex        =   17
      Top             =   120
      Width           =   5535
      Begin VB.TextBox Text6 
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
         Index           =   0
         Left            =   600
         TabIndex        =   2
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text6 
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
         Index           =   1
         Left            =   3000
         TabIndex        =   4
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text9 
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
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text9 
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
         Index           =   1
         Left            =   3960
         TabIndex        =   5
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text3 
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
         Left            =   2760
         TabIndex        =   1
         Top             =   660
         Width           =   2655
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
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   660
         Width           =   2655
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "#"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2040
         Width           =   5295
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Frequency 1"
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
         Left            =   840
         TabIndex        =   25
         Top             =   1080
         WhatsThisHelpID =   290
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Frequency 2"
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
         Left            =   3240
         TabIndex        =   24
         Top             =   1080
         WhatsThisHelpID =   290
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Mhz"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Chan"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Mhz"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Chan"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   20
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "First Name"
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
         TabIndex        =   19
         Top             =   360
         WhatsThisHelpID =   290
         Width           =   2535
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Last Name"
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
         Left            =   2760
         TabIndex        =   18
         Top             =   360
         WhatsThisHelpID =   290
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      Top             =   1320
      WhatsThisHelpID =   290
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   6000
      TabIndex        =   16
      Top             =   3120
      WhatsThisHelpID =   290
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   360
      WhatsThisHelpID =   290
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   840
      WhatsThisHelpID =   290
      Width           =   1215
   End
End
Attribute VB_Name = "PilotEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MyTable As Recordset
Public FreqDb As Database
Public FreqSet As Recordset
Public Countryset As Recordset

Private Sub Combo1_Click()
 Dim search As String
  
  If Combo1.Text <> "" Then
    search = "Country = '" & Combo1.Text & "'"
    Countryset.MoveFirst
    Countryset.FindFirst search
    Text10.Text = Countryset!CountryABV
  End If
End Sub


Private Sub Combo1_DropDown()
Combo1.Clear
  
    Countryset.MoveFirst
    Do Until Countryset.EOF
      Combo1.AddItem Countryset!Country
      Countryset.MoveNext
    Loop
  
End Sub


Private Sub Command1_Click()
Dim search As String
  
  MyTable.Edit
  If Text2.Text <> "" Then
    MyTable![First_Name] = Text2.Text
  End If
  If Text3.Text <> "" Then
    MyTable![Last_Name] = Text3.Text
  End If
  If Text4.Text <> "" Then
    MyTable![Address1] = Text4.Text
  End If
  If Text5.Text <> "" Then
    MyTable![Address2] = Text5.Text
  End If
  MyTable![Freq1] = Val(Text6(0).Text)
  MyTable![Freq2] = Val(Text6(1).Text)
  MyTable!channel1 = Val(Text9(0).Text)
  MyTable!channel2 = Val(Text9(1).Text)
  If Text8.Text <> "" Then
    MyTable![Club] = Text8.Text
  End If
  MyTable![Number] = Text1.Text
  If Text10.Text <> "" Then
    Countryset.MoveFirst
    search = "CountryABV = '" & Text10.Text & "'"
    Countryset.FindFirst search
    If Countryset.NoMatch = True Then
      MsgBox ("This is not a valid Country Abbreviation")
    Else
      MyTable!Res1 = Text10.Text
    End If
  End If
  MyTable.Update
  PilotSelect.UpdateBox1
  Unload Me
End Sub

Private Sub Command2_Click()
  On Error GoTo errhandler
  Dim search As String
  
  
  If Text2.Text = "" Then
    MsgBox ("There must be an entry in the field First Name")
  ElseIf Text3.Text = "" Then
    MsgBox ("There must be an entry in the field Last Name")
  ElseIf Text6(0).Text = "" Then
    MsgBox ("There must be an entry in the field Frequency 1")
  Else
    MyTable.AddNew
    MyTable![First_Name] = Text2.Text
    MyTable![Last_Name] = Text3.Text
    If Text1.Text <> "" Then
      MyTable![Number] = Text1.Text
    End If
    If Text4.Text <> "" Then
      MyTable![Address1] = Text4.Text
    End If
    If Text5.Text <> "" Then
      MyTable![Address2] = Text5.Text
    End If
    MyTable![Freq1] = Val(Text6(0).Text)
    If Text6(1).Text <> "" Then
      MyTable!Freq2 = Val(Text6(1).Text)
    End If
    If Text9(0).Text <> "" Then
      MyTable!channel1 = Text9(0).Text
    End If
    If Text9(1).Text <> "" Then
      MyTable!channel2 = Text9(1).Text
    End If
    
    If Text8.Text <> "" Then
      MyTable![Club] = Text8.Text
    End If
    If Text10.Text <> "" Then
      Countryset.MoveFirst
      search = "CountryABV = '" & Text10.Text & "'"
      Countryset.FindFirst search
      If Countryset.NoMatch = True Then
        MsgBox ("This is not a valid Country Abbreviation")
      Else
        MyTable!Res1 = Text10.Text
      End If
    End If
    MyTable.Update
    ClearFields
    MsgBox ("Pilot entered in Database")
    DataMaint.UpdateFrame1
  End If
  Exit Sub
errhandler:
  MsgBox ("There was an error with one of your data fields")
  Unload Me
End Sub

Private Sub Command3_Click()
  FreqDisp.Show
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
  
  PilotSelect.UpdateBox1
  Unload Me
  
End Sub

Private Sub Form_Load()
Dim datafile As String
Dim SQLString As String
Dim search As String
  
  Lookup = True
  datafile = App.Path & "\FREQ.MDB"
  Set FreqDb = DBEngine.Workspaces(0).OpenDatabase(datafile)
  SQLString = "SELECT * FROM Countries ORDER BY Countries.Country;"
  Set Countryset = FreqDb.OpenRecordset(SQLString, dbOpenDynaset)
  search = "CountryID = " & Str(Country)
  Countryset.FindFirst search
  If Countryset!Freqs = True Then
    SQLString = "SELECT DISTINCTROW Countries.Freqs, Frequency.Frequency, Frequency.Channel, Countries.CountryID FROM Countries INNER JOIN Frequency ON Countries.CountryID = Frequency.CountryID Where ((Countries.CountryID =" & Str(Country) & ")) ORDER BY Frequency.Frequency DESC;"
    Set FreqSet = FreqDb.OpenRecordset(SQLString, dbOpenDynaset)
    Label15.Caption = "Press ENTER after typing the Frequency/Channel to validate entry"
  Else
    Text9(0).Enabled = False
    Text9(1).Enabled = False
    Label13.Enabled = False
    Label11.Enabled = False
    Label15.Caption = "Frequency Look-up feature not available for current Country Setting"
    Lookup = False
  End If
  
  Set MyTable = F3JDb.OpenRecordset("Pilot", dbOpenTable)
  MyTable.Index = "Pilot_ID"
  If FromChange = True Then
    Command1.Caption = "Update"
    Command5.Caption = "Cancel"
    Command3.Caption = "Lookup Freqs"
    Command2.Visible = False
    MyTable.Seek "=", EditPilotID
    If MyTable!First_Name <> "" Then
      Text2.Text = MyTable!First_Name
    End If
    If MyTable!Last_Name <> "" Then
      Text3.Text = MyTable!Last_Name
    End If
    If MyTable!Address1 <> "" Then
      Text4.Text = MyTable!Address1
    End If
    If MyTable!Address2 <> "" Then
      Text5.Text = MyTable!Address2
    End If
    Text6(0).Text = Format(MyTable!Freq1, "#00.000")
    If MyTable!Freq2 <> 0 Then
      Text6(1).Text = Format(MyTable!Freq2, "#00.000")
    End If
    If MyTable!channel1 <> 0 Then
      Text9(0).Text = MyTable!channel1
    End If
    If MyTable!channel2 <> 0 Then
      Text9(1).Text = MyTable!channel2
    End If
    If MyTable!Club <> "" Then
      Text8.Text = MyTable!Club
    End If
    Text1.Text = MyTable!Number
    If MyTable!Res1 <> "" Then
      Text10.Text = MyTable!Res1
      search = "CountryABV = '" & MyTable!Res1 & "'"
      Countryset.FindFirst search
      Combo1.Text = Countryset!Country
    Else
      Combo1.Text = ""
    End If
    FromChange = False
   Else
    If DisplayCountry = True Then
      search = "CountryID = " & Str(Country)
      Countryset.FindFirst search
      Combo1.Text = Countryset!Country
      Text10.Text = Countryset!CountryABV
    End If
    Command1.Visible = False
    Command3.Visible = True
    Command2.Caption = "ADD"
    Command5.Caption = "Done"
    Command3.Caption = "Freq Lookup"
   End If
   
  

End Sub



Public Sub ClearFields()
  Text1.Text = ""
  Text2.Text = ""
  Text3.Text = ""
  Text4.Text = ""
  Text5.Text = ""
  Text6(0).Text = ""
  Text6(1).Text = ""
  Text9(0).Text = ""
  Text9(1).Text = ""
  Text8.Text = ""
 
End Sub

Private Sub Option1_Click()
Dim search As String
Dim SQLString As String

  
  Countryset.MoveFirst
  search = "CountryABV = '" & Text10.Text & "'"
  Countryset.FindFirst search
  If Countryset!Freqs = True Then
    SQLString = "SELECT DISTINCTROW Countries.Freqs, Frequency.Frequency, Frequency.Channel, Countries.CountryID FROM Countries INNER JOIN Frequency ON Countries.CountryID = Frequency.CountryID Where ((Countries.CountryID =" & Str(Countryset!CountryID) & ")) ORDER BY Frequency.Frequency DESC;"
    Set FreqSet = FreqDb.OpenRecordset(SQLString, dbOpenDynaset)
    MsgBox ("Using Frequencies from " & Countryset!Country)
    Label15.Caption = "Press ENTER after typing the Frequency/Channel to validate entry"
  Else
    MsgBox ("This country does not have frequencies defined" & Chr$(13) & "There will be no Lookup")
    Text9(0).Enabled = False
    Text9(1).Enabled = False
    Label13.Enabled = False
    Label11.Enabled = False
    Label15.Caption = "Frequency Look-up feature not available for current Country Setting"
    Lookup = False
  End If

End Sub

Private Sub Option1_DblClick()
Dim SQLString As String

  SQLString = "SELECT DISTINCTROW Countries.Freqs, Frequency.Frequency, Frequency.Channel, Countries.CountryID FROM Countries INNER JOIN Frequency ON Countries.CountryID = Frequency.CountryID Where ((Countries.CountryID =" & Str(Country) & ")) ORDER BY Frequency.Frequency DESC;"
  Set FreqSet = FreqDb.OpenRecordset(SQLString, dbOpenDynaset)
  Option1.Value = False
  MsgBox ("Using Default Countries Frequencies")
  Text9(0).Enabled = True
  Text9(1).Enabled = True
  Label13.Enabled = True
  Label11.Enabled = True
  Label15.Caption = "Press ENTER after typing the Frequency/Channel to validate entry"

End Sub


Private Sub Text10_KeyPress(KeyAscii As Integer)
Dim search As String
  
  If KeyAscii = 13 Or KeyAscii = 9 Then
    Countryset.MoveFirst
    search = "CountryABV = '" & Text10.Text & "'"
    Countryset.FindFirst search
    If Countryset.NoMatch = True Then
      MsgBox ("This is not a valid Country Abbreviation")
    Else
      Text10.Text = Countryset!CountryABV
    
    End If
  End If
End Sub

Private Sub Text6_Change(Index As Integer)
  If Val(Text6(Index).Text) > 100 Then
    MsgBox ("You enter the frequency in Mhz")
    Text6(Index).Text = " "
  End If
  
End Sub



Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)
  'On Error GoTo err_hand
  Dim Freq As Single
  Dim search As String
  
  If Lookup = True Then
    If KeyAscii = 13 Or KeyAscii = 9 Then
      Freq = Val(Trim(Text6(Index)))
      Freq = Round(Freq, 3)
      search = "Frequency = " & Format(Freq, "#.000")
      FreqSet.FindFirst search
      If FreqSet.NoMatch = False Then
        Text9(Index).Text = FreqSet!Channel
        Text6(Index).Text = Format(FreqSet!Frequency, "#00.000")
      Else
        MsgBox ("This is not a valid Frequency")
      End If
    End If
    Countryset.MoveFirst
  End If
Exit Sub

err_hand:
  MsgBox ("You entered an Incorrect Value")
  Exit Sub
End Sub

Private Sub Text7_Change()

End Sub

Private Sub Text9_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Or KeyAscii = 9 Then
    FreqSet.FindFirst "Channel = " & Text9(Index)
    If FreqSet.NoMatch = False Then
      Text9(Index).Text = FreqSet!Channel
      Text6(Index).Text = Format(FreqSet!Frequency, "#00.000")
    Else
      MsgBox ("This is not a valid Channel")
    End If
  End If
  Countryset.MoveFirst
End Sub




