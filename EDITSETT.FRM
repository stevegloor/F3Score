VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{625E24A3-B09D-101D-85F5-6EBA1EE93AF4}#3.3#0"; "IniCon32.OCX"
Begin VB.Form EditSettings 
   Caption         =   "Change INI settings"
   ClientHeight    =   4080
   ClientLeft      =   1125
   ClientTop       =   1710
   ClientWidth     =   5295
   HelpContextID   =   580
   LinkTopic       =   "Form4"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4080
   ScaleWidth      =   5295
   Begin IniconLib.Init Init1 
      Left            =   120
      Top             =   3360
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change INI settings"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   2175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5741
      _Version        =   393216
      Tabs            =   6
      Tab             =   1
      TabHeight       =   529
      TabCaption(0)   =   "Country "
      TabPicture(0)   =   "EDITSETT.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(3)=   "Combo1"
      Tab(0).Control(4)=   "Text1(0)"
      Tab(0).Control(5)=   "Text1(1)"
      Tab(0).Control(6)=   "Check1"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Flight Maxes"
      TabPicture(1)   =   "EDITSETT.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Championship"
      TabPicture(2)   =   "EDITSETT.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Option1(0)"
      Tab(2).Control(1)=   "Option1(1)"
      Tab(2).Control(2)=   "Frame2"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Club Setup"
      TabPicture(3)   =   "EDITSETT.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label7"
      Tab(3).Control(1)=   "Label9"
      Tab(3).Control(2)=   "Label10"
      Tab(3).Control(3)=   "Label11"
      Tab(3).Control(4)=   "Text5"
      Tab(3).Control(5)=   "Text6"
      Tab(3).Control(6)=   "Text7"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "General"
      TabPicture(4)   =   "EDITSETT.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label12"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Check2"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Text8"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "File Options"
      TabPicture(5)   =   "EDITSETT.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame4"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame3"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   "Remote Settings"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   42
         Top             =   2040
         Width           =   4455
         Begin VB.TextBox Text9 
            Height          =   375
            Left            =   1800
            TabIndex        =   43
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label14 
            Caption         =   "eg:\\myserver\apps\data1.mdb"
            Height          =   255
            Left            =   360
            TabIndex        =   45
            Top             =   720
            Width           =   3975
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Network Filename and address"
            Height          =   375
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Database Location"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   39
         Top             =   720
         Width           =   4455
         Begin VB.OptionButton Option2 
            Caption         =   "Remote on the Network"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   41
            Top             =   720
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Local to the PC"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   40
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "Defaults to the path that the program is run from and a file named data.mdb "
            Height          =   615
            Left            =   2160
            TabIndex        =   46
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -74640
         TabIndex        =   37
         Top             =   1560
         Width           =   255
      End
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
         Left            =   -72480
         TabIndex        =   34
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "   10 Khz Frequency Splits Allowed"
         Height          =   255
         Left            =   -74640
         TabIndex        =   33
         Top             =   1080
         Width           =   3855
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
         Left            =   -72480
         TabIndex        =   31
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text5 
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
         Left            =   -72480
         TabIndex        =   29
         Top             =   1620
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Caption         =   "Set Place Scores"
         Height          =   1095
         Left            =   -74760
         TabIndex        =   22
         Top             =   1980
         Width           =   4215
         Begin VB.CommandButton Command3 
            Caption         =   "Change"
            Height          =   375
            Left            =   2760
            TabIndex        =   36
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox Text4 
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
            Left            =   1920
            TabIndex        =   26
            Top             =   600
            Width           =   735
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   375
            Left            =   1440
            TabIndex        =   24
            Top             =   600
            Width           =   255
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
            Height          =   375
            Left            =   720
            TabIndex        =   23
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Points"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   27
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Place"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   25
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Championship standing decided by points awarded for placings (20 places maximum)"
         Height          =   375
         Index           =   1
         Left            =   -74640
         TabIndex        =   21
         Top             =   1380
         Width           =   3975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Championship standing decided by cummlative scores"
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   20
         Top             =   900
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Display Default Country in Pilot Entry Form"
         Height          =   255
         Left            =   -74760
         TabIndex        =   19
         Top             =   2760
         Width           =   4095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Max Time in Minutes"
         Height          =   1695
         Left            =   240
         TabIndex        =   10
         Top             =   900
         Width           =   4095
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
            Index           =   0
            Left            =   1320
            TabIndex        =   14
            Top             =   480
            Width           =   495
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
            Index           =   1
            Left            =   1320
            TabIndex        =   13
            Top             =   1080
            Width           =   495
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
            Index           =   2
            Left            =   3120
            TabIndex        =   12
            Top             =   480
            Width           =   495
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
            Index           =   3
            Left            =   3120
            TabIndex        =   11
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "F3J"
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
            Index           =   0
            Left            =   720
            TabIndex        =   18
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "F3J Flyoff"
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
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "F3B"
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
            Index           =   2
            Left            =   2520
            TabIndex        =   16
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "Aust. Open"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2040
            TabIndex        =   15
            Top             =   1200
            Width           =   975
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   -71760
         TabIndex        =   8
         Text            =   "Text4"
         Top             =   1500
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   -74760
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1500
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -73080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label12 
         Caption         =   $"EDITSETT.frx":00A8
         Height          =   1095
         Left            =   -74160
         TabIndex        =   38
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label11 
         Caption         =   "Flight Maximum"
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
         Left            =   -74280
         TabIndex        =   35
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of flights allowed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74400
         TabIndex        =   32
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Time to aim for"
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
         Left            =   -74160
         TabIndex        =   30
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "The usual format of a simple club contest is to have a number of flights to make a set time.  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74760
         TabIndex        =   28
         Top             =   780
         Width           =   4335
      End
      Begin VB.Label Label8 
         Caption         =   "Abbreviation"
         Height          =   255
         Left            =   -71760
         TabIndex        =   9
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Select Country form List for new INI setting"
         Height          =   495
         Left            =   -74760
         TabIndex        =   7
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Current country INI Setting "
         Height          =   255
         Left            =   -74520
         TabIndex        =   6
         Top             =   1260
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Type over current settings  to set new maximum times for the tasks in the contests"
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   2760
         Width           =   3495
      End
   End
End
Attribute VB_Name = "EditSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ThisCountry As String
Dim FreqDb As Database
Dim FreqSet As Recordset


Private Sub Combo1_DropDown()
  Combo1.Clear
  If FreqSet.RecordCount <> 0 Then
    FreqSet.MoveFirst
    Do Until FreqSet.EOF
      Combo1.AddItem FreqSet!Country
      FreqSet.MoveNext
    Loop
  End If
  
End Sub


Private Sub Command1_Click()
  UpdateINIFile
  
  
End Sub

Private Sub Command2_Click()
  FreqDb.Close
  Unload Me
End Sub

Private Sub Command3_Click()
  Place(Val(Text3.Text)) = Val(Text4.Text)
End Sub

Private Sub Form_Load()
Dim datafile As String
Dim SQLString As String
Dim Search As String
Dim PlaceNum As String
Dim I As Integer


On Error GoTo errhandler:
  ReDim Place(20)
  datafile = App.Path & "\FREQ.MDB"
  Set FreqDb = DBEngine.Workspaces(0).OpenDatabase(datafile)
  'SQLString = "SELECT * From Countries WHERE ((Countries.Freqs=Yes)) ORDER By Countries.Country;"
  SQLString = "SELECT * From Countries ORDER By Countries.Country;"
  Set FreqSet = FreqDb.OpenRecordset(SQLString, dbOpenDynaset)
  Init1.FileName = App.Path & "\f3score.ini"
  Init1.Application = "Program"
  Init1.Parameter = "RoundDiscarded"
  Text8.Text = Init1.Value
  DiscardRound = Init1.Value + 1
  Init1.Parameter = "TenK"
  Check2.Value = Init1.Value
  TenK = Init1.Value
  Init1.Parameter = "Country"
  Search = "CountryID = " & Init1.Value
  FreqSet.FindFirst Search
  Text1(0).Text = FreqSet!Country
  Text1(1).Text = FreqSet!CountryABV
  Init1.Parameter = "CountryDisplay"
  Check1.Value = Init1.Value
  DisplayCountry = Init1.Value
  Init1.Application = "Max"
  Init1.Parameter = "F3J"
  Text2(0).Text = Init1.Value
  Init1.Parameter = "F3JFO"
  Text2(1).Text = Init1.Value
  Init1.Parameter = "F3B"
  Text2(2).Text = Init1.Value
  Init1.Parameter = "AustOpen"
  Text2(3).Text = Init1.Value
  Init1.Parameter = "Club"
  Text7.Text = Init1.Value
  
  Init1.Application = "Club"
  Init1.Parameter = "Time"
  Text5.Text = Init1.Value
  ClubTime = Init1.Value
  Init1.Parameter = "Flights"
  Text6.Text = Init1.Value
  ClubFlights = Init1.Value
  Init1.Application = "Championship"
  Init1.Parameter = "Cummlative"
  If Init1.Value = 1 Then
    Option1(0).Value = True
    Frame2.Visible = False
    Cummlative = True
  Else
    Option1(1).Value = True
    Frame2.Visible = True
    Cummlative = False
  End If
  For I = 1 To 20
    PlaceNum = "Place" & Trim(Str(I))
    Init1.Parameter = PlaceNum
    Place(I) = Val(Init1.Value)
  Next I
  Init1.Application = "Database"
  Init1.Parameter = "Location"
  If Init1.Value = "Local" Then
    Option2(0).Value = True
    
  ElseIf Init1.Value = "Remote" Then
    Option2(1).Value = True
    Init1.Parameter = "Location"
    Text9.Text = Init1.Value
  End If
Exit Sub



errhandler:
  MsgBox ("Could not open Frequency Database")
  Unload Me
End Sub






Public Sub UpdateINIFile()
Dim Search As String

Init1.FileName = App.Path & "\f3score.ini"
  Init1.Application = "Program"
  Init1.Parameter = "RoundDiscarded"
  Init1.Value = Int(Text8.Text)
  DiscardRound = Init1.Value + 1
  If Combo1.Text <> "" Then
    Search = "Country = '" & Combo1.Text & "'"
    FreqSet.MoveFirst
    FreqSet.FindFirst Search
    Init1.Parameter = "Country"
    Init1.Value = FreqSet!CountryID
    Country = Val(Init1.Value)
  End If
  Init1.Parameter = "CountryDisplay"
  Init1.Value = Check1.Value
  DisplayCountry = Init1.Value
  Init1.Parameter = "TenK"
  Init1.Value = Check2.Value
  TenK = Init1.Value
  Init1.Application = "Max"
  Init1.Parameter = "F3J"
  Init1.Value = Text2(0).Text
  Init1.Parameter = "F3JFO"
  Init1.Value = Text2(1).Text
  Init1.Parameter = "F3B"
  Init1.Value = Text2(2).Text
  Init1.Parameter = "AustOpen"
  Init1.Value = Text2(3).Text
  Init1.Parameter = "Club"
  Init1.Value = Text7.Text
  Init1.Parameter = "F3J"
  maxF3J = Init1.Value
  Init1.Parameter = "F3JFO"
  maxF3JFO = Init1.Value
  Init1.Parameter = "F3B"
  maxF3B = Init1.Value
  Init1.Parameter = "8-11"
  max811 = Init1.Value
  Init1.Parameter = "Club"
  maxClub = Init1.Value
  Init1.Application = "Club"
  Init1.Parameter = "Time"
  Text5.Text = Init1.Value
  ClubTime = Init1.Value
  Init1.Parameter = "Flights"
  Text6.Text = Init1.Value
  ClubFlights = Init1.Value
  Init1.Application = "Championship"
  Init1.Parameter = "Cummlative"
  If Option1(0).Value = True Then
    Init1.Value = 1
    Cummlative = True
  Else
    Init1.Value = 0
    Cummlative = False
  End If
  For I = 1 To 20
    PlaceNum = "Place" & Trim(Str(I))
    Init1.Parameter = PlaceNum
    Init1.Value = Place(I)
  Next I
  Init1.Application = "Database"
  Init1.Parameter = "Location"
  If Option2(0).Value = True Then
    Init1.Value = "Local"
  ElseIf Option2(1).Value = True Then
    Init1.Value = "Remote"
    Init1.Parameter = "Location"
    Init1.Value = Text9.Text
  End If
  
  
  
  MsgBox ("INI file parameters set - The changes will take effect immediatly")
  Unload Me
End Sub

Private Sub Option1_Click(Index As Integer)
  If Option1(0).Value = True Then
    Frame2.Visible = False
  Else
    Frame2.Visible = True
  End If
End Sub

Private Sub VScroll1_Change()
  VScroll1.max = 20
  VScroll1.Min = 1
  Text3.Text = VScroll1.Value
  Text4.Text = Place(VScroll1.Value)
  
End Sub


