VERSION 5.00
Begin VB.Form FreqDisp 
   Caption         =   "Frequencies"
   ClientHeight    =   5535
   ClientLeft      =   2325
   ClientTop       =   1440
   ClientWidth     =   4200
   HelpContextID   =   450
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5535
   ScaleWidth      =   4200
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   2160
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print Frequency List"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   5040
      Width           =   1815
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
      Height          =   420
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "FREQDISP.frx":0000
      Left            =   240
      List            =   "FREQDISP.frx":0002
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   5040
      WhatsThisHelpID =   450
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      WhatsThisHelpID =   450
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Enter Frequency to be Searched - Then Hit  ENTER"
      Height          =   615
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Country"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      WhatsThisHelpID =   450
      Width           =   1575
   End
End
Attribute VB_Name = "FreqDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ThisCountry As Integer
Dim FreqDb As Database
Dim FreqSet As Recordset

Private Sub Combo1_Click()
  Dim Search As String
  
  FreqSet.MoveFirst
  Search = "Country = '" & Combo1.Text & "'"
  FreqSet.FindFirst Search
  ThisCountry = FreqSet!CountryID
  UpdateList1
End Sub

Public Sub Combo1_DropDown()
Dim SQLString As String

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
  FreqDb.Close
  Unload Me
End Sub

Private Sub Command2_Click()


End Sub

Private Sub Command3_Click()
  PrintList
End Sub

Public Sub Form_Load()

Dim datafile As String
Dim Search As String
Dim SQLString As String

On Error GoTo errhandler:
      
  
  datafile = App.Path & "\FREQ.MDB"
  Set FreqDb = DBEngine.Workspaces(0).OpenDatabase(datafile)
  SQLString = "SELECT * FROM Countries Where ((Countries.Freqs = Yes))ORDER BY Countries.Country;"
  Set FreqSet = FreqDb.OpenRecordset(SQLString, dbOpenDynaset)
  FreqSet.MoveFirst
  Search = "CountryID = " & Str(Country)
  FreqSet.FindFirst Search
  Combo1.Text = FreqSet!Country
  ThisCountry = Country
  UpdateList1
Exit Sub



errhandler:
  MsgBox ("Could not open Frequency Database")
  Unload Me
End Sub



Public Sub UpdateList1()
Dim Countryset As Recordset
Dim Line As String
Dim SQLString As String
  
  SQLString = "SELECT * FROM Frequency WHERE (Frequency.CountryID = " & ThisCountry & ") ORDER BY Frequency.Frequency;"
  Set Countryset = FreqDb.OpenRecordset(SQLString, dbOpenDynaset)
  
  List1.Clear
  If Countryset.RecordCount > 1 Then
    Countryset.MoveFirst
    List1.AddItem "Frequency   Channel"
    Do Until Countryset.EOF
      Line = "  " & Format(Str(Countryset!Frequency), "##.000") & "            " & Countryset!Channel
      List1.AddItem Line
      Countryset.MoveNext
    Loop
  Else
    Line = "No Records"
    List1.AddItem Line
  End If
End Sub

Public Sub Search()
Dim SearchStr As String
Dim SearchSet As Recordset
Dim Line As String
Dim SQLString As String
Dim Done As Boolean
  
  SQLString = "SELECT DISTINCTROW Frequency.Frequency, Frequency.CountryID From Frequency WHERE ((Frequency.Frequency= " & Text1.Text & "));"
  Set SearchSet = FreqDb.OpenRecordset(SQLString, dbOpenDynaset)
  If SearchSet.RecordCount = 0 Then
    List2.Clear
    List2.AddItem "No Listing for Frequency"
  Else
    List2.Clear
    SearchSet.MoveFirst
    Do Until SearchSet.EOF = True
      FreqSet.MoveFirst
      SearchStr = "CountryID = " & Trim(Str(SearchSet!CountryID))
      FreqSet.FindFirst SearchStr
      List2.AddItem FreqSet!Country
      SearchSet.MoveNext
    Loop
  End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
  If Text1.Text <> "" Then
    Search
  Else
    List2.Clear
    List2.AddItem "Nothing Entered"
  End If
End If
End Sub





Public Sub PrintList()
Dim SQLString1 As String
Dim Tab1 As Integer
Dim tabset(10) As Integer
Dim NumPilots As Integer
Dim Linecount As Integer
Dim Col As Integer
Dim P As Integer
Dim Count As Integer
Dim Countryset As Recordset
Dim Pagelength As Integer

Pagelength = 75


  tabset(0) = 1
  tabset(1) = 15
  tabset(2) = 35
  tabset(3) = 45
  tabset(4) = 55
  tabset(5) = 65
  tabset(6) = 75
  tabset(7) = 85
  tabset(8) = 95
  tabset(9) = 105

  SQLString1 = "SELECT * FROM Frequency WHERE (Frequency.CountryID = " & ThisCountry & ") ORDER BY Frequency.Frequency;"
  Set Countryset = FreqDb.OpenRecordset(SQLString1, dbOpenDynaset)
  Printer.FontSize = 16
  Printer.CurrentX = Printer.ScaleWidth / 2 - TextWidth("List of Frequencies for       ")
  Printer.Print "List of Frequencies for "; Trim(Combo1.Text)
  Linecount = 4
  If Countryset.RecordCount > 1 Then
    Countryset.MoveFirst
    Printer.FontSize = 12
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.Print "Frequency"
    Printer.CurrentY = TextHeight("sample") * Linecount
    Printer.CurrentX = TextWidth("g") * tabset(2)
    Printer.Print "Channel"
    Linecount = Linecount + 2
    Printer.CurrentY = TextHeight("sample") * Linecount
    Do Until Countryset.EOF
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.CurrentX = TextWidth("g") * tabset(0)
      Printer.Print Format(Countryset!Frequency, "##.000")
      Printer.CurrentX = TextWidth("g") * tabset(2)
      Printer.CurrentY = TextHeight("sample") * Linecount
      Printer.Print Countryset!Channel
      If Linecount < Pagelength Then
        Linecount = Linecount + 1
      Else
        Printer.NewPage
        Linecount = 1
      End If
      Countryset.MoveNext
    Loop
  Else
    Printer.Print "No Records"
  End If
  Printer.EndDoc
End Sub
