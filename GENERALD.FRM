VERSION 5.00
Begin VB.Form GeneralDisplay 
   Caption         =   "#"
   ClientHeight    =   5385
   ClientLeft      =   210
   ClientTop       =   510
   ClientWidth     =   6720
   HelpContextID   =   440
   LinkTopic       =   "Form4"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5385
   ScaleWidth      =   6720
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   4920
      Width           =   2655
   End
End
Attribute VB_Name = "GeneralDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  
  Unload Me
End Sub


Private Sub Form_Resize()
  List1.Top = 120
  List1.Left = 120
  List1.Height = ScaleHeight - 1000
  List1.Width = ScaleWidth - 240
End Sub


