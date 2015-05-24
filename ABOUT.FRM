VERSION 5.00
Begin VB.Form HelpAbout 
   BackColor       =   &H00FFFFFF&
   Caption         =   "F3J Scoring Program"
   ClientHeight    =   4485
   ClientLeft      =   510
   ClientTop       =   750
   ClientWidth     =   5340
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FF0000&
   HelpContextID   =   430
   LinkTopic       =   "Form4"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4485
   ScaleWidth      =   5340
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "EXIT"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   3960
      WhatsThisHelpID =   430
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   5055
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Version 4.181"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Written By Stephen Gloor"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "F3 Score 32"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Image Image1 
      Height          =   1875
      Left            =   1200
      Picture         =   "ABOUT.frx":0000
      Top             =   120
      Width           =   2625
   End
End
Attribute VB_Name = "HelpAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload Me
End Sub


