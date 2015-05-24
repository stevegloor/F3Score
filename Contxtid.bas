Attribute VB_Name = "ContextIDs"
Option Explicit
'=====================================================================
'                  Copyright 1993-1996 by Teletech Systems, Inc. All rights reserved
'
'
'This source code may not be distributed in part or as a whole without
'express written permission from Teletech Systems.
'=====================================================================
'
'This source code contains the following routines:
'  o SetAppHelp() 'Called in the main Form_Load event to register your
'                 'program with WINHELP.EXE
'  o QuitHelp()    'Deregisters your program with WINHELP.EXE. Should
'                  'be called in your main Form_Unload event
'  o ShowHelpTopic(Topicnum) 'Brings up context sensitive help based on
'                  'any of the following CONTEXT IDs
'  o ShowContents  'Displays the startup topic
'********** Shameless Plug <g> **********
'The Standard and Professional editions of HelpWriter
' also include the following routines to add sizzle to your
' helpfile presentation...
'  o HelpWindowSize(x,y,dx,dy) ' Position help window in a screen
'                              ' independent manner
'  o SearchHelp()  'Brings up the windows help KEYWORD SEARCH dialog box
'***********************************************************************
'
'=====================================================================
'List of Context IDs for <F3JSCORE>
'=====================================================================
Global Const Hlp_Disclaimer = 10    'Main Help Window
Global Const Hlp_File_ = 210    'Main Help Window
Global Const Hlp_Competition_ = 220    'Main Help Window
Global Const Hlp_Scores_ = 240    'Main Help Window
Global Const Hlp_DataBase_ = 250    'Main Help Window
Global Const Hlp_Tools_ = 260    'Main Help Window
Global Const Hlp_Help_ = 270    'Main Help Window
Global Const Hlp_ContestEntry_ = 280    'Main Help Window
Global Const Hlp_Input_Window = 290    'Main Help Window
Global Const Hlp_Contest_Select = 300    'Main Help Window
Global Const Hlp_EnterF3JScores_ = 310    'Main Help Window
Global Const Hlp_Frequency_Change = 330    'Main Help Window
Global Const Hlp_Report_Selection = 340    'Main Help Window
Global Const Hlp_Slot_Allocation = 350    'Main Help Window
Global Const Hlp_Team_Select = 360    'Main Help Window
Global Const Hlp_Frequencies_ = 380    'Main Help Window
Global Const Hlp_DataMaint_ = 390    'Main Help Window
Global Const Hlp_RoundxSlot_Allocation = 400    'Main Help Window
Global Const Hlp_Move_Pilot = 410    'Main Help Window
Global Const Hlp_Main_Screen = 470    'Main Help Window
Global Const Hlp_Select_Pilots = 500    'Main Help Window
Global Const Hlp_Current_Contest = 520    'Main Help Window
Global Const Hlp_Entering_a = 540    'Main Help Window
Global Const Hlp_View = 560    'Main Help Window
Global Const Hlp_EnterF3BScore = 570    'Main Help Window
Global Const Hlp_EditINIFile = 580    'Main Help Window
Global Const Hlp_Entering_a1 = 590    'Main Help Window
Global Const Hlp_Selecting_Teams = 620    'Main Help Window
Global Const Hlp_Drawing_the = 630    'Main Help Window
Global Const Hlp_F3B = 650    'Main Help Window
'=====================================================================
'
'
'  Help engine section.

' Commands to pass WinHelp()
Global Const HELP_CONTEXT = &H1 '  Display topic in ulTopic
Global Const HELP_QUIT = &H2    '  Terminate help
Global Const HELP_FINDER = &HB  '  Display Contents tab
Global Const HELP_INDEX = &H3   '  Display index
Global Const HELP_HELPONHELP = &H4      '  Display help on using help
Global Const HELP_SETINDEX = &H5        '  Set the current Index for multi index help
Global Const HELP_KEY = &H101           '  Display topic for keyword in offabData
Global Const HELP_MULTIKEY = &H201
Global Const HELP_CONTENTS = &H3     ' Display Help for a particular topic
Global Const HELP_SETCONTENTS = &H5  ' Display Help contents topic
Global Const HELP_CONTEXTPOPUP = &H8 ' Display Help topic in popup window
Global Const HELP_FORCEFILE = &H9    ' Ensure correct Help file is displayed
Global Const HELP_COMMAND = &H102    ' Execute Help macro
Global Const HELP_PARTIALKEY = &H105 ' Display topic found in keyword list
Global Const HELP_SETWINPOS = &H203  ' Display and position Help window


Type HELPWININFO
wStructSize As Long
X As Long
Y As Long
dX As Long
dY As Long
wMax As Long
rgChMember As String * 2
End Type
    Declare Function WinHelp Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
    Declare Function WinHelpByInfo Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, dwData As HELPWININFO) As Long
    Declare Function WinHelpByStr Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData$) As Long
    Declare Function WinHelpByNum Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData&) As Long
    Dim m_hWndMainWindow As Long ' hWnd to tell WINHELP the helpfile owner


Dim MainWindowInfo As HELPWININFO
Sub SetAppHelp(ByVal hWndMainWindow)
'=====================================================================
'To use these subroutines to access WINHELP, you need to add
'at least this one subroutine call to your code
'     o  In the Form_Load event of your main Form enter:
'        Call SetAppHelp(Me.hWnd) 'To setup helpfile variables
'         (If you are not interested in keyword searching or context
'         sensitive help, this is the only call you need to make!)
'=====================================================================
    m_hWndMainWindow = hWndMainWindow
    If Right$(Trim$(App.Path), 1) = "\" Then
        App.HelpFile = App.Path + "F3JSCORE.HLP"
    Else
        App.HelpFile = App.Path + "\F3JSCORE.HLP"
    End If
    MainWindowInfo.wStructSize = 26
    MainWindowInfo.X = 256
    MainWindowInfo.Y = 256
    MainWindowInfo.dX = 512
    MainWindowInfo.dY = 512
    MainWindowInfo.rgChMember = Chr$(0) + Chr$(0)
End Sub
Sub QuitHelp()
    Dim Result As Variant
    Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_QUIT, Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0))
End Sub
Sub ShowHelpTopic(ByVal ContextID As Long)
'=====================================================================
'  FOR CONTEXT SENSITIVE HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic(<any Hlpxxx entry above>)
'=====================================================================
'  TO ADD FORM LEVEL CONTEXT SENSITIVE HELP...
'=====================================================================
'     o  For FORM level context sensetive help, you should set each
'        Me.HelpContext=<any Hlp_xxx entry above>
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTEXT, CLng(ContextID))

End Sub
Sub ShowHelpTopic2(ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 2 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic2(<any Hlpxxx entry above>)
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile & ">HlpWnd02", HELP_CONTEXT, CLng(ContextID))

End Sub
Sub ShowHelpTopic3(ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 3 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic3(<any Hlpxxx entry above>)
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile & ">HlpWnd03", HELP_CONTEXT, CLng(ContextID))

End Sub
Sub ShowGlossary()
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTEXT, CLng(64000))

End Sub
Sub ShowPopupHelp(ByVal ContextID As Long)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTEXTPOPUP, CLng(ContextID))

End Sub
Sub DoHelpMacro(ByVal MacroString As String)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result As Variant

    Result = WinHelpByStr(m_hWndMainWindow, App.HelpFile, HELP_COMMAND, ByVal (MacroString))

End Sub
Sub ShowHelpContents()
'=====================================================================
'  DISPLAY HELP STARTUP TOPIC IN RESPONSE TO A COMMAND BUTTON or MENU ...
'=====================================================================
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTENTS, CLng(0))

End Sub
Sub ShowContentsTab()
'=====================================================================
'  DISPLAY Contents tab (*.CNT)
'=====================================================================
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_FINDER, CLng(0))

End Sub
Sub ShowHelpOnHelp()
'=====================================================================
'  DISPLAY HELP for WINHELP.EXE  ...
'=====================================================================
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_HELPONHELP, CLng(0))

End Sub
