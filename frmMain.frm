VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Answering Machine - VB-TAPI"
   ClientHeight    =   2445
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4440
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSetup 
      Caption         =   "Se&tup"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   720
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1860
      Width           =   735
   End
   Begin VB.CommandButton cmdRepeat 
      Caption         =   "&Repeat"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1860
      Width           =   735
   End
   Begin VB.CommandButton cmdGreeting 
      Caption         =   "&Greeting"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1860
      Width           =   855
   End
   Begin VB.CommandButton cmdOnOff 
      Caption         =   "&On"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1860
      Width           =   735
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1860
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Message Information"
      Height          =   1455
      Left            =   960
      TabIndex        =   6
      ToolTipText     =   "Information for the last played or recorded message."
      Top             =   180
      Width           =   3375
      Begin VB.Label lblName 
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblNumber 
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label LblTime 
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Time"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Number"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label lblMsgCount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1140
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CleanUp As Long              'Resets the state after recording, see Timer1
Private m_TapiInit As Boolean       'Is TAPI initialized?
Private g_cMsgs As New Collection   'A collection of .wav files for msg playback
Private g_lMsgPlaying As Long       'Is a message currently playing?

Private Sub cmdGreeting_Click()
    dlgGreeting.Show vbModal
End Sub
'myInit can get called from dlgSetup and takes a flag fSetup so it knows that.
'If for setup, it initializes TAPI and attempts to negotiate a TAPI version of
'1.4, see SetupHack in the proc for more.  If not for setup it does all that
'and opens the line then sets up the app. to get all the status messages.  It
'only does any of this if TAPI isn't already initialized as signaled by the
'global flag m_TapiInit.  If TAPI is already initialized it shuts it down and
'allows the other functions of the application to be used.
Public Sub myInit(fSetup As Boolean)
Dim nError As Long

Dim lpExtensionID As lineextensionid
Dim lUnused As Long
Dim lLineID As Long
    lLineID = GetSetting("VB-TAPI", "Settings", "DeviceID", "0")
    
    If m_TapiInit = False Then
  
        cmdStop_Click   'No playing messages while running the answering machine
        DoEvents
        nError = lineInitialize(hTAPI, App.hInstance, _
                    AddressOf LineCallBack, 0, lNumLines)
        If nError <> 0 Then
            ProcessTAPIError nError
            err.Raise nError, "Init TAPI", "Can not initialize TAPI"
        Else
            DebugString 4, "lineInitialize -> Success"
        End If
                
'Since we don't have a setup program to find the initial line, try them all
'to negotiate a compatible TAPI version before bailing out.  This should only
'happen if the application has not been setup properly, which will be the case
'on the first launch of the app. or if there are no TAPI 1.4 lines available.
SetupHack:
        nError = lineNegotiateAPIVersion(hTAPI, lLineID, TAPIVERSION, _
            TAPIVERSION, lNegVer, lpExtensionID)
            
        If nError <> 0 Then
            If fSetup = True Then
                lLineID = lLineID + 1
                If lLineID <= lNumLines Then GoTo SetupHack
                'Uh Oh... Can't negotiate a TAPI verision on any line, bail.
            End If
            ProcessTAPIError nError
            err.Raise nError, "Init TAPI", "Can not negotiate TAPI version 1.4"
        Else
            If fSetup = True Then Exit Sub    'Just init for dlgSetup
            nError = lineOpen(hTAPI, lLineID, hLine, lNegVer, lUnused, lUnused, _
                LINECALLPRIVILEGE_OWNER, LINEMEDIAMODE_AUTOMATEDVOICE, 0)
            If nError <> 0 Then ProcessTAPIError nError
        End If
        
        lpLineDevCaps.dwTotalSize = Len(lpLineDevCaps)
        nError = lineGetDevCaps(hTAPI, lLineID, lNegVer, lUnused, lpLineDevCaps)
        If nError <> 0 Then ProcessTAPIError nError
        
        nError = lineSetStatusMessages(hLine, lpLineDevCaps.dwLineStates, 0)
        If nError <> 0 Then
            ProcessTAPIError nError
            err.Raise nError, "Init TAPI", "Can not setup for status messages"
        End If
        
        m_TapiInit = True
        
        DebugString 3, "m_TapiInit=" & CStr(m_TapiInit)
        
        'Disable all other command buttons
        cmdPlay.Enabled = False
        cmdRepeat.Enabled = False
        cmdStop.Enabled = False
        cmdGreeting.Enabled = False
        cmdSetup.Enabled = False
        
        cmdOnOff.Caption = "&Off"
        
    Else    'TAPI is already initialized, shut it down
        If hTAPI <> 0 Then
            nError = lineShutdown(hTAPI)
            If nError <> 0 Then ProcessTAPIError nError
            hTAPI = 0
            m_TapiInit = False
            
            'Enable all other command buttons
            cmdPlay.Enabled = True
            cmdRepeat.Enabled = True
            cmdStop.Enabled = True
            cmdGreeting.Enabled = True
            cmdSetup.Enabled = True
            
            cmdOnOff.Caption = "&On"
            
        End If
    End If

End Sub

'*Init or shutdown TAPI.  When TAPI is initialized all other functions are
'disabled.
Private Sub cmdOnOff_Click()

On Error GoTo EH
    Screen.MousePointer = vbHourglass
    myInit False
    Screen.MousePointer = vbNormal
    Exit Sub
EH:
Screen.MousePointer = vbNormal
MsgBox err.Number & " : " & err.Description
End Sub

'Calls clsFile to get a collection of .wav files and sets the
'"Message Information" frame text on the main application form.
Private Sub GetFiles()
On Error GoTo EH

Dim oFs As New clsFile

    Set g_cMsgs = oFs.GetMessages
    lblMsgCount.Caption = g_cMsgs.Count
    If g_cMsgs.Count > 0 Then SetCallerInfo (g_cMsgs.Item(g_cMsgs.Count))
    
    Exit Sub
EH:

MsgBox err.Number & ": " & err.Description

End Sub

'*Tri-State Play button.  Calls DoPlay and selects off of the caption, not very
'international...
Private Sub cmdPlay_Click()

lMediaID = -1   'The wave mapper, -1 or WAVE_MAPPER as a device id selects
                'the default wave device on the system that can handle the
                'wave format, usually, and hopefully the sound card.

DoPlay

End Sub

'Tri-State DoPlay function to go with Tri-State play button!  This switches
'from play -> pause -> paused then back to pause.  You can't get back to play
'from here, that is set in PlayMessages and StopPlaying and as the default for
'the button.  This function is called from the play function or when the
'remote access code is dialed in for remote message retrieval from Timer_1.
'No actual Wave file playing is done here, that is in PlayMessages, for for
'the overload on the Play word...
Public Sub DoPlay()

If cmdPlay.Caption = "&Pause" Then  'was playing, going to paused
    PausePlay
    cmdPlay.Caption = "&Paused"
    cmdRepeat.Enabled = False

    Exit Sub
End If


If cmdPlay.Caption = "&Paused" Then 'was paused, going to playing (pause caption)
    ResumePlay
    cmdPlay.Caption = "&Pause"
    cmdRepeat.Enabled = True
    Exit Sub
End If

'First time through, get a collection if messages to play and start playing!
Dim oFiles As New clsFile
Set g_cMsgs = oFiles.GetMessages

Set oFiles = Nothing

    If g_cMsgs.Count > 0 Then
        g_lMsgPlaying = 0
        cmdPlay.Caption = "&Pause"
        PlayMessages
    End If
    
End Sub

Private Sub cmdStop_Click()
    StopPlaying
End Sub

'This is called from either the cmdStop button (see line of code above) or
'from the disconnect event when playing messages remotely.
Public Sub StopPlaying()
Dim oFSO As New clsFile
Dim myColl As New Collection
On Error GoTo EH
    StopPlay
    g_lMsgPlaying = -1  'Try not to erase the messages if stop was clicked.
    Set g_cMsgs = Nothing
    Set myColl = oFSO.GetMessages
    lblMsgCount.Caption = CStr(myColl.Count)
    cmdPlay.Caption = "&Play"
    If fInTollSaver = False Then cmdRepeat.Enabled = True

    Set myColl = Nothing
    Set oFSO = Nothing
    
    Exit Sub
EH:

DebugString 0, err.Number & " : " & err.Description
End Sub

'Finally, PlayMessages loops through the collection and plays the Wave files.
'If it plays all the way through the last message it erases all of the saved
'messages.  Why does it erase them all?  Because that is the way my answering
'machine works, annoying isn't it?
Private Sub PlayMessages()
On Error GoTo EH
    g_lMsgPlaying = g_lMsgPlaying + 1
    If g_lMsgPlaying > g_cMsgs.Count Then   'Reached the end of the messages...
        'Kill 'em
        Dim oColl As New Collection
        ''Cut Here
        Dim oVar As Variant
        Dim oFiles As New clsFile
        Dim rc As Long
        Set oColl = oFiles.GetMessages
            For Each oVar In oColl
                Kill ".\Messages\" & CStr(oVar)
            Next
        ''Cut Here
        'Okay, if it really bothers you to have the messages blown away at
        'the bitter end, cut between the lines, add an 'Erase' button, paste
        'and be done with it.  You will need to add a :
        'Dim oColl As New Collection' to the new sub.
        Set oColl = oFiles.GetMessages
        lblMsgCount.Caption = oColl.Count
        Set oFiles = Nothing
        g_lMsgPlaying = 0
        Set g_cMsgs = Nothing
        cmdPlay.Caption = "&Play"
        Exit Sub
    End If
    
    If g_lMsgPlaying = 0 Then Exit Sub  '=0 when stop has been clicked.
    
    lblMsgCount.Caption = CStr(g_lMsgPlaying)
    SetCallerInfo CStr(g_cMsgs.Item(g_lMsgPlaying))
    
    If m_BPlayRec <> False Then
        If (fPlaying = False) Then
            LoadFile App.Path & "\Messages\" & g_cMsgs.Item(g_lMsgPlaying)
            Play lMediaID
            m_BPlayRec = True
            
        End If
    End If
    Exit Sub
EH:

DebugString 0, err.Number & " : " & err.Description
End Sub

'All the messages are saved with any caller info that is known and the
'time/date as the filename.  This info is put back on the main form when
'playing the message.  Why not use the time/date stamp on the file?  This
'naming convention almost guarantees uniqueness since we can't record files
'faster than 1 second resolution.  However it would probably be nicer to use
'the time/date off the file for display purposes.  Okay TODO:
Private Sub SetCallerInfo(sInfo As String)
Dim lPos As Long
On Error Resume Next

    lPos = InStr(1, sInfo, "~")
    lblName.Caption = Mid(sInfo, 1, lPos - 1)
    
    sInfo = Mid(sInfo, lPos + 1, Len(sInfo))
    lPos = InStr(1, sInfo, "~")
    lblNumber.Caption = Mid(sInfo, 1, lPos - 1)
    
    sInfo = Mid(sInfo, lPos + 1, Len(sInfo))
    LblTime.Caption = Mid(sInfo, 1, Len(sInfo) - 4) 'Trim off the .wav
    
End Sub

Private Sub cmdRepeat_Click()
    StopPlay
    g_lMsgPlaying = g_lMsgPlaying - 1
End Sub

Private Sub cmdSetup_Click()
    dlgSetup.Show vbModal
End Sub

'Even if we never display the dialogs, they are loaded by referencing whether
'they are visible or not (for dlgGreeting) or for getting properties (dlgSetup)
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Unload dlgGreeting
    Unload dlgSetup
End Sub

'*Main timer routine.  This runs the show, basically a small state machine.
Private Sub Timer1_Timer()
On Error GoTo EH
    
    'The lTollSaver stuff is used for remote message retrieval.  If the
    'correct access code is entered the playback function is called
    'If over a second goes by without a digit being entered it is reset
    'to give another chance.
    
    If sSecret = dlgSetup.txtSecret Then
        sSecret = ""
        CloseWaveOut
        fInTollSaver = True
        DoEvents
        DoPlay
        'Should disable digit detection here, or do something so you can't
        'get back to this state during the current call.
    End If
    
    lTollSaver = lTollSaver + 1
    
    If lTollSaver > 10 Then 'Over a second has elapsed since getting a digit.
        lTollSaver = 0      'Reset.  lTollSaver is set to zero in Globals.bas
        sSecret = ""        'whenever a digit is detected.
    End If
    
    If m_BPlayRec = False Then  'We are recording
    
        If CleanUp = 1 Then     'We are done recording, as set by our callback
                Dim nErr As Long 'proc. waveInProc in Wave.bas
                DebugString 5, "CleanUp in Timer"
                CloseWaveIn
                Timer1.Enabled = False  'Move??
                SaveToFileAsStream m_FileName
                m_BPlayRec = True   'Since we aren't recording anymore...
                
                'Don't do call stuff when recording a greeting.
                If dlgGreeting.Visible = False Then
                    lblMsgCount.Caption = lblMsgCount.Caption + 1
                    DropCall
                    hCall = 0
                    CleanUp = 0 'Should this be here?  TODO:
                End If
        End If
   Else
        If fPlaying = False Then    'As set by waveOutProc in Wave.bas
            DebugString 5, "fPlaying shutting down wave out device"
            CloseWaveOut
            Timer1.Enabled = False
            
            If hmem <> 0 Then
                GlobalFree (hmem)
                hmem = 0
            End If

            If m_TapiInit = True Then   'Do call stuff
                CleanUp = 0
                If fInTollSaver = False Then
                    m_BPlayRec = False
                    SendDigit
                Else
                    PlayMessages    'Playing messages remotely
                End If
            Else
                If dlgGreeting.Visible = False Then
                    PlayMessages    'Playing messages locally
                End If
            End If
            
        End If
    
   End If 'm_BPlayRec = False
   Exit Sub
EH:

DebugString 0, err.Number & " : " & err.Description
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    If Timer1.Enabled = True Then   'Don't exit if playing or recording
        Cancel = 1
    Else
        cmdStop_Click
        If m_TapiInit = True Then myInit False
    End If
    
End Sub
Private Sub Form_Load()
    CleanUp = 0
    m_BPlayRec = True
    
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)

    
    On Error Resume Next    'Directory probably exists, so an error is thrown
    MkDir "Messages"
    GetFiles
    
End Sub



