VERSION 5.00
Begin VB.Form dlgGreeting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Record A Greeting"
   ClientHeight    =   975
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2295
   Icon            =   "dlgGreeting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRecord 
      Height          =   375
      Left            =   1500
      Picture         =   "dlgGreeting.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Record"
      Top             =   300
      Width           =   495
   End
   Begin VB.CommandButton cmdStop 
      Height          =   375
      Left            =   900
      Picture         =   "dlgGreeting.frx":0A1C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Stop"
      Top             =   300
      Width           =   495
   End
   Begin VB.CommandButton cmdPlay 
      Height          =   375
      Left            =   300
      Picture         =   "dlgGreeting.frx":103A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Play"
      Top             =   300
      Width           =   495
   End
End
Attribute VB_Name = "dlgGreeting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPlay_Click()
On Error GoTo EH
    If m_BPlayRec <> False Then
        If (fPlaying = False) Then
          ' -1 specifies the wave mapper
          LoadFile App.Path & "\" & "Greeting.wav"
          Play -1
          m_BPlayRec = True
        End If
    End If
    Exit Sub
EH:
MsgBox CStr(err.Number) & " : " & err.Description
End Sub

Private Sub cmdRecord_Click()
On Error GoTo EH
    If fPlaying = False Then
        RecStart 20, -1, "Greeting.wav"
        m_BPlayRec = False
    End If
    Exit Sub
EH:
    MsgBox CStr(err.Number) & " : " & err.Description
End Sub

Private Sub cmdStop_Click()
Dim nErr As Long
    If m_BPlayRec = False Then
        nErr = waveInReset(h_wavein)
    Else
        nErr = waveOutReset(hWaveOut)
    End If
    'If nErr <> 0 Then MsgBox "Error Resetting Wave Device"
End Sub

Private Sub Form_Load()
    Me.Top = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
    Me.Left = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If frmMain.Timer1.Enabled = True Then Cancel = 1 'Still playing greeting
End Sub
