VERSION 5.00
Begin VB.Form dlgSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup"
   ClientHeight    =   2190
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4275
   Icon            =   "dlgSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbDevice 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox txtSecret 
      Height          =   285
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtRTA 
      Height          =   285
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtTollSaver 
      Height          =   285
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtMsgLen 
      Height          =   285
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblSecret 
      Alignment       =   1  'Right Justify
      Caption         =   "Remote &Code:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblMsgLen 
      Alignment       =   1  'Right Justify
      Caption         =   "&Max. Message Length:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblTollSaver 
      Alignment       =   1  'Right Justify
      Caption         =   "&Toll Saver Rings:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblRings 
      Alignment       =   1  'Right Justify
      Caption         =   "&Rings To Answer:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "dlgSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'dlgSetup contains a few trivial items (such as which voice card to use) and
'stores/retrieves them from:
'HKEY_CURRENT_USER\Software\VB and VBA Program Settings\VB-TAPI\Settings
Option Explicit

'Center the dialog on the main form and get the settings from the registry.
Private Sub Form_Load()
    Me.Top = frmMain.Top + ((frmMain.Height - Me.Height) / 2)
    Me.Left = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
    txtMsgLen = GetSetting("VB-TAPI", "Settings", "MaxMessage", "60")
    txtRTA = GetSetting("VB-TAPI", "Settings", "NumRings", "5")
    txtTollSaver = GetSetting("VB-TAPI", "Settings", "TollSaver", "3")
    txtSecret = GetSetting("VB-TAPI", "Settings", "Secret", "123")
End Sub
'The Resize event is called when ever the form is shown, not loaded.  This is
'a good thing since the enumeration of the TAPI devices is lengthy and this
'form is used to store/retrieve all are settings at runtime.
'frmMain.myInit is called with a flag denoting its use for this form.  All
'the init function will do is init TAPI and attempt to negotiate a TAPI version
'of 1.4.  See frmMain.myInit for more on 'attempt'.
'Once we have completed initialization we loop through all the devices and put
'their names in the list box cmbDevice.  When this form unloads the index of
'cmbDevice is saved to the registry and that is the device we will use for the
'application.
Private Sub Form_Resize()
Dim lpLineDevCaps As linedevcaps
Dim lLoop As Long
Dim lUnused As Long
Dim nError As Long
On Error GoTo EH

    Screen.MousePointer = vbHourglass
    frmMain.myInit True
    
    For lLoop = 0 To lNumLines
    lpLineDevCaps.dwTotalSize = Len(lpLineDevCaps)
    nError = lineGetDevCaps(hTAPI, lLoop, lNegVer, lUnused, lpLineDevCaps)
    If nError <> 0 Then
        'Skip it
    Else
        Dim sTemp As String
        Dim lTemp As Long
        Dim lStart As Long
        'Find the true length of the lpLineDevCaps UDT (Subtract The Hack!)
        lStart = Len(lpLineDevCaps) - UBound(lpLineDevCaps.bBytes())
        'With the actual length we can now index into the appended info.
        'Somebody is off by one here, must be the VB start at one instead
        'of zero thing.
        lStart = lpLineDevCaps.dwLineNameOffset - lStart + 1

        For lTemp = 0 To lpLineDevCaps.dwLineNameSize
            If lpLineDevCaps.bBytes(lStart + lTemp) = 0 Then Exit For
            'Could I simply use Left() here as opposed to the loop? (Nope ed.)
            sTemp = sTemp & CStr(Chr(lpLineDevCaps.bBytes(lStart + lTemp)))
        Next
       
    End If
    
    DebugString 4, sTemp
    If sTemp = "" Then
        cmbDevice.AddItem "Unknown"
    Else
        cmbDevice.AddItem sTemp
    End If
    sTemp = ""
    Next
    
    Screen.MousePointer = vbNormal
    
    cmbDevice.ListIndex = GetSetting("VB-TAPI", "Settings", "DeviceID", "0")
    
    Exit Sub
EH:
Screen.MousePointer = vbNormal
MsgBox err.Number & " : " & err.Description
    
End Sub
'Make sure we shut down TAPI and reset the mouse cursor
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmMain.myInit True
    Screen.MousePointer = vbNormal
End Sub

Private Sub OKButton_Click()
SaveSetting "VB-TAPI", "Settings", "MaxMessage", txtMsgLen
SaveSetting "VB-TAPI", "Settings", "NumRings", txtRTA
SaveSetting "VB-TAPI", "Settings", "TollSaver", txtTollSaver
SaveSetting "VB-TAPI", "Settings", "Secret", txtSecret
SaveSetting "VB-TAPI", "Settings", "DeviceID", CStr(cmbDevice.ListIndex)

Unload Me
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

'Correct, these do not check if you put in a negative number.  Hey if you don't
'want it to work, don't turn it on...
Private Sub txtMsgLen_Change()
    If Not IsNumeric(txtMsgLen.Text) Then
        Beep
        txtMsgLen = GetSetting("VB-TAPI", "Settings", "MaxMessage", "60")
    End If
End Sub

Private Sub txtRTA_Change()
    If Not IsNumeric(txtRTA.Text) Then
        Beep
        txtRTA = GetSetting("VB-TAPI", "Settings", "NumRings", "5")
    End If
End Sub

Private Sub txtSecret_Change()
    If Not IsNumeric(txtSecret.Text) Then
        Beep
        txtSecret = GetSetting("VB-TAPI", "Settings", "Secret", "123")
    End If
End Sub

Private Sub txtTollSaver_Change()
    If Not IsNumeric(txtTollSaver.Text) Then
        Beep
        txtTollSaver = GetSetting("VB-TAPI", "Settings", "TollSaver", "3")
    End If
End Sub
