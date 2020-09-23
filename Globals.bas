Attribute VB_Name = "Globals"
Option Explicit
'nMessages is used to update the counter frmMain.lblMsgCount while running
Public nMessages As Long
''TAPI Stuff
Public hCall As Long
Public hTAPI As Long
Public lNumLines As Long
Public hLine As Long
Public lpLineDevCaps As linedevcaps
Public lMediaID As Long

''Feature specific
Public CallIDName As String
Public CallIDNumber As String
Public sSecret As String    'Holds the toll saver access code digits
Public lTollSaver As Long   'Counter for validating access code in timer
Public fInTollSaver As Boolean 'Are we currently playing messages over the phone?
Public lNegVer As Long      'Negotiated TAPI version

Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

'*LineCallBack is the callback function passed to lineInitialize.  The heart
'of the TAPI stuff.  TAPI sends all call state to the application through this
'callback. Anything we are interested in we either drill down on to get more
'info, or act on immediately.
Public Function LineCallBack(ByVal dwDevice As Long, ByVal dwMessage As Long, _
    ByVal dwInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long) As Long
  
    Select Case dwMessage
        Case TapiEvent.LINE_ADDRESSSTATE
            DebugString 4, "LINE_ADDRESSSTATE"
        Case TapiEvent.LINE_CALLINFO
            DebugString 3, "LINE_CALLINFO"
            If dwParam1 = LINECALLINFOSTATE_CALLERID Then
            'Got some Caller ID info, GetCallerInfo pulls it out.
                DebugString 4, "LINECALLINFOSTATE_CALLERID"
                GetCallerInfo dwDevice
            Else
                DebugString 5, "LINE_CALLINFO -> " & CStr(dwParam1)
            End If
        Case TapiEvent.LINE_CALLSTATE
            DebugString 3, "LINE_CALLSTATE"
            LineCallStateProc dwDevice, dwInstance, dwParam1, dwParam2, dwParam3
        Case TapiEvent.LINE_CLOSE
            DebugString 4, "LINE_CLOSE"
        Case TapiEvent.LINE_CREATE
            DebugString 4, "LINE_CREATE:"
        Case TapiEvent.LINE_DEVSPECIFIC
            DebugString 4, "LINE_DEVSPECIFIC"
        Case TapiEvent.LINE_DEVSPECIFICFEATURE
            DebugString 4, "LINE_DEVSPECIFICFEATURE"
        Case TapiEvent.LINE_GATHERDIGITS
            DebugString 4, "LINE_GATHERDIGITS"
        Case TapiEvent.LINE_GENERATE
            'We generate a tone after playing the greeting.  This message
            'is thoughtfully provided by TAPI to prompt us to continue on and
            'record a message.
            DebugString 3, "LINE_GENERATE"
            RecordMessage
        Case TapiEvent.LINE_LINEDEVSTATE
            DebugString 3, "LINE_LINEDEVSTATE"
            LineDevStateProc dwDevice, dwInstance, dwParam1, dwParam2, dwParam3
        Case TapiEvent.LINE_MONITORDIGITS
        'Fairly obvious, called everytime a digit is detected.  We get this
        'message because we called lineMonitorDigits when we answered the call.
            DebugString 3, "LINE_MONITORDIGITS -> " & Chr(LoWord(dwParam1))
            sSecret = sSecret & CStr(Chr(LoWord(dwParam1)))
            lTollSaver = 0  'Reset the toll saver counter
        Case TapiEvent.LINE_MONITORMEDIA
            DebugString 4, "LINE_MONITORMEDIA"
        Case TapiEvent.LINE_MONITORTONE
            DebugString 4, "LINE_MONITORTONE"
        Case TapiEvent.LINE_REPLY
        'If dwParam2 is less than zero an error has occured.
        'TODO: raise debug level if dwParam2 is less than zero.
            DebugString 3, "LINE_REPLY -> " & CStr(dwParam2)
        Case TapiEvent.LINE_REQUEST
            DebugString 4, "LINE_REQUEST"
        Case TapiEvent.PHONE_BUTTON
            DebugString 4, "PHONE_BUTTON"
        Case TapiEvent.PHONE_CLOSE
            DebugString 4, "PHONE_CLOSE"
        Case TapiEvent.PHONE_CREATE
            DebugString 4, "PHONE_CREATE"
        Case TapiEvent.PHONE_DEVSPECIFIC
            DebugString 4, "PHONE_DEVSPECIFIC"
        Case TapiEvent.PHONE_REPLY
            DebugString 4, "PHONE_REPLY"
        Case TapiEvent.PHONE_STATE
            DebugString 4, "PHONE_STATE"
        Case Else
            DebugString 2, "Unknown TAPI Event: " & CStr(dwMessage)
    End Select
    
End Function

'*GetCallerInfo takes a handle to a call and calls the TAPI function
' lineGetCallInfo.  This is supposed to be a variable length struct. Not easy
' to do with VB.  The HACK, er.. um.. Solution!
' I have added a bByte() field to the end of the struct and set it to
' 2k.  Hey, if 640k is good enough then 2k is alright...  There are more
' complicated ways to do this that are more dynamic.  CopyMemory could be used.
'HACK: This is a hack but I have used it all over the place, TODO: consolidate
'the hack and I could switch on all of the LINECALLPARTYID_ Constants to see
'if the number/name info is OUTOFAREA, BLOCKED, etc.  Instead I just put
'"UNKNOWN" in the global variables CallIDNumber and CallIDName.
Public Function GetCallerInfo(hCall As Long) As Long
Dim lpCallInfo As lineCallInfo
Dim nErr As Long
Dim sCallerID As String
Dim sCallerName As String
Dim lStart As Long
Dim lLength As Long
Dim lLoop As Long

On Error GoTo EH

    lpCallInfo.dwTotalSize = Len(lpCallInfo)
    nErr = lineGetCallInfo(hCall, lpCallInfo)
        
    'We'll bail here if the 2k HACK isn't big enough
    If nErr <> 0 Then
        ProcessTAPIError nErr
        GetCallerInfo = -1
        Exit Function
    End If
    'Check the LINECALLPARTYID_ Constant to see if we have good info.
    If (lpCallInfo.dwCallerIDFlags And LINECALLPARTYID_ADDRESS) <> False Then
    
    'Find the true length of the lpCallInfo UDT (i.e. subtract the bBytes added)
    lStart = Len(lpCallInfo) - UBound(lpCallInfo.bBytes())
    'With the actual length we can now index into the appended info.
    'Somebody is off by one here, must be the VB start at one instead
    'of zero thing.
    lStart = lpCallInfo.dwCallerIDOffset - lStart + 1
    lLength = lpCallInfo.dwCallerIDSize
        
    'Looping to lLength isn't really needed here, we could just look for
    'the terminating NULL.  Since we aren't verifying the info though,
    'maybe lLength is zero and we win.  This is repeated for the Caller Name.
    For lLoop = 0 To lLength
        If lpCallInfo.bBytes(lStart + lLoop) = 0 Then Exit For
        'Could I simply use Left() here as opposed to the loop?
        sCallerID = sCallerID & CStr(Chr(lpCallInfo.bBytes(lStart + lLoop)))
    Next
    
    CallIDNumber = sCallerID
    
    Else
        CallIDNumber = "UNKNOWN"
    
    End If
        
    'This is an exact repeat of above, except for the name not the number.
    If (lpCallInfo.dwCallerIDFlags And LINECALLPARTYID_NAME) <> False Then
    lStart = Len(lpCallInfo) - UBound(lpCallInfo.bBytes())
    lStart = lpCallInfo.dwCallerIDNameOffset - lStart + 1
    lLength = lpCallInfo.dwCallerIDNameSize
    
    For lLoop = 0 To lLength
        If lpCallInfo.bBytes(lStart + lLoop) = 0 Then Exit For
        sCallerName = sCallerName & CStr(Chr(lpCallInfo.bBytes(lStart + lLoop)))
    Next
    
    CallIDName = sCallerName
    
    Else
        CallIDName = "UNKNOWN"
    End If
        
    DebugString 5, "CallerID:"
    DebugString 4, sCallerID & " : " & sCallerName
    GetCallerInfo = 0
    Exit Function
        
EH:

Debug.Print err.Number & " " & err.Description

GetCallerInfo = 1
End Function

'*LineDevStateProc drills down into the 'LINE_LINEDEVSTATE' messages recieved
' from LineCallBack().  This provides for more verbose logging and we can use
' these state messages to answer the phone, etc...
Public Sub LineDevStateProc(ByVal dwDevice As Long, ByVal dwInstance As Long, _
            ByVal dwParam1 As Long, ByVal dwParam2 As Long, _
            ByVal dwParam3 As Long)

    Select Case dwParam1
        Case LINEDEVSTATE_OTHER:
            DebugString 5, "LINEDEVSTATE_OTHER:"
        Case LINEDEVSTATE_RINGING:  'The only LineDevState case we use..
            DebugString 3, "LINEDEVSTATE_RINGING:"
            DebugString 3, "Ring Count = " & CStr(dwParam3)
            'How many rings we answer on depends on whether we have any
            'messages or not, typical toll saver functionality.
            If frmMain.lblMsgCount > 0 Then
                If dwParam3 >= dlgSetup.txtTollSaver Then Answer
            Else
                If dwParam3 >= dlgSetup.txtRTA Then Answer
            End If
        Case LINEDEVSTATE_CONNECTED:
            DebugString 5, "LINEDEVSTATE_CONNECTED:"
        Case LINEDEVSTATE_DISCONNECTED:
            DebugString 5, "LINEDEVSTATE_DISCONNECTED:"
        Case LINEDEVSTATE_MSGWAITON:
            DebugString 5, "LINEDEVSTATE_MSGWAITON:"
        Case LINEDEVSTATE_MSGWAITOFF:
            DebugString 5, "LINEDEVSTATE_MSGWAITOFF:"
        Case LINEDEVSTATE_INSERVICE:
            DebugString 5, "LINEDEVSTATE_INSERVICE:"
        Case LINEDEVSTATE_OUTOFSERVICE:
            DebugString 5, "LINEDEVSTATE_OUTOFSERVICE:"
        Case LINEDEVSTATE_MAINTENANCE:
            DebugString 5, "LINEDEVSTATE_MAINTENANCE:"
        Case LINEDEVSTATE_OPEN:
            DebugString 5, "LINEDEVSTATE_OPEN:"
        Case LINEDEVSTATE_CLOSE:
            DebugString 5, "LINEDEVSTATE_CLOSE:"
        Case LINEDEVSTATE_NUMCALLS:
            DebugString 5, "LINEDEVSTATE_NUMCALLS:"
        Case LINEDEVSTATE_NUMCOMPLETIONS:
            DebugString 5, "LINEDEVSTATE_NUMCOMPLETIONS:"
        Case LINEDEVSTATE_TERMINALS:
            DebugString 5, "LINEDEVSTATE_TERMINALS:"
        Case LINEDEVSTATE_ROAMMODE:
            DebugString 5, "LINEDEVSTATE_ROAMMODE:"
        Case LINEDEVSTATE_BATTERY:
            DebugString 5, "LINEDEVSTATE_BATTERY:"
        Case LINEDEVSTATE_SIGNAL:
            DebugString 5, "LINEDEVSTATE_SIGNAL:"
        Case LINEDEVSTATE_DEVSPECIFIC:
            DebugString 5, "LINEDEVSTATE_DEVSPECIFIC:"
        Case LINEDEVSTATE_REINIT:
            DebugString 5, "LINEDEVSTATE_REINIT:"
        Case LINEDEVSTATE_LOCK:
            DebugString 5, "LINEDEVSTATE_LOCK:"
        Case LINEDEVSTATE_CAPSCHANGE:
            DebugString 5, "LINEDEVSTATE_CAPSCHANGE:"
        Case LINEDEVSTATE_CONFIGCHANGE:
            DebugString 5, "LINEDEVSTATE_CONFIGCHANGE:"
        Case LINEDEVSTATE_TRANSLATECHANGE:
            DebugString 5, "LINEDEVSTATE_TRANSLATECHANGE:"
        Case LINEDEVSTATE_COMPLCANCEL:
            DebugString 5, "LINEDEVSTATE_COMPLCANCEL:"
        Case LINEDEVSTATE_REMOVED:
            DebugString 5, "LINEDEVSTATE_REMOVED:"
        Case Else:
            DebugString 2, "LINEDEVSTATE_UNKNOWN:"
    End Select
    
End Sub
'Self explanatory, comments are in each state we handle.
Public Sub LineCallStateProc(ByVal dwDevice As Long, ByVal dwInstance As Long, _
                ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)

    Select Case dwParam1
        Case LINECALLSTATE_IDLE:
            DebugString 3, "LINECALLSTATE_IDLE:"
        Case LINECALLSTATE_OFFERING:
            DebugString 3, "LINECALLSTATE_OFFERING:"
            'Set the CallerID stuff now in case we don't get a LINE_CALLINFO msg
            CallIDName = "UNKNOWN"
            CallIDNumber = "UNKNOWN"
            hCall = dwDevice
        Case LINECALLSTATE_ACCEPTED:
            DebugString 4, "LINECALLSTATE_ACCEPTED:"
        Case LINECALLSTATE_DIALTONE:
            DebugString 5, "LINECALLSTATE_DIALTONE:"
        Case LINECALLSTATE_DIALING:
            DebugString 5, "LINECALLSTATE_DIALING:"
        Case LINECALLSTATE_RINGBACK:
            DebugString 5, "LINECALLSTATE_RINGBACK:"
        Case LINECALLSTATE_BUSY:
            DebugString 5, "LINECALLSTATE_BUSY:"
        Case LINECALLSTATE_SPECIALINFO:
            DebugString 4, "LINECALLSTATE_SPECIALINFO:"
        Case LINECALLSTATE_CONNECTED:
            'Answered the call as a result of LINEDEVSTATE_RINGING messages
            'in the LineDevStateProc so now we are connected!
            DebugString 4, "LINECALLSTATE_CONNECTED:"
            fInTollSaver = False
            PlayGreeting
        Case LINECALLSTATE_PROCEEDING:
            DebugString 5, "LINECALLSTATE_PROCEEDING:"
        Case LINECALLSTATE_ONHOLD:
            DebugString 5, "LINECALLSTATE_ONHOLD:"
        Case LINECALLSTATE_CONFERENCED:
            DebugString 5, "LINECALLSTATE_CONFERENCED:"
        Case LINECALLSTATE_ONHOLDPENDCONF:
            DebugString 5, "LINECALLSTATE_ONHOLDPENDCONF:"
        Case LINECALLSTATE_ONHOLDPENDTRANSFER:
            DebugString 5, "LINECALLSTATE_ONHOLDPENDTRANSFER:"
        Case LINECALLSTATE_DISCONNECTED:
        'We dropped the call or the call was dropped on us, either way we
        'simply reset some state variables and wait for the next call.
            DebugString 4, "LINECALLSTATE_DISCONNECTED:"
            Call waveInReset(h_wavein)
            DebugString 5, "m_BPlayRec -> " & CStr(m_BPlayRec)
            If fInTollSaver = True Then
                frmMain.StopPlaying
                DropCall
            End If
        Case LINECALLSTATE_UNKNOWN:
            DebugString 4, "LINECALLSTATE_UNKNOWN:"
    End Select
End Sub
'DebugString, a simple but highly effective logging mechanism
Public Sub DebugString(nSeverity As Long, Message As String)

    OutputDebugString Message & vbCrLf
End Sub
'Give me 16 of those 32 bits
Function LoWord(ByVal dw As Long) As Integer
    If dw And &H8000& Then
        LoWord = dw Or &HFFFF0000
    Else
        LoWord = dw And &HFFFF
    End If
End Function
'Only works for positive numbers
Function LShiftWord(w As Long, c As Integer) As Long
    LShiftWord = w * (2 ^ c)
End Function
'Pick up the phone!
Public Sub Answer()
On Error Resume Next
Dim nError As Long

    nError = lineAnswer(hCall, "", 0)
    If nError < 0 Then ProcessTAPIError nError
    MonitorDigits
End Sub

'Get the wave/in or wave/out lineID so we know where to send the wave output.
Public Sub GetLineID(sWave As String)
On Error Resume Next
Dim nError As Long
Dim sTemp As String
Dim oVar As varString


    oVar.dwTotalSize = Len(oVar)

    nError = lineGetID(hLine, 0, hCall, LINECALLSELECT_CALL, oVar, sWave)
    
    If nError <> 0 Then
        ProcessTAPIError nError
    Else
        If oVar.dwStringOffset = 0 Then 'Nothing to get!
            lMediaID = -1
            Exit Sub
        End If
        sTemp = Trim(Left(oVar.bBytes(0), oVar.dwStringSize))
        lMediaID = sTemp
        DebugString 5, "LineID: " & CStr(lMediaID)
    End If
        
End Sub

Private Sub PlayGreeting()
On Error GoTo EH
    GetLineID "wave/out"

    frmMain.lblName.Caption = CallIDName
    frmMain.lblNumber.Caption = CallIDNumber
    frmMain.LblTime.Caption = Now

    If m_BPlayRec <> False Then
        If (fPlaying = False) Then
          ' -1 specifies the wave mapper
          LoadFile App.Path & "\" & "Greeting.wav"
          Play lMediaID
          m_BPlayRec = True
        End If
    End If
    Exit Sub
EH:
DebugString 0, CStr(err.Number) & " : " & err.Description
End Sub

Private Sub RecordMessage()
On Error GoTo EH
    GetLineID "wave/in"
    nMessages = nMessages + 1
    RecStart dlgSetup.txtMsgLen, lMediaID, CallIDName & "~" & CallIDNumber & "~" _
        & TimeAsString & ".wav"
    Exit Sub
EH:
DebugString 0, CStr(err.Number) & " : " & err.Description
End Sub

'*TimeAsString returns the system time as a string that does not contain
' the illegal filename characters '/' and ':'.  This could blow up on systems
' that have the time settings different, I haven't checked.  A quick glance at
' the 'Regional Options' control panel applet looks like all other time/date
' seps are legal filename characters.
Private Function TimeAsString() As String
Dim lPos As Long
Dim strNow As String
    strNow = Now
    lPos = InStr(1, strNow, "/")
    Do While lPos > 0
    strNow = Mid(strNow, 1, lPos - 1) & "-" & Mid(strNow, lPos + 1, Len(strNow))
    lPos = InStr(1, strNow, "/")
    Loop
    
    lPos = InStr(1, strNow, ":")
    Do While lPos > 0
    strNow = Mid(strNow, 1, lPos - 1) & "-" & Mid(strNow, lPos + 1, Len(strNow))
    lPos = InStr(1, strNow, ":")
    Loop
    
    TimeAsString = strNow
    
End Function

'Simply calls the TAPI lineMonitorDigits function so we can monitor for the
'Remote Access Code.
Private Sub MonitorDigits()
Dim nError As Long
    nError = lineMonitorDigits(hCall, LINEDIGITMODE_DTMF)
    
    If nError <> 0 Then
        ProcessTAPIError nError
    Else
        DebugString 5, "lineMonitorDigits -> Success"
    End If
        
End Sub

'*SendDigit is the tone to signal the caller to leave a message.
Public Sub SendDigit()
Dim nError As Long

    nError = lineGenerateDigits(hCall, LINEDIGITMODE_DTMF, _
                                "#", 0)
    If nError <> 0 Then ProcessTAPIError nError
    
End Sub

'Hang up.  lineDrop isn't enough, we need to deallocate the call as well.  While
'the call is still allcated TAPI retains as much info as possible about the call
'such as caller info data.
Public Sub DropCall()
    Call lineDrop(hCall, "", 0)
    Call lineDeallocateCall(hCall)
End Sub
