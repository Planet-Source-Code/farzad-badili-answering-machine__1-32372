Attribute VB_Name = "Wave"
'All the stuff for Wave file manipulation is here!  Any of the types and
'functions befor the line:
'App. Specific
'can be looked up in the platform SDK and from MMSystem.h
'
'Record and playback are hard coded for the
'CCITT u-Law 8.000 kHz, 8 Bit, Mono CODEC
'
'According to Microsoft Knowledge Base Article ID: Q142745
'The Microsoft Consultative Committee for International Telephone and Telegraph
'(CCITT) G.711 A-Law and u-Law codec can also achieve only a 2:1 compression
'ratio, but is best when compatibility with current Telephony Application
'Programming Interface (TAPI) standards is a concern.
'
'And the fact that it works with my voice and sound cards.

Option Explicit

Public Const CALLBACK_FUNCTION = &H30000

Public Const WAVE_FORMAT_QUERY = &H1
Public Const WAVE_MAPPED = &H4
Public Const SND_FILENAME = &H20000

Public Const MMIO_READ = &H0
Public Const MMIO_FINDCHUNK = &H10
Public Const MMIO_FINDRIFF = &H20

Public Const MM_WOM_DONE = &H3BD
Public Const MM_WIM_OPEN = &H3BE
Public Const MM_WIM_CLOSE = &H3BF
Public Const MM_WIM_DATA = &H3C0

Public Const WAVE_FORMAT_PCM = &H1

Public Const GMEM_FIXED = &H0
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = GMEM_FIXED Or GMEM_ZEROINIT

Private Type FileHeader
    dwRiff As Long
    dwFileSize As Long
    dwWave As Long
    dwFormat As Long
    dwFormatLength As Long
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    dwData As Long
    dwDataLength As Long
End Type

Type WAVEHDR
   lpData As Long
   dwBufferLength As Long
   dwBytesRecorded As Long
   dwUser As Long
   dwFlags As Long
   dwLoops As Long
   lpNext As Long
   Reserved As Long
End Type

Type WAVEINCAPS
   wMid As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * 32
   dwFormats As Long
   wChannels As Integer
End Type

Type WAVEFORMAT
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type

Type mmioinfo
   dwFlags As Long
   fccIOProc As Long
   pIOProc As Long
   wErrorRet As Long
   htask As Long
   cchBuffer As Long
   pchBuffer As String
   pchNext As String
   pchEndRead As String
   pchEndWrite As String
   lBufOffset As Long
   lDiskOffset As Long
   adwInfo(4) As Long
   dwReserved1 As Long
   dwReserved2 As Long
   hmmio As Long
End Type

Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type

Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, _
    ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, _
    ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
                                             
Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, _
    lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
                                                      
Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, _
    lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
                                                   
Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, _
    lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
                                          
Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long

Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long

Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long

Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long

Declare Function waveOutOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal _
    uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, _
    ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
    
Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, _
    lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
    
Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long

Declare Function waveOutStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long

Declare Function waveOutStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long

Declare Function waveOutPause Lib "winmm.dll" (ByVal hWaveOut As Long) As Long

Declare Function waveOutRestart Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
 
Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, _
    lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
    
Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long

Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" _
    (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
    
Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long

Declare Function waveOutGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" _
    (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
    
Declare Function waveOutAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, _
    lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
    
Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, _
    lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
    
Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags _
    As Long) As Long
    
Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As _
    MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
    
Declare Function mmioDescendParent Lib "winmm.dll" Alias "mmioDescend" (ByVal _
    hmmio As Long, lpck As MMCKINFO, ByVal x As Long, ByVal uFlags As Long) As Long
    
Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName _
    As String, lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long
    
Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As _
    Long, ByVal cch As Long) As Long
    
Declare Function mmioReadFormat Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio _
    As Long, ByRef pch As WAVEFORMAT, ByVal cch As Long) As Long
    
Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" _
    (ByVal sz As String, ByVal uFlags As Long) As Long
    
Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As _
    MMCKINFO, ByVal uFlags As Long) As Long
    
Declare Function apisndPlaySound Lib "winmm" Alias _
         "sndPlaySoundA" (ByVal filename As String, ByVal snd_async _
         As Long) As Long
     

Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal _
                                                dwBytes As Long) As Long
                                                
Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long

Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
         pvDest As Any, ByVal pvSource As Any, ByVal lBytes As Long)
   
'TODO:   Double Declare!!
Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As _
    Any, ByVal ptr As Long, ByVal cb As Long)
   
'App Specific
   
Public whdr As WAVEHDR      'Our wave in header
Public format As WAVEFORMAT 'used for playing and recording
Dim rc As Long              'Global error code
Dim msg As String * 200     'For error message lookup from the Wave API

' variables for managing wave file
Public hWaveOut As Long     'Handle to the wave out device
Dim bufferIn As Long        'wave in buffer pointer
Public hmem As Long         'Allocated memory
Dim outHdr As WAVEHDR       'Wave out header
Public h_wavein As Long     'Handle to the wave in device
Public numSamples As Long   'number of samples to play
Public fFileLoaded As Boolean   'Did we load the file okay?
'fPlaying and m_BPlayRec unfortunately can not be the same, see frmMain.Timer1
Public fPlaying As Boolean  'Are we currently playing a file
Public m_BPlayRec As Boolean    'True for playing, False for recording
Public m_FileName As String 'Recorded file name

'This may be hard to recognize anymore but it comes from the Platform SDK
'under the Multimedia, DirectSound VB sample StreamTo.  Check that for more
'info.  As always, any errors and comments are all mine.
'On Error Resume Next is used because:
'If the file is opened by another app, this one's going to fail bad.  No point
'in throwing an exception though.  Just ignore it, so the new greeting doesn't
'get saved, the old one probably won't get deleted either.  This shouldn't
'happen with a new message since we are guaranteed a unique message name, but
'the system time could get changed, etc...
Public Sub SaveToFileAsStream(sName As String)

On Error Resume Next
    Dim fh As FileHeader, Status As Long
    Dim fFile, File_1Holder() As Byte
    Dim nErr As Long
    
    fFile = FreeFile
    
    If dlgGreeting.Visible = True Then  'Recording a greeting
        Kill App.Path + "\" + sName
        Open App.Path + "\" + sName For Binary Access Write As #fFile
    Else
        Open App.Path + "\Messages\" + sName For Binary Access Write As #fFile
    End If
    
    'RIFF specific
    fh.dwRiff = &H46464952          'RIFF
    fh.dwWave = &H45564157          'WAVE
    fh.dwFormat = &H20746D66        'fmt_chnk
    fh.dwFormatLength = 16
    'End RIFF specific
    
    'The format tag for u-Law isn't in MMSystem.H
    fh.wFormatTag = 7
    fh.nChannels = 1
    fh.nSamplesPerSec = 8000
    fh.nAvgBytesPerSec = 8000
    fh.wBitsPerSample = 8
    fh.nBlockAlign = 1
    'RIFF specific
    fh.dwData = &H61746164            '                 // data_chnk
    fh.dwDataLength = whdr.dwBytesRecorded
    fh.dwFileSize = whdr.dwBytesRecorded + Len(fh)
    'End of header!
    
    ReDim File_1Holder(whdr.dwBytesRecorded + 1)
    CopyMemory File_1Holder(0), whdr.lpData, whdr.dwBytesRecorded + 1

    Put #fFile, , fh
    Put #fFile, , File_1Holder()
        
    nErr = GlobalFree(whdr.lpData)
    
    If nErr <> 0 Then DebugString 2, "SaveToFileAsStream leaked memory"
    
    Close #fFile
    
End Sub

'* waveOutProc is the callback for playing wave files.  Don't even breathe in
' this sub.  Anything that can cause a hardware interrupt can GPF the app.  All
' we are looking for here is the MM_WOM_DONE message then we set a flag that
' causes the timer routine in frmMain to follow the state machine.
Sub waveOutProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance _
                As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long)
' Wave IO Callback function
   If (uMsg = MM_WOM_DONE) Then
      fPlaying = False
   End If
End Sub

'* waveInProc is just as sensitive as waveOutProc.  The 'Debug.Print "MM_WIM_DATA"
'statement will cause the app to hang.  To demonstrate uncomment it, record a
'new greeting and click the stop button before the timer expires.  "MM_WIM_DATA"
'is all we look for, set the waveheader to the one delivered in the message and
'set our CleanUp flag which gets serviced in the frmMain timer proc.
Sub waveInProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance _
                As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long)
                
    Select Case uMsg
        Case MM_WIM_OPEN
            'Debug.Print "MM_WIM_OPEN"
        Case MM_WIM_CLOSE
            'Debug.Print "MM_WIM_CLOSE"
        Case MM_WIM_DATA
            'Debug.Print "MM_WIM_DATA"
            whdr = hdr
            frmMain.CleanUp = 1
    End Select
    
End Sub

'Both CloseWaveOut and CloseWaveIn shutdown the respective wave device.  All
'errors are pretty much ignored, what can you do?  With a Dialogic 160/SC you
'reboot, with a Creative Labs DI 5630 it is safe to ignore...  These are high
'priority errors, given the severe nature that Dialogic treats them (which is
'most likely correct)  Though the application calls these functions at times
'that may not be appropriate.  In other words, an error log indicating an error
'from these two functions may be safe to ignore.
Sub CloseWaveOut()
' Close the waveout device
    rc = waveOutReset(hWaveOut)
    If rc <> 0 Then DebugString 0, CStr(rc) & ": " & "waveOutReset Error"
    rc = waveOutUnprepareHeader(hWaveOut, outHdr, Len(outHdr))
    If rc <> 0 Then DebugString 0, CStr(rc) & ": " & "waveOutUnprepareHeader Error"
    rc = waveOutClose(hWaveOut)
    If rc <> 0 Then DebugString 0, CStr(rc) & ": " & "waveOutClose Error"
End Sub

Sub CloseWaveIn()
    rc = waveInReset(h_wavein)
    If rc <> 0 Then DebugString 0, CStr(rc) & ": " & "waveInReset Error"
    rc = waveInUnprepareHeader(h_wavein, whdr, Len(whdr))
    If rc <> 0 Then DebugString 0, CStr(rc) & ": " & "waveInUnprepareHeader Error"
    rc = waveInClose(h_wavein)
    If rc <> 0 Then DebugString 0, CStr(rc) & ": " & "waveInClose Error"
End Sub

'The two subs, LoadFile and Play are from the Microsoft Knowledge Base Article
'ID: Q182983
'FILE: Playwave.exe Demonstrates How To Play a Sound File
'See that for more info.  Any comments and/or errors are all mine though.
Sub LoadFile(inFile As String)
' Load wavefile into memory

Dim mmckinfoParentIn As MMCKINFO
Dim mmckinfoSubchunkIn As MMCKINFO

   Dim hmmioIn As Long
   Dim mmioinf As mmioinfo
   
   fFileLoaded = False
   
   If (inFile = "") Then
       GlobalFree (hmem)
       Exit Sub
   End If
       
   ' Open the input file
   hmmioIn = mmioOpen(inFile, mmioinf, MMIO_READ)
   If hmmioIn = 0 Then
       err.Raise mmioinf.wErrorRet, "LoadFile", _
       "Error opening input file: " & App.Path & inFile
       Exit Sub
   End If
   
   ' Check if this is a wave file
   mmckinfoParentIn.fccType = mmioStringToFOURCC("WAVE", 0)
   rc = mmioDescendParent(hmmioIn, mmckinfoParentIn, 0, MMIO_FINDRIFF)
   If (rc <> 0) Then
       rc = mmioClose(hmmioIn, 0)
       err.Raise -1, "LoadFile", "Not a WAVE file"
       Exit Sub
   End If
   
   ' Get format info
   mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("fmt", 0)
   rc = mmioDescend(hmmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
   If (rc <> 0) Then
       rc = mmioClose(hmmioIn, 0)
       err.Raise -1, "LoadFile", "Couldn't get format chunk"
       Exit Sub
   End If
   rc = mmioReadFormat(hmmioIn, format, Len(format))
   If (rc = -1) Then
      rc = mmioClose(hmmioIn, 0)
      err.Raise -1, "LoadFile", "Error reading format"
      Exit Sub
   End If
   rc = mmioAscend(hmmioIn, mmckinfoSubchunkIn, 0)
   
   ' Find the data subchunk
   mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("data", 0)
   rc = mmioDescend(hmmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
   If (rc <> 0) Then
      rc = mmioClose(hmmioIn, 0)
      err.Raise -1, "LoadFile", "Couldn't get data chunk"
      Exit Sub
   End If
   
   ' Allocate soundbuffer and read sound data
   GlobalFree hmem
   hmem = GlobalAlloc(&H40, mmckinfoSubchunkIn.ckSize)
   bufferIn = GlobalLock(hmem)
   rc = mmioRead(hmmioIn, bufferIn, mmckinfoSubchunkIn.ckSize)
   
   numSamples = mmckinfoSubchunkIn.ckSize / format.nBlockAlign
   
   ' Close file
   rc = mmioClose(hmmioIn, 0)
   
   If rc <> 0 Then DebugString 2, "Error: " & rc & " in Wave->LoadFile"

   
   fFileLoaded = True
    
End Sub

'The two subs, LoadFile and Play are from the Microsoft Knowledge Base Article
'ID: Q182983
'FILE: Playwave.exe Demonstrates How To Play a Sound File
'See that for more info.  Any comments and/or errors are all mine though.
'TODO: Use the waveOutGetErrorText (which is an alias for waveIn...) error
'function to return richer error info.
Sub Play(ByVal soundcard As Integer)

Dim lFlags As Long
    
'Soundcards seem to dislike the WAVE_MAPPED flag, though Voice Cards (modems)
'work with it, or in my case don't work without it.  Same thing with wave/in...
'I am curious as to how many bugs this causes, if you get around a bug by
'fiddling with the next few lines (or their counter parts in RecStart) then
'let me know sfrare@yahoo.com  I don't think it is caused by using the
'WAVE_MAPPER ID I did check it but not very well.
    If soundcard = -1 Then
        lFlags = CALLBACK_FUNCTION
    Else
        lFlags = CALLBACK_FUNCTION Or WAVE_MAPPED
    End If
    
    format.cbSize = 0   'HACK  I have messed up on the save or record somewhere
    
    rc = waveOutOpen(hWaveOut, soundcard, format, AddressOf waveOutProc, 0, lFlags)
    If (rc <> 0) Then
      GlobalFree (hmem)
      waveOutGetErrorText rc, msg, Len(msg)
      err.Raise rc, "Play", msg & ""
      Exit Sub
    End If

    outHdr.lpData = bufferIn
    outHdr.dwBufferLength = numSamples * format.nBlockAlign
    outHdr.dwFlags = 0
    outHdr.dwLoops = 0

    rc = waveOutPrepareHeader(hWaveOut, outHdr, Len(outHdr))
    If (rc <> 0) Then
      waveOutGetErrorText rc, msg, Len(msg)
      err.Raise rc, "Play", msg & ""
      Exit Sub
    End If

    rc = waveOutWrite(hWaveOut, outHdr, Len(outHdr))
    If (rc <> 0) Then
      GlobalFree (hmem)
    Else
      fPlaying = True
      frmMain.Timer1.Enabled = True
    End If
End Sub

'If my life depended on it I couldn't find a sample of recording a wave file
'in VB.  The RecStart and AddWaveInBuffer procedures come from the MSJ Voice
'sample application (converted from C of course).  In that application both
'functions were in the same procedure, I split them up so that it would be
'easier to work with multiple buffers if I decide to do so in the future.
'Again, any errors and comments are all mine.
'TODO: Use the waveOutGetErrorText (which is an alias for waveIn...) error
'function to return richer error info.
Public Sub RecStart(nMaxLength As Long, dwDevice As Long, sFileName As String)
Dim lFlags As Long

Dim bufsz As Long
Dim hData As Long

    If dwDevice = -1 Then
        lFlags = CALLBACK_FUNCTION
    Else
        lFlags = CALLBACK_FUNCTION Or WAVE_MAPPED
    End If
  
    frmMain.CleanUp = 0
    m_FileName = sFileName

    format.wFormatTag = 7
    format.nChannels = 1
    format.wBitsPerSample = 8
    format.nSamplesPerSec = 8000
    format.nBlockAlign = 1
    format.nAvgBytesPerSec = 8000
    format.cbSize = 0

    Dim myNull As Long
    
    'Since we only deal with one wave format, querying the device is folly.
'    rc = waveInOpen(myNull, dwDevice, format, frmMain.hWnd, 0, WAVE_FORMAT_QUERY)
'
'    If rc <> 0 Then
'        err.Raise nErr, "RecStart", "Unsupported Wave Format"
'        Exit Sub
'    End If
    
    rc = waveInOpen(h_wavein, dwDevice, format, AddressOf waveInProc, 0, lFlags)
    If rc <> 0 Then
        'Okay, I would have better debug info had I queried the device...
        'Or better yet used the waveOutGetErrorText function
        err.Raise rc, "RecStart", "Can Not Open the Device"
        Exit Sub
    End If
      
    'I don't know, math.. It is all VooDoo to me ;->
    bufsz = nMaxLength * format.nSamplesPerSec * format.wBitsPerSample / 8
    
    AddWaveInBuffer bufsz, whdr
    
    rc = waveInStart(h_wavein)
    
    If rc <> 0 Then
        err.Raise rc, "RecStart", "Error in waveInStart"
        Exit Sub
    End If
    
    frmMain.Timer1.Enabled = True

    DebugString 5, "RecStart Completed"
End Sub

'This sub allocates the memory for and calls the appropriate Wave functions
'to submit a buffer for wave use.  You can use multiple smaller buffers (a
'good thing) or one huge one as I do (a bad thing)...  I have read, but not
'tried, that multiple smaller buffers result in less CPU usage.  Since this
'puppy uses ~10-20% of a 400Mhz CPU when recording one would hope there was
'a way to put a head on that.
'TODO: Use the waveOutGetErrorText (which is an alias for waveIn...) error
'function to return richer error info.
Private Sub AddWaveInBuffer(bufsz As Long, whdr As WAVEHDR)
    Dim nErr As Long
    Dim hData As Long
    
    hData = GlobalAlloc(GPTR, bufsz)

    If hData = 0 Then
        err.Raise nErr, "AddWaveInBuffer", "Can Not Allocate Memory"
        Exit Sub
    End If
    
    whdr.lpData = GlobalLock(hData)
    
    If whdr.lpData = 0 Then
        err.Raise nErr, "AddWaveInBuffer", "Can Not Lock Memory"
        Exit Sub
    End If
        
    whdr.dwBufferLength = bufsz
    
    nErr = waveInPrepareHeader(h_wavein, whdr, Len(whdr))
    
    If nErr <> 0 Then
        err.Raise nErr, "AddWaveInBuffer", "Error in waveInPrepareHeader"
        Exit Sub
    End If
    
    nErr = waveInAddBuffer(h_wavein, whdr, Len(whdr))
    
    If nErr <> 0 Then
        err.Raise nErr, "AddWaveInBuffer", "Error in waveInAddBuffer"
        Exit Sub
    End If
    
    whdr.dwUser = hData

End Sub

'These next three functions simply call their Win32 counterparts.  All errors
'are ignored.  It takes to much babysitting to see if these functions are
'called at an appropriate time to do error checking.  Such as, of course you
'will get a waveOutReset error if you attempt to stop a wave file when none
'are playing...
Sub PausePlay()
    waveOutPause (hWaveOut)
End Sub
Sub ResumePlay()
    waveOutRestart (hWaveOut)
End Sub
Sub StopPlay()
   waveOutReset (hWaveOut)
End Sub
