Attribute VB_Name = "VB_TAPI"
Option Explicit
'Basically TAPI.h goes VB if you are interested each of these constants and
'functions are explained in the Platform SDK.  This is NOT the full TAPI.h file
'though this is from another program so some stuff not used in the answering
'machine app. may be here.

Public Const LINEDIGITMODE_DTMF = &H2

'Public Const HIGHTAPIVERSION = &H20001 'Also available as an upgrade to Win9x
'Public Const WIN95TAPIVERSION = &H10004
'Only support 1.4, with lineInitialize that is the best we can do anyhow.
Public Const TAPIVERSION = &H10004

Public Const LINECALLPRIVILEGE_NONE = &H1
Public Const LINECALLPRIVILEGE_MONITOR = &H2
Public Const LINECALLPRIVILEGE_OWNER = &H4

Public Const LINECALLINFOSTATE_CALLERID = 32768

'LINECALLPARTYID_ Constants
Public Const LINECALLPARTYID_BLOCKED = &H1
Public Const LINECALLPARTYID_OUTOFAREA = &H2
Public Const LINECALLPARTYID_NAME = &H4
Public Const LINECALLPARTYID_ADDRESS = &H8
Public Const LINECALLPARTYID_PARTIAL = &H10
Public Const LINECALLPARTYID_UNKNOWN = &H20
Public Const LINECALLPARTYID_UNAVAIL = &H40

Public Const LINEMEDIAMODE_UNKNOWN = &H2
Public Const LINEMEDIAMODE_INTERACTIVEVOICE = &H4
Public Const LINEMEDIAMODE_AUTOMATEDVOICE = &H8
Public Const LINEMEDIAMODE_DATAMODEM = &H10
Public Const LINEMEDIAMODE_G3FAX = &H20
Public Const LINEMEDIAMODE_TDD = &H40
Public Const LINEMEDIAMODE_G4FAX = &H80
Public Const LINEMEDIAMODE_DIGITALDATA = &H100
Public Const LINEMEDIAMODE_TELETEX = &H200
Public Const LINEMEDIAMODE_VIDEOTEX = &H400
Public Const LINEMEDIAMODE_TELEX = &H800
Public Const LINEMEDIAMODE_MIXED = &H1000
Public Const LINEMEDIAMODE_ADSI = &H2000
Public Const LINEMEDIAMODE_VOICEVIEW = &H4000                          ' TAPI v1.4
''#if (TAPI_CURRENT_VERSION >= 0x00020001)
'Public Const LINEMEDIAMODE_VIDEO = &H8000                              ' TAPI v2.1
''#End If

Public Const LINECALLSELECT_LINE = &H1
Public Const LINECALLSELECT_ADDRESS = &H2
Public Const LINECALLSELECT_CALL = &H4
''#if (TAPI_CURRENT_VERSION > 0x00020000)
'Public Const LINECALLSELECT_DEVICEID = &H8
''#End If

Public Enum TapiEvent

 LINE_ADDRESSSTATE = 0
 LINE_CALLINFO
 LINE_CALLSTATE
 LINE_CLOSE
 LINE_DEVSPECIFIC
 LINE_DEVSPECIFICFEATURE
 LINE_GATHERDIGITS
 LINE_GENERATE
 LINE_LINEDEVSTATE
 LINE_MONITORDIGITS
 LINE_MONITORMEDIA
 LINE_MONITORTONE
 LINE_REPLY
 LINE_REQUEST
 PHONE_BUTTON
 PHONE_CLOSE
 PHONE_DEVSPECIFIC
 PHONE_REPLY
 PHONE_STATE
 
 LINE_CREATE                                          ' TAPI v1.4
 PHONE_CREATE                                         ' TAPI v1.4

End Enum

'LINECALLSTATE Constants
Public Const LINECALLSTATE_IDLE = &H1
Public Const LINECALLSTATE_OFFERING = &H2
Public Const LINECALLSTATE_ACCEPTED = &H4
Public Const LINECALLSTATE_DIALTONE = &H8
Public Const LINECALLSTATE_DIALING = &H10
Public Const LINECALLSTATE_RINGBACK = &H20
Public Const LINECALLSTATE_BUSY = &H40
Public Const LINECALLSTATE_SPECIALINFO = &H80
Public Const LINECALLSTATE_CONNECTED = &H100
Public Const LINECALLSTATE_PROCEEDING = &H200
Public Const LINECALLSTATE_ONHOLD = &H400
Public Const LINECALLSTATE_CONFERENCED = &H800
Public Const LINECALLSTATE_ONHOLDPENDCONF = &H1000
Public Const LINECALLSTATE_ONHOLDPENDTRANSFER = &H2000
Public Const LINECALLSTATE_DISCONNECTED = &H4000
Public Const LINECALLSTATE_UNKNOWN = &H8000

'TAPI Error Constants
Public Const LINEERR_ALLOCATED = &H80000001
Public Const LINEERR_BADDEVICEID = &H80000002
Public Const LINEERR_BEARERMODEUNAVAIL = &H80000003
Public Const LINEERR_CALLUNAVAIL = &H80000005
Public Const LINEERR_COMPLETIONOVERRUN = &H80000006
Public Const LINEERR_CONFERENCEFULL = &H80000007
Public Const LINEERR_DIALBILLING = &H80000008
Public Const LINEERR_DIALDIALTONE = &H80000009
Public Const LINEERR_DIALPROMPT = &H8000000A
Public Const LINEERR_DIALQUIET = &H8000000B
Public Const LINEERR_INCOMPATIBLEAPIVERSION = &H8000000C
Public Const LINEERR_INCOMPATIBLEEXTVERSION = &H8000000D
Public Const LINEERR_INIFILECORRUPT = &H8000000E
Public Const LINEERR_INUSE = &H8000000F
Public Const LINEERR_INVALADDRESS = &H80000010
Public Const LINEERR_INVALADDRESSID = &H80000011
Public Const LINEERR_INVALADDRESSMODE = &H80000012
Public Const LINEERR_INVALADDRESSSTATE = &H80000013
Public Const LINEERR_INVALAPPHANDLE = &H80000014
Public Const LINEERR_INVALAPPNAME = &H80000015
Public Const LINEERR_INVALBEARERMODE = &H80000016
Public Const LINEERR_INVALCALLCOMPLMODE = &H80000017
Public Const LINEERR_INVALCALLHANDLE = &H80000018
Public Const LINEERR_INVALCALLPARAMS = &H80000019
Public Const LINEERR_INVALCALLPRIVILEGE = &H8000001A
Public Const LINEERR_INVALCALLSELECT = &H8000001B
Public Const LINEERR_INVALCALLSTATE = &H8000001C
Public Const LINEERR_INVALCALLSTATELIST = &H8000001D
Public Const LINEERR_INVALCARD = &H8000001E
Public Const LINEERR_INVALCOMPLETIONID = &H8000001F
Public Const LINEERR_INVALCONFCALLHANDLE = &H80000020
Public Const LINEERR_INVALCONSULTCALLHANDLE = &H80000021
Public Const LINEERR_INVALCOUNTRYCODE = &H80000022
Public Const LINEERR_INVALDEVICECLASS = &H80000023
Public Const LINEERR_INVALDEVICEHANDLE = &H80000024
Public Const LINEERR_INVALDIALPARAMS = &H80000025
Public Const LINEERR_INVALDIGITLIST = &H80000026
Public Const LINEERR_INVALDIGITMODE = &H80000027
Public Const LINEERR_INVALDIGITS = &H80000028
Public Const LINEERR_INVALEXTVERSION = &H80000029
Public Const LINEERR_INVALGROUPID = &H8000002A
Public Const LINEERR_INVALLINEHANDLE = &H8000002B
Public Const LINEERR_INVALLINESTATE = &H8000002C
Public Const LINEERR_INVALLOCATION = &H8000002D
Public Const LINEERR_INVALMEDIALIST = &H8000002E
Public Const LINEERR_INVALMEDIAMODE = &H8000002F
Public Const LINEERR_INVALMESSAGEID = &H80000030
Public Const LINEERR_INVALPARAM = &H80000032
Public Const LINEERR_INVALPARKID = &H80000033
Public Const LINEERR_INVALPARKMODE = &H80000034
Public Const LINEERR_INVALPOINTER = &H80000035
Public Const LINEERR_INVALPRIVSELECT = &H80000036
Public Const LINEERR_INVALRATE = &H80000037
Public Const LINEERR_INVALREQUESTMODE = &H80000038
Public Const LINEERR_INVALTERMINALID = &H80000039
Public Const LINEERR_INVALTERMINALMODE = &H8000003A
Public Const LINEERR_INVALTIMEOUT = &H8000003B
Public Const LINEERR_INVALTONE = &H8000003C
Public Const LINEERR_INVALTONELIST = &H8000003D
Public Const LINEERR_INVALTONEMODE = &H8000003E
Public Const LINEERR_INVALTRANSFERMODE = &H8000003F
Public Const LINEERR_LINEMAPPERFAILED = &H80000040
Public Const LINEERR_NOCONFERENCE = &H80000041
Public Const LINEERR_NODEVICE = &H80000042
Public Const LINEERR_NODRIVER = &H80000043
Public Const LINEERR_NOMEM = &H80000044
Public Const LINEERR_NOREQUEST = &H80000045
Public Const LINEERR_NOTOWNER = &H80000046
Public Const LINEERR_NOTREGISTERED = &H80000047
Public Const LINEERR_OPERATIONFAILED = &H80000048
Public Const LINEERR_OPERATIONUNAVAIL = &H80000049
Public Const LINEERR_RATEUNAVAIL = &H8000004A
Public Const LINEERR_RESOURCEUNAVAIL = &H8000004B
Public Const LINEERR_REQUESTOVERRUN = &H8000004C
Public Const LINEERR_STRUCTURETOOSMALL = &H8000004D
Public Const LINEERR_TARGETNOTFOUND = &H8000004E
Public Const LINEERR_TARGETSELF = &H8000004F
Public Const LINEERR_UNINITIALIZED = &H80000050
Public Const LINEERR_USERUSERINFOTOOBIG = &H80000051
Public Const LINEERR_REINIT = &H80000052
Public Const LINEERR_ADDRESSBLOCKED = &H80000053
Public Const LINEERR_BILLINGREJECTED = &H80000054
Public Const LINEERR_INVALFEATURE = &H80000055
Public Const LINEERR_NOMULTIPLEINSTANCE = &H80000056

'LINEDEVSTATE Constants
Public Const LINEDEVSTATE_OTHER = &H1
Public Const LINEDEVSTATE_RINGING = &H2
Public Const LINEDEVSTATE_CONNECTED = &H4
Public Const LINEDEVSTATE_DISCONNECTED = &H8
Public Const LINEDEVSTATE_MSGWAITON = &H10
Public Const LINEDEVSTATE_MSGWAITOFF = &H20
Public Const LINEDEVSTATE_INSERVICE = &H40
Public Const LINEDEVSTATE_OUTOFSERVICE = &H80
Public Const LINEDEVSTATE_MAINTENANCE = &H100
Public Const LINEDEVSTATE_OPEN = &H200
Public Const LINEDEVSTATE_CLOSE = &H400
Public Const LINEDEVSTATE_NUMCALLS = &H800
Public Const LINEDEVSTATE_NUMCOMPLETIONS = &H1000
Public Const LINEDEVSTATE_TERMINALS = &H2000
Public Const LINEDEVSTATE_ROAMMODE = &H4000
Public Const LINEDEVSTATE_BATTERY = &H8000
Public Const LINEDEVSTATE_SIGNAL = &H10000
Public Const LINEDEVSTATE_DEVSPECIFIC = &H20000
Public Const LINEDEVSTATE_REINIT = &H40000
Public Const LINEDEVSTATE_LOCK = &H80000
Public Const LINEDEVSTATE_CAPSCHANGE = &H100000         ' TAPI v1.4
Public Const LINEDEVSTATE_CONFIGCHANGE = &H200000       ' TAPI v1.4
Public Const LINEDEVSTATE_TRANSLATECHANGE = &H400000    ' TAPI v1.4
Public Const LINEDEVSTATE_COMPLCANCEL = &H800000        ' TAPI v1.4
Public Const LINEDEVSTATE_REMOVED = &H1000000           ' TAPI v1.4


Type linedialparams
  dwDialPause As Long
  dwDialSpeed As Long
  dwDigitDuration As Long
  dwWaitForDialtone As Long
End Type

Type lineCallInfo
    
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long
    hLine As Long
    dwLineDeviceID As Long
    dwAddressID As Long
    dwBearerMode As Long
    dwRate As Long
    dwMediaMode As Long
    dwAppSpecific As Long
    dwCallID As Long
    dwRelatedCallID As Long
    dwCallParamFlags As Long
    dwCallStates As Long
    dwMonitorDigitModes As Long
    dwMonitorMediaModes As Long
    DialParams As linedialparams
    dwOrigin As Long
    dwReason As Long
    dwCompletionID As Long
    dwNumOwners As Long
    dwNumMonitors As Long
    dwCountryCode As Long
    dwTrunk As Long
    dwCallerIDFlags As Long
    dwCallerIDSize As Long
    dwCallerIDOffset As Long
    dwCallerIDNameSize As Long
    dwCallerIDNameOffset As Long
    dwCalledIDFlags As Long
    dwCalledIDSize As Long
    dwCalledIDOffset As Long
    dwCalledIDNameSize As Long
    dwCalledIDNameOffset As Long
    dwConnectedIDFlags As Long
    dwConnectedIDSize As Long
    dwConnectedIDOffset As Long
    dwConnectedIDNameSize As Long
    dwConnectedIDNameOffset As Long
    dwRedirectionIDFlags As Long
    dwRedirectionIDSize As Long
    dwRedirectionIDOffset As Long
    dwRedirectionIDNameSize As Long
    dwRedirectionIDNameOffset As Long
    dwRedirectingIDFlags As Long
    dwRedirectingIDSize As Long
    dwRedirectingIDOffset As Long
    dwRedirectingIDNameSize As Long
    dwRedirectingIDNameOffset As Long
    dwAppNameSize As Long
    dwAppNameOffset As Long
    dwDisplayableAddressSize As Long
    dwDisplayableAddressOffset As Long
    dwCalledPartySize As Long
    dwCalledPartyOffset As Long
    dwCommentSize As Long
    dwCommentOffset As Long
    dwDisplaySize As Long
    dwDisplayOffset As Long
    dwUserUserInfoSize As Long
    dwUserUserInfoOffset As Long
    dwHighLevelCompSize As Long
    dwHighLevelCompOffset As Long
    dwLowLevelCompSize As Long
    dwLowLevelCompOffset As Long
    dwChargingInfoSize As Long
    dwChargingInfoOffset As Long
    dwTerminalModesSize As Long
    dwTerminalModesOffset As Long
    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
    
''#if (TAPI_CURRENT_VERSION >= 0x00020000)
'    dwCallTreatment As Long                                ' TAPI v2.0
'    dwCallDataSize As Long                                 ' TAPI v2.0
'    dwCallDataOffset As Long                               ' TAPI v2.0
'    dwSendingFlowspecSize As Long                          ' TAPI v2.0
'    dwSendingFlowspecOffset As Long                        ' TAPI v2.0
'    dwReceivingFlowspecSize As Long                        ' TAPI v2.0
'    dwReceivingFlowspecOffset As Long                      ' TAPI v2.0
''#End If
    bBytes(2000) As Byte 'HACK Added to TAPI structure for callinfo data.

End Type

Type lineextensionid
    dwExtensionID0 As Long
    dwExtensionID1 As Long
    dwExtensionID2 As Long
    dwExtensionID3 As Long
End Type

Type linedevcaps
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long
    dwProviderInfoSize As Long
    dwProviderInfoOffset As Long
    dwSwitchInfoSize As Long
    dwSwitchInfoOffset As Long
    dwPermanentLineID As Long
    dwLineNameSize As Long
    dwLineNameOffset As Long
    dwStringFormat As Long
    dwAddressModes As Long
    dwNumAddresses As Long
    dwBearerModes As Long
    dwMaxRate As Long
    dwMediaModes As Long
    dwGenerateToneModes As Long
    dwGenerateToneMaxNumFreq As Long
    dwGenerateDigitModes As Long
    dwMonitorToneMaxNumFreq As Long
    dwMonitorToneMaxNumEntries As Long
    dwMonitorDigitModes As Long
    dwGatherDigitsMinTimeout As Long
    dwGatherDigitsMaxTimeout As Long
    dwMedCtlDigitMaxListSize As Long
    dwMedCtlMediaMaxListSize As Long
    dwMedCtlToneMaxListSize As Long
    dwMedCtlCallStateMaxListSize As Long
    dwDevCapFlags As Long
    dwMaxNumActiveCalls As Long
    dwAnswerMode As Long
    dwRingModes As Long
    dwLineStates As Long
    dwUUIAcceptSize As Long
    dwUUIAnswerSize As Long
    dwUUIMakeCallSize As Long
    dwUUIDropSize As Long
    dwUUISendUserUserInfoSize As Long
    dwUUICallInfoSize As Long
    MinDialParams As linedialparams
    MaxDialParams As linedialparams
    DefaultDialParams As linedialparams
    dwNumTerminals As Long
    dwTerminalCapsSize As Long
    dwTerminalCapsOffset As Long
    dwTerminalTextEntrySize As Long
    dwTerminalTextSize As Long
    dwTerminalTextOffset As Long
    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
    dwLineFeatures As Long                                 ' TAPI v1.4
''#if (TAPI_CURRENT_VERSION >= 0x00020000)
'    dwSettableDevStatus As Long                            ' TAPI v2.0
'    dwDeviceClassesSize As Long                            ' TAPI v2.0
'    dwDeviceClassesOffset As Long                          ' TAPI v2.0
''#End If
    bBytes(2000) As Byte
End Type

Type varString
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long
    dwStringFormat As Long
    dwStringSize As Long
    dwStringOffset As Long
    bBytes(2000) As Byte 'HACK Added to TAPI structure for lineGetID data.
End Type

Public Declare Function lineMonitorDigits Lib "Tapi32" (ByVal hCall As Long, _
    ByVal dwDigitModes As Long) As Long

Public Declare Function lineGenerateDigits Lib "Tapi32" (ByVal hCall As Long, _
    ByVal dwDigitMode As Long, ByVal lpszDigits As String, ByVal dwDuration _
    As Long) As Long
    
Public Declare Function lineGetCallInfo Lib "Tapi32" (ByVal hCall As Long, _
    ByRef lpCallInf As lineCallInfo) As Long

Public Declare Function lineInitialize Lib "Tapi32" (ByRef hTAPI As Long, _
    ByVal hInst As Long, ByVal fnPtr As Long, ByRef szAppName As Long, _
    ByRef dwNumLines As Long) As Long

Public Declare Function lineNegotiateAPIVersion Lib "Tapi32" _
    (ByVal hTAPI As Long, ByVal dwDeviceID As Long, _
    ByVal dwAPILowVersion As Long, ByVal dwAPIHighVersion As Long, _
    ByRef lpdwAPIVersion As Long, ByRef lpExtensionID As lineextensionid) _
    As Long

Public Declare Function lineOpen Lib "Tapi32" (ByVal hLineApp As Long, _
    ByVal dwDeviceID As Long, ByRef lphLine As Long, ByVal dwAPIVersion As _
    Long, ByVal dwExtVersion As Long, ByRef dwCallbackInstance As Long, _
    ByVal dwPrivileges As Long, ByVal dwMediaModes As Long, _
    ByRef lpCallParams As Long) As Long
    
Public Declare Function lineGetDevCaps Lib "Tapi32" (ByVal hLineApp As Long, _
    ByVal dwDeviceID As Long, ByVal dwAPIVersion As Long, ByVal dwExtVersion _
    As Long, ByRef lpLineDevCaps As linedevcaps) As Long
  
Public Declare Function lineSetStatusMessages Lib "Tapi32" (ByVal hLine As _
    Long, ByVal dwLineStates As Long, ByVal dwAddressStates As Long) As Long

Public Declare Function lineMakeCall Lib "Tapi32" (ByVal hLine As Long, _
    ByRef lphCall As Long, ByVal lpszDestAddress As String, _
    ByVal dwCountryCode As Long, ByVal lpCallParams As Long) As Long
    
Public Declare Function lineDrop Lib "Tapi32" (ByVal hCall As Long, _
    ByVal lpsUserUserInfo As String, ByVal dwSize As Long) As Long

Public Declare Function lineShutdown Lib "Tapi32" (ByVal hLineApp As Long) _
    As Long

Public Declare Function lineAnswer Lib "Tapi32" (ByVal hCall As Long, _
    ByRef lpsUserUserInfo As String, ByVal dwSize As Long) As Long
    
    
Public Declare Function lineGetID Lib "Tapi32" (ByVal hLine As Long, _
    ByVal dwAddressID As Long, ByVal hCall As Long, ByVal dwSelect As Long, _
    ByRef lpDevice As varString, ByVal lpszDeviceClass As String) As Long
    
Public Declare Function lineDeallocateCall Lib "Tapi32" (ByVal hCall As Long) _
    As Long

