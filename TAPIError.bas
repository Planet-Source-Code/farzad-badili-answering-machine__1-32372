Attribute VB_Name = "TAPIError"
Option Explicit
'Simply, switch down the possible errors and report the message.  I think there
'is a shorter way to do this.  This comes from the MSJ Voice sample.
'/////////////////////////////////////////////////////////////////////
'//  ProcessTAPIError - print TAPI error message
'////////////////////////////////////////////////////////////////////
Public Sub ProcessTAPIError(lrc As Long)
Select Case lrc
    Case LINEERR_ALLOCATED
        DebugString 0, " LINEERR_ALLOCATED"
        
    Case LINEERR_BADDEVICEID:
        DebugString 0, " LINEERR_BADDEVICEID"
        
    Case LINEERR_BEARERMODEUNAVAIL:
        DebugString 0, " LINEERR_BEARERMODEUNAVAIL"
        
    Case LINEERR_CALLUNAVAIL:
        DebugString 0, " LINEERR_CALLUNAVAIL"
        
    Case LINEERR_COMPLETIONOVERRUN:
        DebugString 0, " LINEERR_COMPLETIONOVERRUN"
        
    Case LINEERR_CONFERENCEFULL:
        DebugString 0, " LINEERR_CONFERENCEFULL"
        
    Case LINEERR_DIALBILLING:
        DebugString 0, " LINEERR_DIALBILLING"
        
    Case LINEERR_DIALDIALTONE:
        DebugString 0, " LINEERR_DIALDIALTONE"
        
    Case LINEERR_DIALPROMPT:
        DebugString 0, " LINEERR_DIALPROMPT"
        
    Case LINEERR_DIALQUIET:
        DebugString 0, " LINEERR_DIALQUIET"
        
    Case LINEERR_INCOMPATIBLEAPIVERSION:
        DebugString 0, " LINEERR_INCOMPATIBLEAPIVERSION"
        
    Case LINEERR_INCOMPATIBLEEXTVERSION:
        DebugString 0, " LINEERR_INCOMPATIBLEEXTVERSION"
        
    Case LINEERR_INIFILECORRUPT:
        DebugString 0, " LINEERR_INIFILECORRUPT"
        
    Case LINEERR_INUSE:
        DebugString 0, " LINEERR_INUSE"
        
    Case LINEERR_INVALADDRESS:
        DebugString 0, " LINEERR_INVALADDRESS"
        
    Case LINEERR_INVALADDRESSID:
        DebugString 0, " LINEERR_INVALADDRESSID"
        
    Case LINEERR_INVALADDRESSMODE:
        DebugString 0, " LINEERR_INVALADDRESSMODE"
        
    Case LINEERR_INVALADDRESSSTATE:
        DebugString 0, " LINEERR_INVALADDRESSSTATE"
        
    Case LINEERR_INVALAPPHANDLE:
        DebugString 0, " LINEERR_INVALAPPHANDLE"
        
    Case LINEERR_INVALAPPNAME:
        DebugString 0, " LINEERR_INVALAPPNAME"
        
    Case LINEERR_INVALBEARERMODE:
        DebugString 0, " LINEERR_INVALBEARERMODE"
        
    Case LINEERR_INVALCALLCOMPLMODE:
        DebugString 0, " LINEERR_INVALCALLCOMPLMODE"
        
    Case LINEERR_INVALCALLHANDLE:
        DebugString 0, " LINEERR_INVALCALLHANDLE"
        
    Case LINEERR_INVALCALLPARAMS:
        DebugString 0, " LINEERR_INVALCALLPARAMS"
        
    Case LINEERR_INVALCALLPRIVILEGE:
        DebugString 0, " LINEERR_INVALCALLPRIVILEGE"
        
    Case LINEERR_INVALCALLSELECT:
        DebugString 0, " LINEERR_INVALCALLSELECT"
        
    Case LINEERR_INVALCALLSTATE:
        DebugString 0, " LINEERR_INVALCALLSTATE"
        
    Case LINEERR_INVALCALLSTATELIST:
        DebugString 0, " LINEERR_INVALCALLSTATELIST"
        
    Case LINEERR_INVALCARD:
        DebugString 0, " LINEERR_INVALCARD"
        
    Case LINEERR_INVALCOMPLETIONID:
        DebugString 0, " LINEERR_INVALCOMPLETIONID"
        
    Case LINEERR_INVALCONFCALLHANDLE:
        DebugString 0, " LINEERR_INVALCONFCALLHANDLE"
        
    Case LINEERR_INVALCONSULTCALLHANDLE:
        DebugString 0, " LINEERR_INVALCONSULTCALLHANDLE"
        
    Case LINEERR_INVALCOUNTRYCODE:
        DebugString 0, " LINEERR_INVALCOUNTRYCODE"
        
    Case LINEERR_INVALDEVICECLASS:
        DebugString 0, " LINEERR_INVALDEVICECLASS"
        
    Case LINEERR_INVALDEVICEHANDLE:
        DebugString 0, " LINEERR_INVALDEVICEHANDLE"
        
    Case LINEERR_INVALDIALPARAMS:
        DebugString 0, " LINEERR_INVALDIALPARAMS"
        
    Case LINEERR_INVALDIGITLIST:
        DebugString 0, " LINEERR_INVALDIGITLIST"
        
    Case LINEERR_INVALDIGITMODE:
        DebugString 0, " LINEERR_INVALDIGITMODE"
        
    Case LINEERR_INVALDIGITS:
        DebugString 0, " LINEERR_INVALDIGITS"
        
    Case LINEERR_INVALEXTVERSION:
        DebugString 0, " LINEERR_INVALEXTVERSION"
        
    Case LINEERR_INVALGROUPID:
        DebugString 0, " LINEERR_INVALGROUPID"
        
    Case LINEERR_INVALLINEHANDLE:
        DebugString 0, " LINEERR_INVALLINEHANDLE"
        
    Case LINEERR_INVALLINESTATE:
        DebugString 0, " LINEERR_INVALLINESTATE"
        
    Case LINEERR_INVALLOCATION:
        DebugString 0, " LINEERR_INVALLOCATION"
        
    Case LINEERR_INVALMEDIALIST:
        DebugString 0, " LINEERR_INVALMEDIALIST"
        
    Case LINEERR_INVALMEDIAMODE:
        DebugString 0, " LINEERR_INVALMEDIAMODE"
        
    Case LINEERR_INVALMESSAGEID:
        DebugString 0, " LINEERR_INVALMESSAGEID"
        
    Case LINEERR_INVALPARAM:
        DebugString 0, " LINEERR_INVALPARAM"
        
    Case LINEERR_INVALPARKID:
        DebugString 0, " LINEERR_INVALPARKID"
        
    Case LINEERR_INVALPARKMODE:
        DebugString 0, " LINEERR_INVALPARKMODE"
        
    Case LINEERR_INVALPOINTER:
        DebugString 0, " LINEERR_INVALPOINTER"
        
    Case LINEERR_INVALPRIVSELECT:
        DebugString 0, " LINEERR_INVALPRIVSELECT"
        
    Case LINEERR_INVALRATE:
        DebugString 0, " LINEERR_INVALRATE"
        
    Case LINEERR_INVALREQUESTMODE:
        DebugString 0, " LINEERR_INVALREQUESTMODE"
        
    Case LINEERR_INVALTERMINALID:
        DebugString 0, " LINEERR_INVALTERMINALID"
        
    Case LINEERR_INVALTERMINALMODE:
        DebugString 0, " LINEERR_INVALTERMINALMODE"
        
    Case LINEERR_INVALTIMEOUT:
        DebugString 0, " LINEERR_INVALTIMEOUT"
        
    Case LINEERR_INVALTONE:
        DebugString 0, " LINEERR_INVALTONE"
        
    Case LINEERR_INVALTONELIST:
        DebugString 0, " LINEERR_INVALTONELIST"
        
    Case LINEERR_INVALTONEMODE:
        DebugString 0, " LINEERR_INVALTONEMODE"
        
    Case LINEERR_INVALTRANSFERMODE:
        DebugString 0, " LINEERR_INVALTRANSFERMODE"
        
    Case LINEERR_LINEMAPPERFAILED:
        DebugString 0, " LINEERR_LINEMAPPERFAILED"
        
    Case LINEERR_NOCONFERENCE:
        DebugString 0, " LINEERR_NOCONFERENCE"
        
    Case LINEERR_NODEVICE:
        DebugString 0, " LINEERR_NODEVICE"
        
    Case LINEERR_NODRIVER:
        DebugString 0, " LINEERR_NODRIVER"
        
    Case LINEERR_NOMEM:
        DebugString 0, " LINEERR_NOMEM"
        
    Case LINEERR_NOREQUEST:
        DebugString 0, " LINEERR_NOREQUEST"
        
    Case LINEERR_NOTOWNER:
        DebugString 0, " LINEERR_NOTOWNER"
        
    Case LINEERR_NOTREGISTERED:
        DebugString 0, " LINEERR_NOTREGISTERED"
        
    Case LINEERR_OPERATIONFAILED:
        DebugString 0, " LINEERR_OPERATIONFAILED"
        
    Case LINEERR_OPERATIONUNAVAIL:
        DebugString 0, " LINEERR_OPERATIONUNAVAIL"
        
    Case LINEERR_RATEUNAVAIL:
        DebugString 0, " LINEERR_RATEUNAVAIL"
        
    Case LINEERR_RESOURCEUNAVAIL:
        DebugString 0, " LINEERR_RESOURCEUNAVAIL"
        
    Case LINEERR_REQUESTOVERRUN:
        DebugString 0, " LINEERR_REQUESTOVERRUN"
        
    Case LINEERR_STRUCTURETOOSMALL:
        DebugString 0, " LINEERR_STRUCTURETOOSMALL"
        
    Case LINEERR_TARGETNOTFOUND:
        DebugString 0, " LINEERR_TARGETNOTFOUND"
        
    Case LINEERR_TARGETSELF:
        DebugString 0, " LINEERR_TARGETSELF"
        
    Case LINEERR_UNINITIALIZED:
        DebugString 0, " LINEERR_UNINITIALIZED"
        
    Case LINEERR_USERUSERINFOTOOBIG:
        DebugString 0, " LINEERR_USERUSERINFOTOOBIG"
        
    Case LINEERR_REINIT:
        DebugString 0, " LINEERR_REINIT"
        
    Case LINEERR_ADDRESSBLOCKED:
        DebugString 0, " LINEERR_ADDRESSBLOCKED"
        
    Case LINEERR_BILLINGREJECTED:
        DebugString 0, " LINEERR_BILLINGREJECTED"
        
    Case LINEERR_INVALFEATURE:
        DebugString 0, " LINEERR_INVALFEATURE"
        
    Case LINEERR_NOMULTIPLEINSTANCE:
        DebugString 0, " LINEERR_NOMULTIPLEINSTANCE"

    Case Else   'If you see this is it probably a programming error.
        DebugString 0, " Unknown TAPI Error"
        
    End Select
End Sub
