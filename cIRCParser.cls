VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIRCParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'IRC Parsing Engine Class
'©2007 - Jason James Newland
'
'Purpose:
'       The purpose of this class is to parse the individual
'       events coming from an IRC session and splitting them
'       up so it just raises the relevant events.
'       Saves large parsing subs, especially when it comes to
'       PRIVMSG, as you have to parse private/channel action/text,
'       CTCP's and DCC's.
'
'Usage: (Declarations)
'       Private WithEvents cParser As cIRCParser
'Form/Class load:
'       Set cParser = New cIRCParser
'For/Class unload:
'       Set cParser = Nothing
'
'Method:
'       To parse socket data, get the data into a string, pass that
'       directly to ParseRAWSocketData sVariableName.
'       When parsed, it will raise SocketDataOutput sFirst, sSecond, sThird, sFourth.
'       The reason for this is so an IRC client can override any
'       further parsing for various reasons.
'
'       Just simply pass the output of this event back to the Parse
'       routine of this class.
'
'NOTE:
'       strMe is a string that is set to your current IRC
'       nick name.
Public strMe As String

Private WithEvents IRCSocket As CSocket
Attribute IRCSocket.VB_VarHelpID = -1
Private WithEvents tmrParse As CLiteTimer
Attribute tmrParse.VB_VarHelpID = -1
Private WithEvents tmrRaiseDiscon As CLiteTimer
Attribute tmrRaiseDiscon.VB_VarHelpID = -1
Private WithEvents tmrCloseSock As CLiteTimer
Attribute tmrCloseSock.VB_VarHelpID = -1
Private cQueueMSG As New Collection
Private blClose As Boolean
'
'events that will be raised
Public Event SocketDataOutput(ByVal sFirst As String, sSecond As String, sThird As String, sFourth As String)

Public Event IRCSocketConnect()
Public Event IRCSocketDisconnect()
Public Event IRCSocketCancel()
Public Event IRCSocketServerDisconnect(ByVal sNumber As Integer, sErrorMSG As String)

'Main IRC events
Public Event IRCCloseLink(ByVal sLink As String, sErrorMSG As String)
Public Event IRCPing(ByVal sText As String)
Public Event IRCJoin(ByVal ssServer As String, sNick As String, sAddress As String, sChannel As String)
Public Event IRCPart(ByVal sNick As String, sAddress As String, sChannel As String, sText As String)
Public Event IRCKick(ByVal sNick As String, sKnick As String, sChannel As String, sText As String)
Public Event IRCNick(ByVal sNick As String, sNewNick As String)
Public Event IRCQuit(ByVal sNick As String, sAddress As String, sText As String)
Public Event IRCMode(ByVal sNick As String, sChannel As String, sModeString As String, sData As String)
Public Event IRCSelfMode(ByVal sNick As String, sModeString As String)
Public Event IRCNotice(ByVal sNick As String, sAddress As String, sText As String)
Public Event IRCSNotice(ByVal ssServer As String, sText As String)
Public Event IRCTopic(ByVal sNick As String, sChannel As String, sText As String)
Public Event IRCInvite(ByVal sNick As String, sAddress As String, sChannel As String)

'private message events
Public Event IRCTextChan(ByVal sNick As String, sAddress As String, sChannel As String, sText As String)
Public Event IRCActionChan(ByVal sNick As String, sAddress As String, sChannel As String, sText As String)
Public Event IRCTextQuery(ByVal sNick As String, sAddress As String, sText As String)
Public Event IRCActionQuery(ByVal sNick As String, sAddress As String, sText As String)

'ctcps
Public Event IRCCTCPPing(ByVal sNick As String, sAddress As String, sText As String)
Public Event IRCCTCPTime(ByVal sNick As String, sAddress As String)
Public Event IRCCTCPVersion(ByVal sNick As String, sAddress As String)
Public Event IRCCTCPFinger(ByVal sNick As String, sAddress As String)
Public Event IRCCTCPAway(ByVal sNick As String, sAddress As String, sText As String)
Public Event IRCCTCPOther(ByVal sNick As String, sAddress As String, sText As String)

Public Event IRCCTCPReply(ByVal sNick As String, sCTCP As String, sText As String)

'DCC
Public Event IRCDCCSend(ByVal sNick As String, sAddress As String, sFile As String, sNickIP As String, sPort As String, sFSize As String)
Public Event IRCDCCChat(ByVal sNick As String, sAddress As String, sNickIP As String, sPort As String)
Public Event IRCDCCResume(ByVal sNick As String, sAddress As String, sFile As String, sPort As String, sPos As String)
Public Event IRCDCCAccept(ByVal sNick As String, sAddress As String, sFile As String, sPort As String, sPos As String)

'RAW events
Public Event IRCUnknown(ByVal ssServer As String, sText As String)

Public Event IRCNotifyUpdate(ByVal ssServer As String, sNick As String, sAddress As String)
Public Event IRCNotify(ByVal ssServer As String, sNick As String, sAddress As String)
Public Event IRCUNotify(ByVal ssServer As String, sNick As String, sAddress As String)

Public Event IRCWelcome1(ByVal ssServer As String, sText As String)
Public Event IRCWelcome2(ByVal ssServer As String, sText As String)

Public Event IRCProtocols(ByVal ssServer As String, sText As String)

Public Event IRCMOTDStart(ByVal ssServer As String, sText As String)
Public Event IRCMOTDText(ByVal ssServer As String, sText As String)
Public Event IRCMOTDEnd(ByVal ssServer As String, sText As String)
Public Event IRCMOTDError(ByVal ssServer As String, sText As String)

Public Event IRCLusers(ByVal ssServer As String, sText As String)

Public Event IRCWho(ByVal sNick As String, sChannel As String, sAddress As String)
Public Event IRCWhoEnd(ByVal sNick As String, sChannel As String)

Public Event IRCNames(ByVal sChannel As String, sNames As String)
Public Event IRCNamesEnd(ByVal sChannel As String)

Public Event IRCChannelList(ByVal ssServer As String, sChannel As String, sUsers As String, sTopic As String)
Public Event IRCChannelListEnd(ByVal ssServer As String, sText As String)

Public Event IRCBanList(ByVal sChannel As String, sBanMask As String, sNick As String, sTime As String)
Public Event IRCBanListEnd(ByVal sChannel As String, sText As String)

Public Event IRCChannelModes(ByVal sChannel As String, sModeString As String)
Public Event IRCCannotJoin(ByVal sChannel As String, sReason)
Public Event IRCTopicIs(ByVal sChannel As String, sText As String)
Public Event IRCTopicSet(ByVal sChannel As String, sNick As String, sTime As String)

Public Event IRCRawOther(ByVal ssServer As String, sNumeric As String, sText As String)

'socket control
Public Sub IRCSocketConnect(ByVal ssServer As String, Optional sPort As String)
    On Error Resume Next
    '
    blClose = False
    '
    IRCSocket.Connect ssServer, IIf(LenB(sPort) <> 0, sPort, vbNullString)
End Sub

Public Sub IRCSocketDisconnect(ByVal sQuit As String)
    On Error Resume Next
    '
    If IRCSocket.State = sckConnected Then
        IRCSocket.SendData "QUIT :" & sQuit & vbCrLf
        '
        Set tmrCloseSock = New CLiteTimer
        tmrCloseSock.Interval = 20
        tmrCloseSock.Enabled = True
    ElseIf IRCSocket.State = sckConnecting Or IRCSocket.State = sckResolvingHost Then
        blClose = True
        RaiseEvent IRCSocketCancel
        IRCSocket.CloseSocket
    End If
End Sub

'socket senddata routine
Public Sub IRCSocketSendData(ByVal sData As String)
    On Error Resume Next
    '
    If IRCSocket.State = sckConnected Then IRCSocket.SendData sData
    '
End Sub

'local socket properties
Public Property Get IRCSocketState() As StateConstants
    On Error Resume Next
    '
    IRCSocketState = IRCSocket.State
    '
End Property

Public Property Get IRCSocketLocalHostName() As String
    On Error Resume Next
    '
    IRCSocketLocalHostName = IRCSocket.LocalHostName
    '
End Property

Public Property Get IRCSocketRemoteHost() As String
    On Error Resume Next
    '
    IRCSocketRemoteHost = IRCSocket.RemoteHost
    '
End Property

Public Property Get IRCSocketRemoteHostIP() As String
    On Error Resume Next
    '
    IRCSocketRemoteHostIP = IRCSocket.RemoteHostIP
    '
End Property

Public Property Get IRCSocketLocalIP() As String
    On Error Resume Next
    '
    IRCSocketLocalIP = IRCSocket.LocalIP
    '
End Property

Public Property Get IRCSocketRemotePort() As Long
    On Error Resume Next
    '
    IRCSocketRemotePort = IRCSocket.RemotePort
    '
End Property

Public Property Let IRCSocketRemotePort(sNewValue As Long)
    On Error Resume Next
    '
    IRCSocket.RemotePort = sNewValue
    '
End Property

Public Property Let IRCSocketRemoteHost(sNewValue As String)
    On Error Resume Next
    '
    IRCSocket.RemoteHost = sNewValue
    '
End Property

Public Property Get IRCSocketProtocol() As ProtocolConstants
    On Error Resume Next
    '
    IRCSocketProtocol = IRCSocket.Protocol
    '
End Property

Public Property Let IRCSocketProtocol(sNewValue As ProtocolConstants)
    On Error Resume Next
    '
    IRCSocket.Protocol = sNewValue
    '
End Property

Public Property Get IRCSocketLocalPort() As Long
    On Error Resume Next
    '
    IRCSocketLocalPort = IRCSocket.LocalPort
    '
End Property

Public Property Let IRCSocketLocalPort(sNewValue As Long)
    On Error Resume Next
    '
    IRCSocket.LocalPort = sNewValue
    '
End Property

'socket events
Private Sub IRCSocket_OnClose()
    On Error Resume Next
    '
    If blClose = True Then Exit Sub
    '
    Set tmrRaiseDiscon = New CLiteTimer
    tmrRaiseDiscon.Interval = 100
    tmrRaiseDiscon.Enabled = True
End Sub

Private Sub IRCSocket_OnConnect()
    On Error Resume Next
    '
    RaiseEvent IRCSocketConnect
End Sub

Private Sub IRCSocket_OnDataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim sData As String
    '
    IRCSocket.GetData sData
    '
    ParseRAWSocketData sData
End Sub

Private Sub IRCSocket_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    '
    RaiseEvent IRCSocketServerDisconnect(Number, Description)
End Sub

'parsing routines
Public Sub ParseRAWSocketData(ByVal sData As String)
    On Error Resume Next
    '
    Dim M As String * 1
    Dim i As Long
    Dim recvBuf As String
    '
    Do While Len(sData) > 0
        i = InStr(1, sData, vbCrLf)
        '
        If i = 0 Then
            recvBuf = recvBuf & sData
            Exit Sub
        Else
            recvBuf = recvBuf & Left$(sData, i - 1)
            '
            cQueueMSG.Add recvBuf
            '
            recvBuf = vbNullString
            If Len(sData) > i + 1 Then
                sData = Mid$(sData, i + 2)
            Else
                sData = vbNullString
            End If
        End If
        '
    Loop
End Sub

Public Sub Parse(ByVal sFirst As String, sSecond As String, sThird As String, sFourth As String)
    On Error Resume Next
    Dim intTemp As Integer
    Dim intTemp2 As Integer
    Dim strTemp As String
    Dim strTemp2 As String
    Dim strAddress As String
    Dim strCTCP As String
    Dim i As Long
    Select Case UCase(sSecond)
        Case "ERROR"
            intTemp = InStr(sFourth, Chr$(32))
            If intTemp = 0 Then
                'no space, which is unusal
                strTemp = Replace(sFourth, ":", vbNullString)
                strTemp2 = vbNullString
            Else
                strTemp = Left$(sFourth, intTemp - 1)
                strTemp2 = Mid$(sFourth, intTemp + 1)
            End If
            '
            RaiseEvent IRCCloseLink(strTemp, strTemp2)
            Exit Sub
        Case "AUTH"
            If UCase(sFirst) = "NOTICE" Then RaiseEvent IRCSNotice(Replace(sFirst, ":", vbNullString), Mid$(sThird, 2) & " " & sFourth)
            Exit Sub
        Case "PING"
            RaiseEvent IRCPing(sFirst)
            Exit Sub
        Case "PRIVMSG"
            intTemp = InStr(sFirst, "!")
            If intTemp = 0 Then
                'must be a server
                strTemp = Mid$(sFirst, 2)
            Else
                strTemp = Replace(Left$(sFirst, intTemp - 1), ":", vbNullString)
                strAddress = Mid$(sFirst, intTemp + 1)
            End If
            '
            'check for CTCP's
            If Left$(GetTok(sFourth, "1", 32), 2) = ":" & Chr$(1) And UCase(GetTok(sFourth, "1", 32)) <> ":" & Chr$(1) & "ACTION" Then
                strCTCP = Replace(GetTok(sFourth, "1", 32), Chr$(1), vbNullString)
                strCTCP = Replace(strCTCP, ":", vbNullString)
                Select Case UCase(strCTCP)
                    Case "PING"
                        RaiseEvent IRCCTCPPing(strTemp, strAddress, Replace(GetTok(sFourth, "2-", 32), Chr$(1), vbNullString))
                        Exit Sub
                    Case "VERSION"
                        RaiseEvent IRCCTCPVersion(strTemp, strAddress)
                        Exit Sub
                    Case "TIME"
                        RaiseEvent IRCCTCPTime(strTemp, strAddress)
                        Exit Sub
                    Case "FINGER"
                        RaiseEvent IRCCTCPFinger(strTemp, strAddress)
                        Exit Sub
                    Case "AWAY"
                        RaiseEvent IRCCTCPAway(strTemp, strAddress, Replace(GetTok(sFourth, "2-", 32), Chr$(1), vbNullString))
                        Exit Sub
                    Case "DCC"
                        Select Case UCase(GetTok(sFourth, "2", 32))
                            Case "SEND"
                                If Val(GetTok(sFourth, "0", 32)) > 6 Then Exit Sub
                                RaiseEvent IRCDCCSend(strTemp, strAddress, GetTok(sFourth, "3", 32), GetTok(sFourth, "4", 32), GetTok(sFourth, "5", 32), Replace(GetTok(sFourth, "6", 32), Chr$(1), vbNullString))
                                Exit Sub
                            Case "ACCEPT"
                                If Val(GetTok(sFourth, "0", 32)) > 6 Then Exit Sub
                                RaiseEvent IRCDCCAccept(strTemp, strAddress, GetTok(sFourth, "3", 32), GetTok(sFourth, "4", 32), Replace(GetTok(sFourth, "5", 32), Chr$(1), vbNullString))
                                Exit Sub
                            Case "RESUME"
                                RaiseEvent IRCDCCResume(strTemp, strAddress, GetTok(sFourth, "3", 32), GetTok(sFourth, "4", 32), Replace(GetTok(sFourth, "5", 32), Chr$(1), vbNullString))
                                Exit Sub
                            Case "CHAT"
                                RaiseEvent IRCDCCChat(strTemp, strAddress, GetTok(sFourth, "4", 32), Replace(GetTok(sFourth, "5", 32), Chr$(1), vbNullString))
                                Exit Sub
                        End Select
                        Exit Sub
                    Case "OTHER"
                        RaiseEvent IRCCTCPOther(strTemp, strAddress, Replace(GetTok(sFourth, "2-", 32), Chr$(1), vbNullString))
                        Exit Sub
                End Select
            Else
                'it isn't a ctcp so its either a private or
                'a channel message
                If LCase(sThird) = LCase(strMe) Then
                    'its a private message
                    If GetTok(sFourth, "1", 32) = ":" & Chr$(1) & "ACTION" Then
                        'its an action
                        RaiseEvent IRCActionQuery(strTemp, strAddress, Replace(GetTok(sFourth, "2-", 32), Chr$(1), vbNullString))
                        Exit Sub
                    Else
                        'normal message
                        RaiseEvent IRCTextQuery(strTemp, strAddress, Mid$(sFourth, 2))
                        Exit Sub
                    End If
                Else
                    'its a channel message
                    If GetTok(sFourth, "1", 32) = ":" & Chr$(1) & "ACTION" Then
                        'its an action
                        RaiseEvent IRCActionChan(strTemp, strAddress, sThird, Replace(GetTok(sFourth, "2-", 32), Chr$(1), vbNullString))
                        Exit Sub
                    Else
                        'normal message
                        RaiseEvent IRCTextChan(strTemp, strAddress, sThird, Mid$(sFourth, 2))
                        Exit Sub
                    End If
                End If
            End If
            Exit Sub
        Case "NOTICE"
            intTemp = InStr(sFirst, "!")
            If intTemp = 0 Then
                'must be a server
                strTemp = Mid$(sFirst, 2)
            Else
                strTemp = Replace(Left$(sFirst, intTemp - 1), ":", vbNullString)
                strAddress = Mid$(sFirst, intTemp + 1)
            End If
            'we have to first parse CTCP replies
            If Left$(sFourth, 2) = ":" & Chr$(1) Then
                strCTCP = UCase(Replace(Mid$(GetTok(sFourth, "1", 32), 2), Chr$(1), vbNullString))
                Select Case strCTCP
                    Case "VERSION", "TIME", "FINGER"
                        RaiseEvent IRCCTCPReply(strTemp, strCTCP, Replace(GetTok(sFourth, "2-", 32), Chr$(1), vbNullString))
                        Exit Sub
                    Case "PING"
                        i = Val(Replace(GetTok(sFourth, "2-", 32), Chr$(1), vbNullString))
                        If i <> 0 Then
                            RaiseEvent IRCCTCPReply(strTemp, "PING", GetDuration(Round((GetTickCount - i) / 1000)))
                            Exit Sub
                        Else
                            RaiseEvent IRCCTCPReply(strTemp, "PING", Replace(GetTok(sFourth, "2-", 32), Chr$(1), vbNullString))
                            Exit Sub
                        End If
                End Select
                Exit Sub
            Else
                'normal notice
                If UCase(sThird) <> "AUTH" Then
                    RaiseEvent IRCNotice(strTemp, strAddress, Mid$(sFourth, 2))
                    Exit Sub
                Else
                    RaiseEvent IRCSNotice(strTemp, Mid$(sFourth, 2))
                    Exit Sub
                End If
            End If
        Case "JOIN"
            intTemp = InStr(sFirst, "!")
            RaiseEvent IRCJoin(Replace(sFirst, ":", vbNullString), Replace(Left$(sFirst, intTemp - 1), ":", vbNullString), Mid$(sFirst, intTemp + 1), Replace(sThird, ":", vbNullString))
            Exit Sub
        Case "PART"
            intTemp = InStr(sFirst, "!")
            RaiseEvent IRCPart(Replace(Left$(sFirst, intTemp - 1), ":", vbNullString), Mid$(sFirst, intTemp + 1), Replace(sThird, ":", vbNullString), Mid$(sFourth, 2))
            Exit Sub
        Case "KICK"
            intTemp = InStr(sFirst, "!")
            intTemp2 = InStr(sFourth, Chr$(32))
            RaiseEvent IRCKick(Replace(Left$(sFirst, intTemp - 1), ":", vbNullString), Left$(sFourth, intTemp2 - 1), sThird, Mid$(sFourth, intTemp2 + 2))
            Exit Sub
        Case "NICK"
            intTemp = InStr(sFirst, "!")
            RaiseEvent IRCNick(Replace(Left$(sFirst, intTemp - 1), ":", vbNullString), Mid$(sThird, 2))
            Exit Sub
        Case "QUIT"
            intTemp = InStr(sFirst, "!")
            RaiseEvent IRCQuit(Replace(Left$(sFirst, intTemp - 1), ":", vbNullString), Mid$(sFirst, intTemp + 1), sFourth)
            Exit Sub
        Case "INVITE"
            intTemp = InStr(sFirst, "!")
            RaiseEvent IRCInvite(Replace(Left$(sFirst, intTemp - 1), ":", vbNullString), Mid$(sFirst, intTemp + 1), Mid$(sFourth, 2))
            Exit Sub
        Case "TOPIC"
            intTemp = InStr(sFirst, "!")
            RaiseEvent IRCTopic(Replace(Left$(sFirst, intTemp - 1), ":", vbNullString), sThird, Mid$(sFourth, 2))
            Exit Sub
        Case "MODE"
            intTemp = InStr(sFirst, "!")
            If LCase(Mid$(sFirst, 2)) = LCase(sThird) Then
                'self mode
                RaiseEvent IRCSelfMode(sThird, Mid$(sFourth, 2))
                Exit Sub
            Else
                'channel mode
                'get the nick
                intTemp = InStr(sFirst, "!")
                If intTemp = 0 Then
                    'must be a server, not a nick
                    intTemp2 = InStr(sFourth, Chr$(32))
                    RaiseEvent IRCMode(Mid$(sFirst, 2), sThird, Left$(sFourth, intTemp2 - 1), Mid$(sFourth, intTemp2 + 1))
                    Exit Sub
                Else
                    'it is a nick
                    intTemp2 = InStr(sFourth, Chr$(32))
                    RaiseEvent IRCMode(Replace(Left$(sFirst, intTemp - 1), ":", vbNullString), sThird, Left$(sFourth, intTemp2 - 1), Mid$(sFourth, intTemp2 + 1))
                    Exit Sub
                End If
            End If
        'RAWS
        'welcome
        Case "001"
            RaiseEvent IRCWelcome1(Mid$(sFirst, 2), Mid$(sFourth, 2))
            Exit Sub
        Case "002"
            RaiseEvent IRCWelcome2(Mid$(sFirst, 2), Mid$(sFourth, 2))
            Exit Sub
        Case "003"
            RaiseEvent IRCWelcome2(Mid$(sFirst, 2), Mid$(sFourth, 2))
            Exit Sub
        Case "004"
            RaiseEvent IRCWelcome2(Mid$(sFirst, 2), sFourth)
            Exit Sub
        'protocols
        Case "005"
            RaiseEvent IRCProtocols(Mid$(sFirst, 2), Replace(sFourth, " :are available on this server", vbNullString))
            Exit Sub
        'lusers
        Case "251", "252", "253", "254", "255", "265", "266"
            RaiseEvent IRCLusers(Mid$(sFirst, 2), Replace(sFourth, ":", vbNullString))
            Exit Sub
        'topic on join
        Case "332"
            intTemp = InStr(sFourth, Chr$(32))
            RaiseEvent IRCTopicIs(Left$(sFourth, intTemp - 1), Mid$(sFourth, intTemp + 2))
            Exit Sub
        'topic set
        Case "333"
            RaiseEvent IRCTopicSet(GetTok(sFourth, "1", 32), GetTok(sFourth, "2", 32), GetTok(sFourth, "3", 32))
            Exit Sub
        'who
        Case "352"
            RaiseEvent IRCWho(sThird, GetTok(sFourth, "1", 32), GetTok(sFourth, "5", 32) & "!" & GetTok(sFourth, "2", 32) & "@" & GetTok(sFourth, "3", 32))
            Exit Sub
        'who end
        Case "315"
            RaiseEvent IRCWhoEnd(sThird, GetTok(sFourth, "1", 32))
            Exit Sub
        'channel list
        Case "322"
            RaiseEvent IRCChannelList(Replace(sFirst, ":", vbNullString), GetTok(sFourth, "1", 32), GetTok(sFourth, "2", 32), Mid$(GetTok(sFourth, "3-", 32), 2))
            Exit Sub
        'end channel list
        Case "323"
            RaiseEvent IRCChannelListEnd(Replace(sFirst, ":", vbNullString), "End of /LIST command.")
            Exit Sub
        'channel modes
        Case "324"
            intTemp = InStr(sFourth, Chr$(32))
            RaiseEvent IRCChannelModes(Left$(sFourth, intTemp - 1), Mid$(sFourth, intTemp + 1))
            Exit Sub
        'names
        Case "353"
            intTemp = InStr(sFourth, Chr$(32))
            intTemp2 = InStr(intTemp + 1, sFourth, Chr$(32))
            RaiseEvent IRCNames(Trim(Mid$(sFourth, intTemp + 1, intTemp2 - intTemp)), Replace(Trim(Mid$(sFourth, intTemp2 + 1)), ":", vbNullString))
            Exit Sub
        Case "366"
            intTemp = InStr(sFourth, Chr$(32))
            RaiseEvent IRCNamesEnd(Left$(sFourth, intTemp - 1))
            Exit Sub
        'ban list
        Case "367"
            intTemp = InStr(sFourth, Chr$(32))
            RaiseEvent IRCBanList(Left$(sFourth, intTemp - 1), GetTok(sFourth, "2", 32), GetTok(sFourth, "3", 32), GetTok(sFourth, "4", 32))
            Exit Sub
        'end of ban list
        Case "368"
            intTemp = InStr(sFourth, Chr$(32))
            RaiseEvent IRCBanListEnd(Left$(sFourth, intTemp - 1), "End of BAN list.")
            Exit Sub
        'unknown
        Case "421"
            intTemp = InStr(sFourth, Chr$(32))
            Select Case UCase(Left$(sFourth, intTemp - 1))
                Case "IRCX": Exit Sub
                Case Else
                    RaiseEvent IRCUnknown(Mid$(sFirst, 2), Mid$(sFourth, intTemp + 2) & ": " & Left$(sFourth, intTemp - 1))
                    Exit Sub
            End Select
        'motd
        Case "372"
            RaiseEvent IRCMOTDText(Mid$(sFirst, 2), Mid$(sFourth, 2))
            Exit Sub
        Case "375"
            RaiseEvent IRCMOTDStart(Mid$(sFirst, 2), Mid$(sFourth, 2))
            Exit Sub
        Case "376"
            RaiseEvent IRCMOTDEnd(Mid$(sFirst, 2), "End of Message Of The Day (MOTD)")
            Exit Sub
        Case "377", "378"
            RaiseEvent IRCMOTDText(Mid$(sFirst, 2), Mid$(sFourth, 2))
            Exit Sub
        'motd error
        Case "422"
            RaiseEvent IRCMOTDError(Mid$(sFirst, 2), Mid$(sFourth, 2))
            Exit Sub
        'cannot join
        Case "471"
            intTemp = InStr(sFourth, Chr$(32))
            RaiseEvent IRCCannotJoin(Left$(sFourth, intTemp - 1), "+l")
            Exit Sub
        Case "473"
            intTemp = InStr(sFourth, Chr$(32))
            RaiseEvent IRCCannotJoin(Left$(sFourth, intTemp - 1), "+i")
            Exit Sub
        Case "474"
            intTemp = InStr(sFourth, Chr$(32))
            RaiseEvent IRCCannotJoin(Left$(sFourth, intTemp - 1), "+b")
            Exit Sub
        Case "475"
            intTemp = InStr(sFourth, Chr$(32))
            RaiseEvent IRCCannotJoin(Left$(sFourth, intTemp - 1), "+k")
            Exit Sub
        Case "477"
            intTemp = InStr(sFourth, Chr$(32))
            RaiseEvent IRCCannotJoin(Left$(sFourth, intTemp - 1), "+r")
            Exit Sub
        'watch
        Case "600"
            RaiseEvent IRCNotify(Mid$(sFirst, 2), GetTok(sFourth, "1", 32), GetTok(sFourth, "2", 32) & "@" & GetTok(sFourth, "3", 32))
            Exit Sub
        Case "601"
            RaiseEvent IRCUNotify(Mid$(sFirst, 2), GetTok(sFourth, "1", 32), GetTok(sFourth, "2", 32) & "@" & GetTok(sFourth, "3", 32))
            Exit Sub
        Case "604"
            RaiseEvent IRCNotifyUpdate(Mid$(sFirst, 2), GetTok(sFourth, "1", 32), GetTok(sFourth, "2", 32) & "@" & GetTok(sFourth, "3", 32))
            Exit Sub
        'ircx
        Case "800"
            RaiseEvent IRCRawOther(Mid$(sFirst, 2), sSecond, sFourth)
            Exit Sub
        Case Else
            RaiseEvent IRCRawOther(Mid$(sFirst, 2), sSecond, IIf(Left$(sFourth, 1) = ":", Mid$(sFourth, 2), sFourth))
            Exit Sub
    End Select
End Sub

'class properties
Private Sub Class_Initialize()
    On Error Resume Next
    Set IRCSocket = New CSocket
    IRCSocketProtocol = sckTCPProtocol
    '
    Set tmrParse = New CLiteTimer
    tmrParse.Interval = 1
    tmrParse.Enabled = True
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    tmrParse.Enabled = False
    Set tmrParse = Nothing
    Set tmrCloseSock = Nothing
    Set tmrRaiseDiscon = Nothing
    '
    Set IRCSocket = Nothing
End Sub

'timer class
Private Sub tmrCloseSock_Timer()
    On Error Resume Next
    tmrCloseSock.Enabled = False
    Set tmrCloseSock = Nothing
    IRCSocket.CloseSocket
End Sub

Private Sub tmrRaiseDiscon_Timer()
    On Error Resume Next
    tmrRaiseDiscon.Enabled = False
    Set tmrRaiseDiscon = Nothing
    RaiseEvent IRCSocketDisconnect
End Sub

Private Sub tmrParse_Timer()
    On Error Resume Next
    'The timer will check for queue message every 1 millisecond.  This is secondary parse
    Dim intCount As Integer
    Dim blnParsed As Boolean
    Dim strData As String
    '
    Dim strFirst As String
    Dim strSecond As String
    Dim strThird As String
    Dim strFourth As String
    '
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim intPos3 As Integer
    Dim intPos4 As Integer
    '
    intCount = 1
    Do While blnParsed = False And intCount <= cQueueMSG.Count
        strData = cQueueMSG.Item(intCount)
        'remove first line feed if there are any
        If Mid$(strData, 1, 1) = Chr$(13) Or Mid$(strData, 1, 1) = Chr$(10) Then
            strData = Mid$(strData, 2)
        End If
        intPos1 = InStr(1, strData, Chr$(32))
        If intPos1 Then
            strFirst = Trim$(Left$(strData, intPos1))
            intPos2 = InStr(intPos1 + 1, strData, Chr$(32))
            If intPos2 Then
                strSecond = Trim$(Mid$(strData, intPos1 + 1, (intPos2 - intPos1)))
                intPos3 = InStr(intPos2 + 1, strData, Chr$(32))
                    If intPos3 Then
                        strThird = Trim$(Mid$(strData, intPos2 + 1, (intPos3 - intPos2)))
                        strFourth = Trim$(Right$(strData, Len(strData) - intPos3))
                    Else
                        'no third space
                        strThird = Trim$(Mid$(strData, intPos2 + 1, Len(strData) - intPos2))
                    End If
            Else
                'no second space, most likely PING or ERROR
                strFirst = Trim$(Right$(strData, Len(strData) - InStr(strData, ":")))
                strSecond = "PING"
                strThird = vbNullString
                strFourth = vbNullString
            End If
        End If
        'Error case
        If UCase(strFirst) = "ERROR" Then
            strFirst = vbNullString
            strSecond = "ERROR"
        End If
        blnParsed = True
        cQueueMSG.Remove intCount
        intCount = intCount + 1
        '
        RaiseEvent SocketDataOutput(strFirst, strSecond, strThird, strFourth)
        '
    Loop
End Sub
