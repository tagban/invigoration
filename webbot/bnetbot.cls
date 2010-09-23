VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BnetBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const ID_EMOTE = &H17
Private Const ID_WHISPTO = &HA
Private Const ID_WHISPFROM = &H4
Private Const ID_TALK = &H5
Private Const ID_USER = &H1
Private Const ID_JOIN = &H2
Private Const ID_LEAVE = &H3
Private Const ID_FLAGS = &H9
Private Const ID_CHAN = &H7
Private Const ID_ERROR = &H13
Private Const ID_INFO = &H12
Private Const ID_BROADCAST = &H6
Private Const BNFLAGS_BLIZZ = &H1
Private Const BNFLAGS_OP = &H2
Private Const BNFLAGS_SPKR = &H4
Private Const BNFLAGS_SYSOP = &H8
Private Const BNFLAGS_PLUG = &H10
Private Const BNFLAGS_SQUELCH = &H20
Private Const BNFLAGS_GLASSES = &H40
Private Const BNFLAGS_GFPLAYER = &H200000
Private Const bnu = &H800000
Private Const HACK = &H8000000
Private Const WARCRAFT3 = &H80000000
Private Const pidSEXP = 9
Private Const pidSTAR = 6
Private Const pidD2DV = 5
Private Const pidD2XP = 13
Private Const pidW2BN = 10
Private Const pidJSTR = 11
Private Const pidSSHR = 7
Private Const pidWAR3 = 61
Private Const pidDSHR = 64
Private Const pidCHAT = 1
Private Const pidELSE = 59
Private Const pidDRTL = 3
Private Const pidBLIZZ = 63
Private Const pidOP = 2
Private Const pidSQU = 14
Private Const pidW3XP = 62
'' BNET.cc Icons
'' id: 66 - BNET.cc Icon
'' id: 67 - BNU`Bot Icon
'' id: 68 - HACKER

Public Event OnUser(ByVal username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
Public Event OnJoin(ByVal username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
Public Event OnLeave(ByVal username As String, ByVal Flags As Long)
Public Event OnWhisperFrom(ByVal username As String, ByVal Flags As Long, ByVal Message As String)
Public Event OnTalk(ByVal username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
Public Event OnChannel(ByVal ChannelName As String, ByVal Flags As Long)
Public Event OnFlags(ByVal username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
Public Event OnWhisperTo(ByVal username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
Public Event OnInfo(ByVal Message As String)
Public Event OnError(ByVal Message As String)
Public Event OnEmote(ByVal username As String, ByVal Flags As Long, ByVal Message As String)
Public Event OnUnknown(ByVal UnknownString As String)
Public Sub DispatchMessage(ByVal databuf As String)
Dim EID As Long, Ping As Long, username As String, Message As String
    EID = MakeLong(Mid$(databuf, 5, 4))
    Flags = MakeLong(Mid$(databuf, 9, 4))
    Ping = MakeLong(Mid$(databuf, 13, 4))
    username = KillNull(Mid$(databuf, 29))
    Message = KillNull(Mid$(databuf, Len(username) + 30))
Select Case EID
    Case ID_TALK
        RaiseEvent OnTalk(username, Flags, Message, Ping)
        Exit Sub
    Case ID_EMOTE
        RaiseEvent OnEmote(username, Flags, Message)
        Exit Sub
    Case ID_CHAN
        RaiseEvent OnChannel(Message, Flags)
        Exit Sub
    Case ID_USER
        RaiseEvent OnUser(username, Flags, Message, Ping)
        Exit Sub
    Case ID_JOIN
        RaiseEvent OnJoin(username, Flags, Message, Ping)
        Exit Sub
    Case ID_LEAVE
        RaiseEvent OnLeave(username, Flags)
        Exit Sub
    Case ID_WHISPTO
        RaiseEvent OnWhisperTo(username, Flags, Message, Ping)
        Exit Sub
    Case ID_WHISPFROM
        RaiseEvent OnWhisperFrom(username, Flags, Message)
        Exit Sub
    Case ID_INFO
        RaiseEvent OnInfo(Message)
        Exit Sub
    Case ID_FLAGS
        RaiseEvent OnFlags(username, Flags, Message, Ping)
        Exit Sub
    Case ID_BROADCAST
        RaiseEvent OnInfo(Message)
        Exit Sub
    Case ID_ERROR
        RaiseEvent OnError(Message)
        Exit Sub
    Case Else
        RaiseEvent OnUnknown(Message)
        Exit Sub
End Select
End Sub
Function MakeLong(X As String) As Long
    If Len(X) < 4 Then
        Exit Function
    End If
    CopyMemory MakeLong, ByVal X, 4
End Function
Public Function GetIconCode(ByVal statstring As String, ByVal Flags As Long) As Integer
Dim Product As String
Product = Mid(statstring, 1, 4)

    If (BNFLAGS_BLIZZ And Flags) = BNFLAGS_BLIZZ Then
        GetIconCode = 63
        Exit Function
    ElseIf (BNFLAGS_OP And Flags) = BNFLAGS_OP Then
        GetIconCode = 2
        Exit Function
    ElseIf (BNFLAGS_SYSOP And Flags) = BNFLAGS_SYSOP Then
        GetIconCode = 64
        Exit Function
    ElseIf (BNFLAGS_SQUELCH And Flags) = BNFLAGS_SQUELCH Then
        GetIconCode = 14
        Exit Function
    ElseIf (BNFLAGS_GLASSES And Flags) = BNFLAGS_GLASSES Then
        GetIconCode = 65
        Exit Function
    ElseIf (BNETcc And Flags) = BNETcc Then
        GetIconCode = 66
        Exit Function
    ElseIf (bnu And Flags) = bnu Then
        GetIconCode = 69
        Exit Function
    ElseIf (HACK And Flags) = HACK Then
        GetIconCode = 70
        'GetIconCode = 64
        Exit Function
    ElseIf (WARCRAFT3 And Flags) = WARCRAFT3 Then
        GetIconCode = 61
        Exit Function
    ElseIf Not Flags Then: GoTo Products
    End If
Products:
    
   Select Case Product
        Case "PXES": GetIconCode = pidSEXP
        Case "RATS": GetIconCode = pidSTAR
        Case "RTSJ": GetIconCode = pidJSTR
        Case "RHSS": GetIconCode = pidSSHR
        Case "NB2W": GetIconCode = pidW2BN
        Case "VD2D": GetIconCode = pidD2DV
        Case "PX2D": GetIconCode = pidD2XP
        Case "LTRD": GetIconCode = pidDRTL
        Case "RHSD": GetIconCode = pidDRTL
        Case "TAHC": GetIconCode = pidCHAT
        Case "3RAW": GetIconCode = pidWAR3
        Case "PX3W": GetIconCode = pidW3XP
        Case Else: GetIconCode = pidELSE
    End Select
End Function


Public Property Get DataSource() As DataSource

End Property

Public Property Set DataSource(ByVal objDataSource As DataSource)

End Property

Public Property Get DataMember() As DataMember

End Property

Public Property Let DataMember(ByVal DataMember As DataMember)

End Property

