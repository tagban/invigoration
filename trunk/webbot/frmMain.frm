VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H001D1D1D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "InvigWebbot"
   ClientHeight    =   4530
   ClientLeft      =   180
   ClientTop       =   930
   ClientWidth     =   5385
   ClipControls    =   0   'False
   FillColor       =   &H0080FFFF&
   FillStyle       =   3  'Vertical Line
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00008000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "None"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   Picture         =   "frmMain.frx":164A
   ScaleHeight     =   4530
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSend 
      Interval        =   600
      Left            =   2160
      Top             =   960
   End
   Begin InetCtlsObjects.Inet webbot 
      Left            =   840
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.Timer tmrConnect 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   240
   End
   Begin VB.Timer ConnectTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3120
      Top             =   240
   End
   Begin VB.Timer tmrDC 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2520
      Top             =   240
   End
   Begin MSWinsockLib.Winsock wsBnls 
      Left            =   1440
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsBnet 
      Left            =   960
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   7646
      _Version        =   393217
      BackColor       =   2368548
      Enabled         =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":2931
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuMail 
      Caption         =   "&Bot"
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect to Battle.net.."
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect from Battle.net"
      End
      Begin VB.Menu Reconnect 
         Caption         =   "Reconnect to Battle.net..."
      End
      Begin VB.Menu sep123 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuConfigure 
      Caption         =   "&Configure"
      Begin VB.Menu mnuSetupOption 
         Caption         =   "Edit Configuration"
         Shortcut        =   {F6}
      End
      Begin VB.Menu AA 
         Caption         =   "-"
      End
      Begin VB.Menu ToTray 
         Caption         =   "Minimize to System Tray"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuClearBufs 
         Caption         =   "Clear Chat Buffers"
         Shortcut        =   %{BKSP}
      End
   End
   Begin VB.Menu mnuOther 
      Caption         =   "&Other"
      Begin VB.Menu mnuAbout 
         Caption         =   "About.."
      End
      Begin VB.Menu mnuWebsite 
         Caption         =   "Website"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public antiidlesecond As Integer
Public newest As String
Public connectstatus As Boolean
Public winampnow As String
Public sUserName As String
Public connectseconds As Long
Public random As Integer
Public dctime As Integer
Public IdleTime As Integer
Public privatever As Boolean
Public intRandom As Integer
Public hexchat As Integer
Public webbotsite As String
Private Const WM_USER = &H400&
Private Const EM_AUTOURLDETECT = (WM_USER + 91)
Public WithEvents ChatBot As BnetBot
Attribute ChatBot.VB_VarHelpID = -1

Public Sub PrepareCheck(ByRef tocheck As String)
    tocheck = Replace(tocheck, "[", "a")
    tocheck = Replace(tocheck, "]", "x")
    tocheck = Replace(tocheck, "#", "y")
    tocheck = Replace(tocheck, "-", "z")
    tocheck = Replace(tocheck, "&", "b")
End Sub

Private Sub Command1_Click()
    ShellExecute Me.hWnd, "Open", "http://www.bnet.cc/", 0&, 0&, 0&
End Sub


Private Sub Form_Load()
    Set ChatBot = New BnetBot
    Dim lRet As Long
    frmMain.Caption = "BNET.cc - Webbot"
    LogChat = 0
    LoadConfig
    webbotsite = "http://webbot.bnetweb.org/webbot.php?u=" & BNET.WebUser & "&p=" & BNET.WebPass
    connectstatus = False
    AddChat2 D2Green, "BNET.cc  Webbot 1.0 Based on BNETWeb's API"
    AddChat2 D2Green, "Webbot Data: " & GoWinInet(webbotsite & "&f=99")
    AddChat2 D2Green, "Setting Webbot Server to: " & BNET.BattlenetServer & " :: " & GoWinInet(webbotsite & "&f=10s=" & BNET.BattlenetServer)
    AddChat D2MedBlue, "BNET.cc Webbot Activated" & GoWinInet(webbotsite & "&f=15 v=BNET.cc+Webbot+v1.0")
    GoWinInet (webbotsite & "&f=11&c=clr+all")
    AddChat D2MedBlue, "---------------------------------------------------"
    
    
    frmConfigBNET.txtCDKey.text = GetStuff("BNET", "CDKey")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveConfig
    frmMain.wsBnet.Close
    frmMain.wsBnls.Close
End
End Sub

Private Sub Form_Terminate()
    SaveConfig
    frmMain.wsBnet.Close
    frmMain.wsBnls.Close
    End
End Sub


Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub
Private Sub mnuBug_Click()
    ShellExecute Me.hWnd, "Open", "http://code.google.com/p/BNET.cc/issues/entry", 0&, 0&, 0&
    AddChat D2White, "Thank you for contributing, we appreciate all of your testing! If a window didn't open, go here:  ohttp://code.google.com/p/BNET.cc/issues/entry "
End Sub

Private Sub mnuClearBufs_Click()
    rtbChat.text = vbNullString
    AddChat D2White, "Cleared Chat Buffers. Old information is lost."
End Sub
Private Sub mnuConnect_Click()
On Error GoTo Error
connectstatus = True
    tmrConnect.Enabled = True
    AddChat D2Green, "Battle.net Login Server Connecting to " & BNET.BNLSServer & "..."
    frmMain.wsBnls.Close
    frmMain.wsBnls.Connect BNET.BNLSServer, 9367
Error:

End Sub

Private Sub tmrConnect_Timer()
    connectseconds = connectseconds + 1
End Sub


Private Sub mnuDisconnect_Click()
    connectstatus = False
    frmMain.Caption = "BNET.ccWebbot - http://www.bnet.cc/"
    frmMain.wsBnet.Close
    frmMain.wsBnls.Close
    AddChat D2White, "Battle.net Closed Connection."
    frmMain.Caption = "BNET.cc - [ Disconnected ]"
    frmMain.tmrDC.Enabled = False
    mnuConnect.Visible = True
    mnuDisconnect.Visible = False
End Sub

Private Sub mnuExit_Click()
    frmMain.wsBnet.Close
    frmMain.wsBnls.Close
    Unload Me
    End
End Sub
Private Sub mnuSetupOption_Click()
    frmConfigBNET.Show
End Sub

Private Sub mnuWebsite_Click()
    ShellExecute Me.hWnd, "Open", "http://code.google.com/p/BNET.cc/", 0&, 0&, 0&
End Sub

Private Sub Reconnect_Click()
On Error GoTo Error
    frmMain.tmrDC.Enabled = False
    frmMain.wsBnet.Close
    frmMain.wsBnls.Close
    connectstatus = False
    frmMain.Caption = "BNET.cc - [ Reconnecting to: " & BNET.BNLSServer & " ]"
    AddChat D2Green, "Battle.net Login Server Reconnecting to " & BNET.BNLSServer & "..."
    frmMain.wsBnls.Close
    frmMain.wsBnls.Connect BNET.BNLSServer, 9367
    connectstatus = True
Error:

End Sub


Private Sub rtbChat_Change()
 If Len(rtbChat.text) >= 8000 Then
        rtbChat.Visible = False
        rtbChat.SelStart = 0
        removed = InStr(1, rtbChat.text, vbCrLf) + 1
        rtbChat.SelLength = removed
        rtbChat.SelText = vbNullString
        rtbChat.SelStart = 0
        removed = InStr(1, rtbChat.text, vbCrLf) + 1
        rtbChat.SelLength = removed
        rtbChat.SelText = vbNullString
        rtbChat.SelStart = 0
        removed = InStr(1, rtbChat.text, vbCrLf) + 1
        rtbChat.SelLength = removed
        rtbChat.SelText = vbNullString
        rtbChat.SelStart = Len(rtbChat.text)
        rtbChat.Visible = True
    End If
End Sub

Private Sub tmrDC_Timer()
dctime = dctime + 1
connectstatus = False
tmrAntiIdle.Enabled = False
     If dctime = 1 Then
        AddChat D2Green, "Reconnecting in 4 minutes"
    ElseIf dctime = 2 Then
        AddChat D2Green, "Reconnecting in 3 minutes"
    ElseIf dctime = 3 Then
        AddChat D2Green, "Reconnecting in 2 minutes"
    ElseIf dctime = 4 Then
        AddChat D2Green, "Reconnecting in 1 minutes"
    ElseIf dctime = 5 Then
        frmMain.tmrDC.Enabled = False
        AddChat D2Green, "Reconnecting..."
        dctime = 0
        AddChat D2Green, "Battle.net Login Server Connecting to " & BNET.BNLSServer & "..."
        frmMain.wsBnls.Close
        frmMain.wsBnls.Connect BNET.BNLSServer, 9367
        frmMain.Caption = "BNET.cc - [ Connecting to: " & BNET.BNLSServer & " ]"
        connectstatus = True
    End If
End Sub

Private Sub ToTray_Click()
TrayToolTip = "BNET.ccWebbot [ " & BNET.username & " ]" & vbNewLine & "Version : " & vernum
       Me.WindowState = vbMinimized
        Me.Hide
        With nID
            .cbSize = Len(nID)
            .hWnd = Me.hWnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_Message
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon      '
            .szTip = TrayToolTip      'tooltip text
        End With
    Shell_NotifyIcon NIM_ADD, nID
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Result, Action As Long
    If Me.ScaleMode = vbPixels Then
        Action = X
    Else
        Action = X / Screen.TwipsPerPixelX
    End If
    
Select Case Action

    Case WM_LBUTTONDBLCLK
        Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hWnd)
        Me.Show
        Shell_NotifyIcon NIM_DELETE, nID
    
    Case WM_RBUTTONUP
        Result = SetForegroundWindow(Me.hWnd)
    End Select
End Sub
Private Sub Update_Click()
    ShellExecute Me.hWnd, "Open", "http://www.bnet.cc/", 0&, 0&, 0&
End Sub
Private Sub wsBnet_Close()
    mnuConnect.Visible = True
    mnuDisconnect.Visible = False
    AddChat D2Red, "Battle.net Disconnected"
    frmMain.Caption = "BNET.ccWebbot - [ Disconnected ]"
    lstChannel.ListItems.Clear
End Sub

Private Sub wsBnet_Connect()
    mnuConnect.Visible = False
    mnuDisconnect.Visible = True
    AddChat D2Green, "Battle.net Connected!"
    wsBnet.SendData Chr(1)
    Send0x50
End Sub

Private Sub wsBnet_DataArrival(ByVal bytesTotal As Long)
Static strBuffer As String
Dim strTemp As String, lngLen As Long
    wsBnet.GetData strTemp, vbString
    strBuffer = strBuffer & strTemp
    While Len(strBuffer) > 4
        lngLen = Val("&H" & StrToHex(StrReverse(Mid(strBuffer, 3, 2))))
        If Len(strBuffer) < lngLen Then: Exit Sub
        ParseBnet (Left(strBuffer, lngLen))
        strBuffer = Mid(strBuffer, lngLen + 1)
    Wend
End Sub

Private Sub wsBnls_Close()
    AddChat D2Red, "Battle.net Login Server Connection Closed."
    frmMain.Caption = "BNET.ccWebbot - [BNLS Connection CLOSED]"
End Sub

Private Sub wsBnls_Connect()
    AddChat D2Green, "Battle.net Login Server Connected!"
    frmMain.Caption = "BNET.ccWebbot - [BNLS Connected to " & wsBnls.RemoteHostIP & "]"
    With PBuffer
        .InsertNTString "BNET.cc"
        .SendBNLSPacket &HE
    End With
End Sub

Private Sub wsBnls_DataArrival(ByVal bytesTotal As Long)
Dim TempData As String
    wsBnls.GetData TempData, vbString
    ParseBNLS TempData
End Sub

Private Sub ChatBot_OnUser(ByVal username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
Dim Product As String ', thing As New BnetBot
Product = Mid(Message, 1, 4)
Product = StrReverse(Product)
AddChat D2White, Flags
    Flags = 0
    GoWinInet (webbotsite & "&f=4&d=" & Format(Time, "HH:MM") & "&n=" & username & "&cl=" & Product & "&pi=" & Ping & "&fl=" & Flags)
    AddChat D2White, "Product: " & Product & " Flags: " & Flags & " Ping: " & Ping & " Username: " & username
End Sub


Private Sub ChatBot_OnChannel(ByVal ChannelName As String, ByVal Flags As Long)
    AddChat D2Orange, "Joining (" & GetChannelType(Flags) & ") " & ChannelName & " w/Flags: " & Flags & "."
    GoWinInet (webbotsite & "&f=2&d=" & Format(Time, "HH:MM") & "&ch=" & ChannelName)
End Sub

Private Sub ChatBot_OnEmote(ByVal username As String, ByVal Flags As Long, ByVal Message As String)
    AddChat D2Beige1, "* " & username & Space(1) & Message & " *"
    GoWinInet (webbotsite & "&f=9&d=" & Format(Time, "HH:MM") & "&n=" & username & "&t=" & Message)
End Sub
Private Sub ChatBot_OnInfo(ByVal Message As String)
        AddChat vbYellow, Message
End Sub

Private Sub ChatBot_OnFlags(ByVal username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
Dim thing As New BnetBot
    GoWinInet (webbotsite & "&f=12&n=" & username)
End Sub



Private Sub ChatBot_OnJoin(ByVal username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
Dim Product As String ', thing As New BnetBot
Product = Mid(Message, 1, 4)
Product = StrReverse(Product)
AddChat D2White, Flags
    Flags = 0
    GoWinInet (webbotsite & "&f=4&d=" & Format(Time, "HH:MM") & "&n=" & username & "&cl=" & Product & "&pi=" & Ping & "&fl=" & Flags)
    AddChat D2White, "Product: " & Product & " Flags: " & Flags & " Ping: " & Ping & " Username: " & username

End Sub
Private Sub ChatBot_OnLeave(ByVal username As String, ByVal Flags As Long)
        AddChat D2Green, username & " has left the channel."
        GoWinInet (webbotsite & "&f=16&d=" & Format(Time, "HH:MM") & "&n=" & username)
End Sub

Private Sub ChatBot_OnTalk(ByVal username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
'Message , username 'send commands to parser
AddChat D2White, username & " : " & Message
          GoWinInet (webbotsite & "&f=8&d=" & Format(Time, "HH:MM") & "&n=" & username & "&t=" & Message)
On Error GoTo Error

Error:


End Sub


Private Sub ChatBot_OnUnknown(ByVal UnknownString As String)
    AddChat D2Purple, "UNKNOWN STRING: ", vbRed, UnknownString
End Sub

Private Sub ChatBot_OnWhisperFrom(ByVal username As String, ByVal Flags As Long, ByVal Message As String)

    AddChat D2Beige1, ":: Whisper From: " & username & " :: ", vbGrey, Message
    LastW = username
    'LastCW = "/w " & username & " "
    LastM = Message
End Sub

Private Sub ChatBot_OnWhisperTo(ByVal username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
    AddChat D2Beige1, ":: Whisper To: " & username & " :: ", vbGrey, Message
    LastSW = username
    LastSM = Message
End Sub
Private Sub mnuQC1_Click()
On Error Resume Next

End Sub
Private Sub wsBnet_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error GoTo Err3:
    AddChat D2Red, "Internet Connection Error: Auto-Reconnecting in 5 mins."
    wsBnet.Close
    wsBnls.Close
    frmMain.tmrDC.Enabled = True
Err3:
    
End Sub

Private Sub wsRealm_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddChat D2Red, "Realm Connection Failed..."
End Sub
