VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invigoration [Nightly]"
   ClientHeight    =   5730
   ClientLeft      =   240
   ClientTop       =   990
   ClientWidth     =   12405
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
   ScaleHeight     =   5730
   ScaleWidth      =   12405
   Begin RichTextLib.RichTextBox rtbsendbnet 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   -2147483640
      Enabled         =   -1  'True
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
   Begin VB.Timer tmrConnect 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5760
      Top             =   240
   End
   Begin VB.Timer ConnectTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5280
      Top             =   240
   End
   Begin VB.Timer IdleTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   240
   End
   Begin VB.Timer tmrDC 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4320
      Top             =   240
   End
   Begin VB.Timer tmrUptime 
      Interval        =   1000
      Left            =   3840
      Top             =   240
   End
   Begin VB.Timer tmrAntiIdle 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   3360
      Top             =   240
   End
   Begin MSComctlLib.ListView lstChannel 
      Height          =   4440
      Left            =   8520
      TabIndex        =   1
      Top             =   360
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   7832
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   3
      _Version        =   393217
      Icons           =   "ClientIcons"
      SmallIcons      =   "ClientIcons"
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "a"
         Object.Width           =   4235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "s"
         Object.Width           =   1059
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsRealm 
      Left            =   1920
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
   Begin MSComctlLib.ImageList ClientIcons 
      Left            =   8040
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   1907997
      ImageWidth      =   28
      ImageHeight     =   14
      MaskColor       =   1907997
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   71
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":304C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3538
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":48E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4DD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":52C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":57AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":70E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":72BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":77AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C97
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8183
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":866F
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8B5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9047
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9533
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9F0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A3F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A8E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ADCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B2BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B7A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BC93
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C17F
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C66B
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CB57
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D043
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D537
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DA2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DF1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E413
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E907
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EDFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F2EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F7E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FCD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":101CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":106BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10BB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":110A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1159B
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A87
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11F73
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1245F
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1294B
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12E37
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13323
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1380F
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13CFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":141E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":146D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14BBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":150AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15597
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1584F
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15EB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":164D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":168FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17482
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17688
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17930
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17DDE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin InetCtlsObjects.Inet devNEWS 
      Left            =   240
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7646
      _Version        =   393217
      BackColor       =   2368548
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":183F8
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
   Begin VB.Label txtChannelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Invigoration"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   8520
      TabIndex        =   2
      Top             =   120
      Width           =   3135
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
      Begin VB.Menu mnuUpdate 
         Caption         =   "Update"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuWebsite 
         Caption         =   "Website"
      End
      Begin VB.Menu mnuBug 
         Caption         =   "Report Bugs"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnuUserList 
      Caption         =   "UserList"
      Visible         =   0   'False
      Begin VB.Menu mnuWhisper 
         Caption         =   "Whisper"
      End
      Begin VB.Menu mnuSquelch 
         Caption         =   "Squelch"
      End
      Begin VB.Menu mnuUnsquelch 
         Caption         =   "UnSquelch"
      End
      Begin VB.Menu mnuUserFocus 
         Caption         =   "User Focus"
      End
      Begin VB.Menu mnuEndFocus 
         Caption         =   "End Focus"
      End
      Begin VB.Menu mnuOp 
         Caption         =   "--Ops--"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBan 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuKick 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuFList 
         Caption         =   "--Friend List--"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuFRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuFListView 
         Caption         =   "View List"
      End
   End
   Begin VB.Menu mnuhahhah 
      Caption         =   "|"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuWinamp 
      Caption         =   "Win&amp"
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu mnuLoadWA 
         Caption         =   "Load"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUnloadWA 
         Caption         =   "Unload"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSKIPMEBITCH 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "&Play"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNext 
         Caption         =   "&Next"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBack 
         Caption         =   "&Back"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMP3 
         Caption         =   "Mp3 Display"
         Begin VB.Menu mnuDispMP3Chan 
            Caption         =   "In Channel"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDispMP32me 
            Caption         =   "Just to me!"
         End
      End
   End
   Begin VB.Menu mnuBOoHooOOo 
      Caption         =   "|"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuCM 
      Caption         =   "Chat Modes"
      Begin VB.Menu mnuCanada 
         Caption         =   "Canada Mode"
      End
      Begin VB.Menu mnuLeet 
         Caption         =   "LeeT SpeaK"
      End
      Begin VB.Menu mnuFudd 
         Caption         =   "Elmer Fudd"
      End
      Begin VB.Menu mnuMoooo 
         Caption         =   "Cows go Mooooo!"
         Shortcut        =   ^{INSERT}
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEncrypt 
      Caption         =   "Encrypt"
      Begin VB.Menu mnuInvigEncrypt 
         Caption         =   "Invig Encrypt"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuHexEncrypt 
         Caption         =   "Hex Encrypt"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "EditMenu"
      Index           =   19
      Visible         =   0   'False
      Begin VB.Menu Copy 
         Caption         =   "Copy"
         Index           =   1
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu SelectAll 
         Caption         =   "SelectAll"
         Shortcut        =   ^A
         Visible         =   0   'False
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
Public AntiIdleSeconds As Long
Public random As Integer
Public dctime As Integer
Public IdleTime As Integer
Public invigchat As Integer
Public invignet As Integer
Public privatever As Boolean
Public intRandom As Integer
Public hexchat As Integer
Private Const WM_USER = &H400&
Private Const EM_AUTOURLDETECT = (WM_USER + 91)
Public WithEvents ChatBot As BnetBot
Attribute ChatBot.VB_VarHelpID = -1


Private Sub Main()
    frmMain.Show
End Sub
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

Private Sub bnu_Click()
On Error Resume Next
Send "/join Town Square", frmMain.wsBnet
AddChat vbYellow, "Joining Channel Town Square...."

End Sub

Private Sub Form_Load()
    s = SetSysColors(1, 9, &H0)
    m = SetSysColors(1, 7, &HF)
    Cm = SetSysColors(1, 4, &HF)
    Dim InvigNews As String
    Set ChatBot = New BnetBot
    Dim lRet As Long
    frmMain.Caption = "Invigoration - http://www.BNET.cc"
    LogChat = 0
    LoadConfig
    'InvigNews = devNEWS.OpenURL("http://www.bnet.cc/invigoration/news.txt")
    'InvigVer = devNEWS.OpenURL("http://www.bnet.cc/invigoration/version.txt")
    'InvigNight = devNEWS.OpenURL("http://www.bnet.cc/invigoration/nversion.txt")
    'InvigRel = devNEWS.OpenURL("http://www.bnet.cc/invigoration/verrelease.txt")
    AddChat D2MedBlue, "Version Check: "
    connectstatus = False
    If InvigVer = vernum Then
        privatever = False
        AddChat D2White, "You have the latest version of Invigoration."
        AddChat D2Orange, "Version ", D2White, InvigVer, D2Orange, " Released on: ", D2White, InvigRel
        mnuUpdate.Enabled = False
    ElseIf InvigVer < vernum Then
        privatever = True
        
        AddChat D2White, "You are running Invigoration Nightly Build...."
        mnuUpdate.Enabled = False
            Else
        privatever = False
        mnuUpdate.Enabled = True
        AddChat HEXPINK, "---------------------------------------------------------------------"
        AddChat D2Orange, "You need to update: Hold ", vbYellow, "CTRL ", D2Orange, "and press ", vbYellow, "U", D2Orange, " to launch update file."
        AddChat D2Orange, "You can also activate the Updater via the ", D2Green, "Other ", D2Orange, "menu."
        AddChat D2Orange, "Your version: " & vernum, D2White, "  Latest version: " & InvigVer
        AddChat D2Orange, "Last Update: ", D2White, X, D2Orange, " on ", D2White, R
        AddChat HEXPINK, "---------------------------------------------------------------------"
        msg = MsgBox("You need to update Invigoration. Please try using CTRL+U, if that doesn't work, close your bot and visit: http://www.clanbnu.ws/upgrade.html", vbOKOnly, "New Version Released!")
    End If
    AddChat D2Purple, "---------------------------------------------------"
    AddChat D2MedBlue, "()()"
    AddChat D2MedBlue, "(--)"
    AddChat D2MedBlue, "(')(')"
    AddChat D2Green, "Invigoration  rNightly  gBunny"
    AddChat D2Orange, "Public Open Source Version: " & vernum
    AddChat D2MedBlue, "---------------------------------------------------"
    frmConfigBNET.txtCDKey.text = GetStuff("BNET", "CDKey")
    random = "180"
    uptimesec = 0
    uptimemin = 0
    uptimehour = 0
    uptimedays = 0
    uptimeweek = 0
    uptimemonth = 0
    rtbsendbnet.SelColor = D2White
    ''.SelText = "this is my red text"
''End With
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

Private Sub IdleTimer_Timer()
    IdleTime = IdleTime + 1
            If IdleTime = idletimeset Then
                Send idleMessage, frmMain.wsBnet
                IdleTime = 0
            End If
End Sub
Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub
Private Sub mnuBan_Click()
    Send "/ban " & lstChannel.SelectedItem, wsBnet
    rtbsendbnet.SetFocus
End Sub

Private Sub mnuCanada_Click()
    If Canada = 0 Then
        Canada = 1
        AddChat D2White, "Canada mode enabled."
    ElseIf Canada = 1 Then
        Canada = 0
        AddChat D2White, "Canada mode disabled."
    End If
End Sub
Private Sub mnuFAdd_Click()
    Send "/f a " & lstChannel.SelectedItem, wsBnet
    rtbsendbnet.SetFocus
End Sub
Private Sub mnuFListView_Click()
    Send "/f l ", wsBnet
    rtbsendbnet.SetFocus
End Sub

Private Sub mnuFRemove_Click()
    Send "/f r " & lstChannel.SelectedItem, wsBnet
    rtbsendbnet.SetFocus
End Sub

Private Sub mnuFudd_Click()
    If fudd = 0 Then
        fudd = 1
        AddChat D2White, "Elmer Fudd mode enabled."
    ElseIf fudd = 1 Then
        fudd = 0
        AddChat D2White, "Elmer Fudd mode disabled."
    End If
End Sub

Private Sub mnuHexEncrypt_Click()
If hexchat = 1 Then
    hexchat = 0
    mnuInvigEncrypt.Enabled = True
    AddChat D2White, "Hex encryption disabled."
ElseIf hexchat = 0 Then
    hexchat = 1
    mnuInvigEncrypt.Enabled = False
    AddChat D2White, "Hex encryption enabled."
End If
End Sub

Private Sub mnuInvigEncrypt_Click()
If invigchat = 1 Then
    invigchat = 0
    mnuHexEncrypt.Enabled = True
    AddChat D2White, "Invig encryption disabled."
ElseIf invigchat = 0 Then
    invigchat = 1
    mnuHexEncrypt.Enabled = False
    AddChat D2White, "Invig encryption enabled."
End If
End Sub
Private Sub mnuKick_Click()
    Send "/kick " & lstChannel.SelectedItem, wsBnet
    rtbsendbnet.SelStart = Len(rtbsendbnet.text)
    rtbsendbnet.SetFocus
End Sub

Private Sub mnuBug_Click()
    ShellExecute Me.hWnd, "Open", "https://github.com/tagban/invigoration/issues", 0&, 0&, 0&
    AddChat D2White, "Thank you for contributing, we appreciate all of your testing! If a window didn't open, go here:  ohttps://github.com/tagban/invigoration/issues "
End Sub

Private Sub mnuClearBufs_Click()
    rtbChat.text = vbNullString
    AddChat D2White, "Cleared Chat Buffers. Old information is lost."
    rtbsendbnet.SetFocus
End Sub
Private Sub lstChannel_Click() '(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
If frmMain.wsBnet.State = sckConnected Then
        strList = lstChannel.SelectedItem
        PopupMenu mnuUserList
    rtbsendbnet.SetFocus2
Else
''
End If
End Sub
Private Sub lstChannel_Change() '(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
If frmMain.wsBnet.State = sckConnected Then
        strList = lstChannel.SelectedItem
        PopupMenu mnuUserList
    rtbsendbnet.SetFocus
Else
''
End If
End Sub
Private Sub mnuConnect_Click()
    With frmMain.lstChannel
        .ListItems.Clear
    End With
On Error GoTo Error
connectstatus = True
    tmrConnect.Enabled = True
    AddChat D2Green, "Battle.net Login Server Connecting to " & BNET.BNLSServer & "..."
    frmMain.wsBnls.Close
    frmMain.wsBnls.Connect BNET.BNLSServer, 9367
    frmMain.Caption = "Invigoration - [ Connecting to: " & BNET.BNLSServer & " ]"
    txtChannelInfo.Caption = "Connecting"
    txtChannelInfo.ForeColor = D2Green
    antiidlesecond = 0
    frmMain.tmrAntiIdle.Enabled = True
Error:

End Sub

Private Sub rtbsendbnet_Change()
    rtbsendbnet.SelColor = D2White
End Sub

Private Sub tmrAntiIdle_Timer()
AntiIdleSeconds = AntiIdleSeconds + 1
''AddChat D2Red, AntiIdleSeconds
If AntiIdleSeconds = 110 Then
    PBuffer.SendPacket &H0
AntiIdleSeconds = 0
End If
End Sub

Private Sub tmrConnect_Timer()
    connectseconds = connectseconds + 1
End Sub


Private Sub mnuDisconnect_Click()
    connectstatus = False
    tmrAntiIdle.Enabled = False
    frmMain.Caption = "Invigoration - http://www.BNET.cc/invigoration"
    frmMain.wsBnet.Close
    frmMain.wsBnls.Close
    AddChat D2White, "Battle.net Closed Connection."
    lstChannel.ListItems.Clear
    frmMain.Caption = "Invigoration - [ Disconnected ]"
    txtChannelInfo.Caption = "Disconnected"
    txtChannelInfo.ForeColor = D2Red
    antiidlesecond = 0
    dctime = 0
    frmMain.tmrAntiIdle.Enabled = False
    frmMain.tmrDC.Enabled = False
    mnuConnect.Visible = True
    mnuDisconnect.Visible = False
End Sub

Private Sub mnuEndFocus_Click()
    AddChat D2White, "User focus OFF"
        targetuser = vbNullString
        targetusername = vbNullString
    rtbsendbnet.SetFocus
End Sub

Private Sub mnuExit_Click()
    frmMain.wsBnet.Close
    frmMain.wsBnls.Close
    Unload Me
    End
End Sub

Private Sub mnuLeet_Click()
If leetspeak = 0 Then
    leetspeak = 1
    AddChat D2White, "Leet Speak enabled."
ElseIf leetspeak = 1 Then
    leetspeak = 0
    AddChat D2White, "Leet Speak disabled."
End If
End Sub
Private Sub mnuMoooo_Click()
    If moo = 0 Then
        moo = 1
        AddChat D2White, "Mooo?"
    ElseIf moo = 1 Then
        moo = 0
        AddChat D2White, "I lost my milk... :("
    End If
End Sub
Private Sub mnuSetupOption_Click()
    frmConfigBNET.Show
End Sub


Private Sub mnuSquelch_Click()
    Send "/squelch " & lstChannel.SelectedItem & Space(1), wsBnet
    rtbsendbnet.SelStart = Len(rtbsendbnet.text)
    rtbsendbnet.SetFocus
End Sub

Private Sub mnuUnSquelch_Click()
    Send "/unsquelch " & lstChannel.SelectedItem & Space(1), wsBnet
    rtbsendbnet.SelStart = Len(rtbsendbnet.text)
    rtbsendbnet.SetFocus
End Sub
Private Sub mnuUpdate_Click()
InvigVer = devNEWS.OpenURL("https://github.com/tagban/invigoration/")
    If InvigVer = vernum Then
    ''
Else
    ShellExecute Me.hWnd, "Open", "https://github.com/tagban/invigoration/", 0&, 0&, 0&
    AddChat D2Red, "If a window didn't open, close Invig, then download the new version from: https://github.com/tagban/invigoration/"
End If
connectstatus = False
End Sub
Private Sub mnuUserFocus_Click()
    targetuser = lstChannel.SelectedItem & " : "
    targetusername = lstChannel.SelectedItem
    
    AddChat D2MedBlue, targetusername & " is in Focus."
    rtbsendbnet.SetFocus
End Sub
Private Sub mnuViewProfile_Click()
    Profile.Caption = lstChannel.SelectedItem
    Profile.Show
    Profile.SetFocus
End Sub

Private Sub mnuWebsite_Click()
    ShellExecute Me.hWnd, "Open", "https://github.com/tagban/invigoration", 0&, 0&, 0&
End Sub
Private Sub mnuWhisper_Click()
    rtbsendbnet.text = "/w " & lstChannel.SelectedItem
    rtbsendbnet.SelColor = HEXPINK
    rtbsendbnet.SetFocus
    rtbsendbnet.SelStart = Len(rtbsendbnet.text)
End Sub

Private Sub Reconnect_Click()
    With frmMain.lstChannel
        .ListItems.Clear
    End With
On Error GoTo Error
    frmMain.tmrDC.Enabled = False
    frmMain.wsBnet.Close
    frmMain.wsBnls.Close
    connectstatus = False
    frmMain.Caption = "Invigoration - [ Reconnecting to: " & BNET.BNLSServer & " ]"
    AddChat D2Green, "Battle.net Login Server Reconnecting to " & BNET.BNLSServer & "..."
    frmMain.wsBnls.Close
    frmMain.wsBnls.Connect BNET.BNLSServer, 9367
    txtChannelInfo.Caption = "Reconnecting"
    txtChannelInfo.ForeColor = D2Orange
    connectstatus = True
    antiidlesecond = 0
    frmMain.tmrAntiIdle.Enabled = True
Error:

End Sub

Private Sub Form_Resize()
On Error GoTo Size
rtbChat.Height = Me.Height - 1450
rtbChat.Width = Me.Width - 3495
rtbsendbnet.Width = Me.Width - 3495
rtbsendbnet.Top = Me.ScaleHeight - 450
lstChannel.Left = Me.Width - 3350
lstChannel.Height = rtbChat.Height - 150
txtChannelInfo.Left = Me.Width - 3350
'frmWinamp.Top = Me.ScaleHeight - 550
'frmWinamp.Left = Me.Width - 1900
imgInvig.Top = Me.ScaleHeight - 450
imgInvig.Left = Me.Width - 3350

Size:

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
        frmMain.Caption = "Invigoration - [ Connecting to: " & BNET.BNLSServer & " ]"
        txtChannelInfo.Caption = "Connecting"
        txtChannelInfo.ForeColor = D2Green
        antiidlesecond = 0
        frmMain.tmrAntiIdle.Enabled = True
        connectstatus = True
    End If
End Sub

Public Sub tmrUptime_Timer()
    uptimesec = uptimesec + 1
        If uptimesec = 60 Then
            uptimemin = uptimemin + 1
            uptimesec = 0
        If uptimemin = 60 Then
            uptimehour = uptimehour + 1
            uptimemin = 0
        If uptimehour = 24 Then
            uptimedays = uptimedays + 1
            uptimehour = 0
        If uptimedays = 7 Then
            uptimeweek = uptimeweek + 1
            uptimedays = 0
        If uptimeweek = 4 Then
            uptimemonth = uptimemonth + 1
            uptimeweek = 0
        End If
        End If
End If
End If
End If

End Sub
Private Sub ToTray_Click()
TrayToolTip = "Invigoration [ " & BNET.username & " ]" & vbNewLine & "Version : " & vernum
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

Private Sub rtbSendBNET_KeyPress(keyascii As Integer)
            rtbsendbnet.SelColor = D2White
            If rtbsendbnet.text = "/r " Then
                rtbsendbnet.text = "/w " & LastW & Space(1)
                rtbsendbnet.SelStart = Len(rtbsendbnet.text)
            ElseIf rtbsendbnet.text = "/t " Then
                rtbsendbnet.text = "/w "
                rtbsendbnet.SelStart = Len(rtbsendbnet.text)
            ElseIf rtbsendbnet.text = "/l " Then
                rtbsendbnet.text = "/fmsg "
                rtbsendbnet.SelStart = Len(rtbsendbnet.text)
            ElseIf rtbsendbnet.text = "/em " Then
                rtbsendbnet.text = "/me "
                rtbsendbnet.SelStart = Len(rtbsendbnet.text)
            End If
    If Len(rtbsendbnet.text) = 0 Then
    ''
    rtbsendbnet.SelColor = D2White
    Else
    rtbsendbnet.SelColor = D2White
    If keyascii = "13" Then
        keyascii = "0"
        If Len(rtbsendbnet.text) >= 1 Then
                Dim Message As String
                Message = rtbsendbnet.text
                DoAddToSendList rtbsendbnet.text
                rtbsendbnet.text = vbNullString
                Dim channel As String
            If Left$(Message, 1) = "/" Then
                ParseCommand Message, BNET.username, True
                Exit Sub
            End If
            'Canada Mode Replacements
            If Canada Then
                Message = Message & ", eh?"
                Message = Replace(Message, "police", "mounties")
                Message = Replace(Message, "house", "igloo")
                Message = Replace(Message, "beer", "Labatt Blue")
                Message = Replace(Message, "drinking", "Seein' the Governor")
                Message = Replace(Message, "drinks", "Molsen Ices")
                Message = Replace(Message, "drink", "Molsen Ice")
                Message = Replace(Message, "weather", "monsoon")
                Message = Replace(Message, "crazy", "cookie")
                Message = Replace(Message, "dollar", "Loonie")
                Message = Replace(Message, "headache", "Skull Cramp")
                Message = Replace(Message, "raingear", "Oil Cloths")
                Message = Replace(Message, "coffee", "timmie")
                Message = Replace(Message, "huh", "eh")
                Message = Replace(Message, "friend", "hoser")
                Message = Replace(Message, "out", "oot")
                Message = Replace(Message, "ues", "ooz")
            End If
            'Leet Speak replacements
            If leetspeak Then
                Message = Replace(Message, "A", "4")
                Message = Replace(Message, "a", "4")
                Message = Replace(Message, "b", "8")
                Message = Replace(Message, "B", "8")
                Message = Replace(Message, "c", "¢")
                Message = Replace(Message, "C", "((")
                Message = Replace(Message, "d", "|)")
                Message = Replace(Message, "D", "|)")
                Message = Replace(Message, "e", "3")
                Message = Replace(Message, "E", "3")
                Message = Replace(Message, "f", "f")
                Message = Replace(Message, "F", "F")
                Message = Replace(Message, "g", "6")
                Message = Replace(Message, "h", "h")
                Message = Replace(Message, "H", "|-|")
                Message = Replace(Message, "i", "!")
                Message = Replace(Message, "I", "!")
                Message = Replace(Message, "k", "|<")
                Message = Replace(Message, "K", "]{")
                Message = Replace(Message, "l", "1")
                Message = Replace(Message, "L", "£")
                Message = Replace(Message, "M", "/\/\")
                Message = Replace(Message, "n", "ñ")
                Message = Replace(Message, "N", "/\/")
                Message = Replace(Message, "o", "0")
                Message = Replace(Message, "O", "0")
                Message = Replace(Message, "P", "|o")
                Message = Replace(Message, "Q", "9")
                Message = Replace(Message, "q", "9")
                Message = Replace(Message, "R", "|2")
                Message = Replace(Message, "s", "5")
                Message = Replace(Message, "S", "5")
                Message = Replace(Message, "t", "t")
                Message = Replace(Message, "T", "7")
                Message = Replace(Message, "U", "(_)")
                Message = Replace(Message, "V", "\/")
                Message = Replace(Message, "W", "\//")
                Message = Replace(Message, "x", "×")
                Message = Replace(Message, "Y", "¥")
                Message = Replace(Message, "y", "y")
                Message = Replace(Message, "Z", "2")
            End If
            'Elmer Fudd... you wascally wabbit!
            If fudd Then
                Message = Replace(Message, "l", "w")
                Message = Replace(Message, "L", "W")
                Message = Replace(Message, "r", "w")
                Message = Replace(Message, "R", "W")
            End If
            'MooMode Trick your friends! tell them to hit CTRL+M
            If moo Then
                If Len(Message) > 25 Then
                    Message = "Mooooo the aliens are coming for me! Mooo!!!"
                ElseIf Len(Message) > 20 Then
                    Message = "Cheeessyyyy"
                ElseIf Len(Message) > 15 Then
                    Message = "Got Milk?"
                Else
                    Message = "Moooo!"
                End If
            End If
            End If
            If Message = "/" Then
            On Error GoTo Err:
                AddChat D2Red, "Cannot send just /. Please specify the command."
            ElseIf invigchat = 1 Then
                If wsBnet.State <> Connected Then
                    Send Chr(149) & InvigEncrypt(Message & "-"), frmMain.wsBnet
                Else
                    AddChat D2Red, "Battle.net is not connected."
                End If
            ElseIf hexchat = 1 Then
                If wsBnet.State <> Connected Then
                    Send Chr(163) & StrToHex(Message), frmMain.wsBnet
                Else
                    AddChat D2Red, "Battle.net is not connected."
                End If
            Else
                If wsBnet.State <> Connected Then
                    Send targetuser & Message & postpend, wsBnet
                Else
                    AddChat D2Red, "Battle.net is not connected."
                End If
            End If
Err:
            
        End If
    End If
End Sub


Private Sub Update_Click()
    ShellExecute Me.hWnd, "Open", "http://www.bnet.cc/invigoration", 0&, 0&, 0&
End Sub
Private Sub wsBnet_Close()
    mnuConnect.Visible = True
    mnuDisconnect.Visible = False
    AddChat D2Red, "Battle.net Disconnected"
    frmMain.Caption = "Invigoration - [ Disconnected ]"
    lstChannel.ListItems.Clear
    txtChannelInfo.Caption = "Disconnected"
    txtChannelInfo.ForeColor = D2Red
End Sub

Private Sub wsBnet_Connect()
    mnuConnect.Visible = False
    mnuDisconnect.Visible = True
    AddChat D2Green, "Battle.net Connected!"
    frmMain.Caption = "Invigoration - [ " & BNET.username & " ]"
    wsBnet.SendData Chr(1)
    Send0x50
    txtChannelInfo.ForeColor = D2LtBlue
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
    frmMain.Caption = "Invigoration - [BNLS Connection CLOSED]"
End Sub

Private Sub wsBnls_Connect()
    AddChat D2Green, "Battle.net Login Server Connected!"
    frmMain.Caption = "Invigoration - [BNLS Connected to " & wsBnls.RemoteHostIP & "]"
    With PBuffer
        .InsertNTString "Invigoration"
        .SendBNLSPacket &HE
    End With
End Sub

Private Sub wsBnls_DataArrival(ByVal bytesTotal As Long)
Dim TempData As String
    wsBnls.GetData TempData, vbString
    ParseBNLS TempData
End Sub

Private Sub ChatBot_OnUser(ByVal username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
Dim ParsedString As String, thing As New BnetBot
    Call ParseStatString(Message, ParsedString)
    newest = username
    '''''''''''''''''''''''''''''''''''''''''''''''
    'Special Flags Fun!? w00t Added 9/14/2010 - Tagban
    'All developers for Invigoration are recommended to add their own.
    If LCase(username) = "tagban" Then
        Flags = &H80000
    ElseIf LCase(username) = "Invigoration" Then
        Flags = &H80000
    ElseIf LCase(username) = "bnu-bot" Then
        Flags = &H800000
    Else
        If BNET.BNCCICON = 1 Then
                If LCase(username) = BNET.username Then
                    Flags = &H800000
                End If
        End If
    End If
    'End of Special Flags Code
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    With frmMain.lstChannel
            '.ListItems(frmMain.lstChannel.ListItems.Count).ListSubItems.Add
        If Flags = "2" Then
            .ListItems.Add 1, , username, , thing.GetIconCode(Message, Flags)
            .ListItems(1).ListSubItems.Add , , GetPingCode(Ping)
            .ListItems(1).ToolTipText = "[" & ParsedString & "](" & Ping & "ms) MODERATOR"
        ElseIf Flags = "1" Then
            .ListItems.Add 1, , username, , thing.GetIconCode(Message, Flags)
            .ListItems(1).ListSubItems.Add , , GetPingCode(Ping)
            .ListItems(1).ToolTipText = "Blizzard Representative"
        ElseIf Flags = "&H80000" Then
            .ListItems.Add , , username, , thing.GetIconCode(Message, Flags)
            .ListItems(frmMain.lstChannel.ListItems.Count).ListSubItems.Add , , GetPingCode(Ping)
            .ListItems(frmMain.lstChannel.ListItems.Count).ToolTipText = "[Invigoration Development Team]"
        Else
            .ListItems.Add , , username, , thing.GetIconCode(Message, Flags)
            .ListItems(frmMain.lstChannel.ListItems.Count).ListSubItems.Add , , GetPingCode(Ping)
            .ListItems(frmMain.lstChannel.ListItems.Count).ToolTipText = "[" & ParsedString & "](" & Ping & "ms)"
        End If
        GetPingColor username, Flags, Ping
        If Flags = &H1 Then
            .ListItems(frmMain.lstChannel.ListItems.Count).ForeColor = D2MedBlue
        Else
            .ListItems(frmMain.lstChannel.ListItems.Count).ForeColor = D2White
        End If
    .Refresh
    End With
    With frmMain.txtChannelInfo
        .Caption = BNET.CurrentChan & " (" & frmMain.lstChannel.ListItems.Count & ")"
    End With
End Sub


Private Sub ChatBot_OnChannel(ByVal ChannelName As String, ByVal Flags As Long)
    With frmMain.lstChannel
        .ListItems.Clear
    End With
    txtChannelInfo.ForeColor = D2LtBlue
    With frmMain.txtChannelInfo
        .Caption = BNET.CurrentChan & " (1)"
    End With
        BanCount = 0
        KickCount = 0
        JoinCount = 0
    With frmMain.lstChannel
        Refresh
    End With
    BNET.CurrentChan = ChannelName
    AddChat D2Orange, "Joining (" & GetChannelType(Flags) & ") " & ChannelName & " w/Flags: " & Flags & "."
End Sub

Private Sub ChatBot_OnEmote(ByVal username As String, ByVal Flags As Long, ByVal Message As String)
    AddChat D2Beige1, "* " & username & Space(1) & Message & " *"
End Sub
Private Sub ChatBot_OnError(ByVal Message As String)

If Message = "That is not a valid command. Type /help or /? for more info." Then
'' nothing "That channel is restricted."
ElseIf Message = "You are banned from that channel." Then
    Send "/join Invig Rejects", frmMain.wsBnet
    AddChat D2LtYellow, "You've been transported to the 'Invig Rejects' channel because the channel you attempted to enter, or were in, was 'restricted'."
Else
    AddChat D2Red, "ERROR: " & Message
End If
End Sub

Private Sub ChatBot_OnInfo(ByVal Message As String)

    If Message = "You are still marked as being away." Or InStr(Message, "(-- Invigoration") Then
    ElseIf InStr(Message, "was kicked out of the channel by") Then
        AddChat D2LtYellow, Message
        KickCount = KickCount + 1
    ElseIf InStr(Message, "was banned by") Then
        AddChat D2LtYellow, Message
        BanCount = BanCount + 1
    ElseIf InStr(Message, "kicked you") Then
        AddChat D2LtYellow, Message
        Send "/join " & BNET.CurrentChan, frmMain.wsBnet
    Else
        AddChat vbYellow, Message
    End If
End Sub

Private Sub ChatBot_OnFlags(ByVal username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
Dim thing As New BnetBot
    '''''''''''''''''''''''''''''''''''''''''''''''
    'Special Flags Fun!? w00t Added 9/14/2010 - Tagban
    'All developers for Invigoration are recommended to add their own.
    If LCase(username) = "tagban" And BNET.BattlenetServer = "connect-use.classic.blizzard.com" Then
        Flags = &H80000
    ElseIf LCase(username) = "tagban" And BNET.BattlenetServer = "atlas.bnet.cc" Then
        Flags = &H80000
    ElseIf LCase(username) = "bnu-bot" Then
        Flags = &H800000
    ElseIf LCase(username) = "tagban" And BNET.BattlenetServer = "us.battle.vet" Then
        Flags = &H80000
    Else
        If BNET.BNCCICON = 1 Then
                If LCase(username) = BNET.username Then
                    Flags = &H800000
                End If
        End If
    End If
    'End of Special Flags Code
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    For X = 1 To frmMain.lstChannel.ListItems.Count
        If frmMain.lstChannel.ListItems.Item(X).text = username Then
            With frmMain.lstChannel
                .ListItems.Item(X).SmallIcon = thing.GetIconCode(Message, Flags)
                    If (Flags And &H2) = &H2 Then
                        .ListItems.Remove X ' Add 1
                        .ListItems.Add 1, , username, , thing.GetIconCode(Message, Flags)
                        .ListItems(1).ListSubItems.Add , , GetPingCode(Ping)
                        .ListItems(1).ToolTipText = "[" & ParsedString & "]"
                    End If
            End With
        End If
    Next X
End Sub



Private Sub ChatBot_OnJoin(ByVal username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
Dim ParsedString As String, thing As New BnetBot
    Call ParseStatString(Message, ParsedString)
    With frmMain.lstChannel
    '''''''''''''''''''''''''''''''''''''''''''''''
    'Special Flags Fun!? w00t Added 9/14/2010 - Tagban
    'All developers for Invigoration are recommended to add their own.
    If LCase(username) = "tagban" And BNET.BattlenetServer = "connect-use.classic.blizzard.com" Then
        Flags = &H80000
    ElseIf LCase(username) = "tagban" And BNET.BattlenetServer = "atlas.bnet.cc" Then
        Flags = &H80000
    ElseIf LCase(username) = "bnu-bot" Then
        Flags = &H800000
    ElseIf LCase(username) = "tagban" And BNET.BattlenetServer = "us.battle.vet" Then
        Flags = &H80000
    Else
        If BNET.BNCCICON = 1 Then
                If LCase(username) = BNET.username Then
                    Flags = &H800000
                End If
        End If
    End If
    'End of Special Flags Code
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    If LCase(frmMain.lstChannel.ListItems.Count) = LCase(BNET.username) Then
         '' Do Something
    Else
    If Flags = "1" Then
        .ListItems.Add 1, , username, , thing.GetIconCode(Message, Flags)
        .ListItems(1).ListSubItems.Add , , GetPingCode(Ping)
        .ListItems(1).ToolTipText = "[" & ParsedString & "](" & Ping & "ms) moderator"
        JoinCount = JoinCount + 1
    Else
        .ListItems.Add , , username, , thing.GetIconCode(Message, Flags)
        .ListItems(frmMain.lstChannel.ListItems.Count).ListSubItems.Add , , GetPingCode(Ping)
        .ListItems(frmMain.lstChannel.ListItems.Count).ToolTipText = "[" & ParsedString & "](" & Ping & "ms)"
        JoinCount = JoinCount + 1
    End If
        GetPingColor username, Flags, Ping
        If BNET.JoinNotify = "1" Then
            AddChat D2Green, username & " has entered the channel. Ping: " & Ping & "ms"
        Else
        End If
End If
    With frmMain.txtChannelInfo
        .Caption = BNET.CurrentChan & " (" & frmMain.lstChannel.ListItems.Count & ")"
    End With
.Refresh
End With
End Sub
Private Sub ChatBot_OnLeave(ByVal username As String, ByVal Flags As Long)
    '''''''''''''''''''''''''''''''''''''''''''''''
    'Special Flags Fun!? w00t Added 9/14/2010 - Tagban
    If LCase(username) = "tagban" And BNET.BattlenetServer = "connect-use.classic.blizzard.com" Then
        Flags = &H80000
    ElseIf LCase(username) = "tagban" And BNET.BattlenetServer = "atlas.bnet.cc" Then
        Flags = &H80000
    ElseIf LCase(username) = "bnu-bot" Then
        Flags = &H800000
    ElseIf LCase(username) = "tagban" And BNET.BattlenetServer = "us.battle.vet" Then
        Flags = &H80000
    Else
        If BNET.BNCCICON = 1 Then
                If LCase(username) = BNET.username Then
                    Flags = &H800000
                End If
        End If
    End If
    'End of Special Flags Code
    ''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
    With frmMain.lstChannel
        .ListItems.Remove frmMain.lstChannel.FindItem(username).Index
    End With
    With frmMain.txtChannelInfo
        .Caption = BNET.CurrentChan & " (" & frmMain.lstChannel.ListItems.Count & ")"
    End With
    
    If BNET.JoinNotify = "1" Then
        AddChat D2Green, username & " has left the channel."
    Else
    End If
End Sub

Private Sub ChatBot_OnTalk(ByVal username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
    Message = Replace(Message, "ÿC1", vbNullString)
    Message = Replace(Message, "ÿC2", vbNullString)
    Message = Replace(Message, "ÿC3", vbNullString)
    Message = Replace(Message, "ÿC4", vbNullString)
    Message = Replace(Message, "ÿC5", vbNullString)
    Message = Replace(Message, "ÿC6", vbNullString)
    Message = Replace(Message, "ÿC7", vbNullString)
    Message = Replace(Message, "ÿC8", vbNullString)
    Message = Replace(Message, "ÿC9", vbNullString)
    Message = Replace(Message, "ÿC:", vbNullString)
    Message = Replace(Message, "ÿC;", vbNullString)
    Message = Replace(Message, "ÿC<", vbNullString)
    Message = Replace(Message, "ÁP", vbNullString)
    Message = Replace(Message, "ÁQ", vbNullString)
    Message = Replace(Message, "ÁR", vbNullString)
    Message = Replace(Message, "ÁS", vbNullString)
    Message = Replace(Message, "ÁT", vbNullString)
    Message = Replace(Message, "ÁU", vbNullString)
    Message = Replace(Message, "ÁV", vbNullString)
    Message = Replace(Message, "ÁW", vbNullString)
    Message = Replace(Message, "ÁX", vbNullString)
    Message = Replace(Message, "ÁY", vbNullString)
    Message = Replace(Message, "ÁZ", vbNullString)
    Message = Replace(Message, "Á[", vbNullString)
    'Message Data
    Select Case Mid$(Message, 1, 1)
        Case Chr(163)
            AddChat D2Beige1, ":: " & username & " :: ", HEXPINK, "[HEX] " & HexToStr(Mid$(Message, 2, Len(Message)))
        Case Chr(149)
            AddChat D2Beige1, ":: " & username & " :: ", HEXPINK, "[INVIG] " & InvigDecrypt(Mid$(Message, 2, Len(Message)))
        Case Else
             AddChat D2Beige1, ":: " & username & " :: ", D2White, Message
        End Select
ParseCommand Message, username 'send commands to parser
Call fBotColors
End Sub


Private Sub ChatBot_OnUnknown(ByVal UnknownString As String)
    AddChat D2Purple, "UNKNOWN STRING: ", vbRed, UnknownString
End Sub

Private Sub ChatBot_OnWhisperFrom(ByVal username As String, ByVal Flags As Long, ByVal Message As String)

    AddChat D2Beige1, ":: Whisper From: " & username & " :: ", vbGrey, Message
    LastW = username
    'LastCW = "/w " & username & " "
    LastM = Message
    ParseCommand Message, username 'send commands to parser
End Sub

Private Sub ChatBot_OnWhisperTo(ByVal username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)

If InStr(Message, "Invigoration BETA version:") Then
''
Else
        AddChat D2Beige1, ":: Whisper To: " & username & " :: ", vbGrey, Message
    LastSW = username
    LastSM = Message
End If
End Sub
Private Sub mnuQC1_Click()
On Error Resume Next
    Send "/join " & lbQC.List(0), frmMain.wsBnet

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
