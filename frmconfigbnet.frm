VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmConfigBNET 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Battle.net Configuration"
   ClientHeight    =   3690
   ClientLeft      =   1875
   ClientTop       =   1935
   ClientWidth     =   7575
   ControlBox      =   0   'False
   FillColor       =   &H00000080&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "frmConfigBNET.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7575
   Begin VB.ComboBox txtBattlenet 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   315
      ItemData        =   "frmConfigBNET.frx":08CA
      Left            =   2520
      List            =   "frmConfigBNET.frx":08DA
      TabIndex        =   24
      Text            =   "BNET Servers"
      Top             =   480
      Width           =   2175
   End
   Begin MSComctlLib.ListView lstchannels 
      Height          =   30
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   53
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdCommit 
      BackColor       =   &H8000000A&
      Caption         =   "&Accept"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   3000
      Width           =   855
   End
   Begin VB.Frame frmBattlenet 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Invigoration Config"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   7335
      Begin VB.CheckBox chkUDP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "-1ms"
         ForeColor       =   &H0000C0C0&
         Height          =   195
         Left            =   6000
         MaskColor       =   &H00000000&
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1935
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   4695
         Begin VB.TextBox txtEmail 
            BackColor       =   &H00000040&
            ForeColor       =   &H000080FF&
            Height          =   285
            Left            =   240
            TabIndex        =   33
            Text            =   "e-mail addr"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtCDKey2 
            BackColor       =   &H00000040&
            ForeColor       =   &H000080FF&
            Height          =   285
            Left            =   3240
            TabIndex        =   31
            Text            =   "1234567891011"
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtTrigger 
            BackColor       =   &H00000040&
            ForeColor       =   &H000080FF&
            Height          =   285
            Left            =   4200
            MaxLength       =   1
            TabIndex        =   22
            Text            =   "!"
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox txtBotMaster 
            BackColor       =   &H00000040&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   285
            Left            =   240
            TabIndex        =   16
            ToolTipText     =   $"frmConfigBNET.frx":0928
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtCDKey 
            BackColor       =   &H00000040&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3240
            TabIndex        =   15
            Text            =   "1234567891011"
            ToolTipText     =   "Input Product CD Key"
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox txtProduct 
            BackColor       =   &H00000040&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   315
            ItemData        =   "frmConfigBNET.frx":09B3
            Left            =   840
            List            =   "frmConfigBNET.frx":09C6
            TabIndex        =   14
            Text            =   "PXES"
            ToolTipText     =   "Product ID (Backwards)"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label3 
            BackColor       =   &H00000000&
            Caption         =   "User e-Mail"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000007&
            BackStyle       =   0  'Transparent
            Caption         =   "EXP Key"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   2400
            TabIndex        =   30
            Top             =   600
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Bot Trigger"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   2280
            TabIndex        =   23
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Bot Master"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "CDKey"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   2520
            TabIndex        =   17
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label12 
            BackColor       =   &H00000000&
            Caption         =   "Product:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Other 
         BackColor       =   &H80000007&
         Caption         =   "Other Options"
         ForeColor       =   &H0000FFFF&
         Height          =   2175
         Left            =   5040
         TabIndex        =   21
         Top             =   240
         Width           =   2175
         Begin VB.ComboBox cmbLanguage 
            BackColor       =   &H00000040&
            Enabled         =   0   'False
            ForeColor       =   &H000080FF&
            Height          =   315
            ItemData        =   "frmConfigBNET.frx":0A02
            Left            =   120
            List            =   "frmConfigBNET.frx":0A15
            TabIndex        =   29
            Text            =   "English"
            Top             =   1680
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.CheckBox chkJoinNotify 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000012&
            Caption         =   "Join Notifications"
            ForeColor       =   &H0000C0C0&
            Height          =   375
            Left            =   240
            TabIndex        =   27
            Top             =   720
            Width           =   1815
         End
         Begin VB.CheckBox chkPing 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000012&
            Caption         =   "Lag Bars"
            ForeColor       =   &H0000C0C0&
            Height          =   255
            Left            =   960
            TabIndex        =   26
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lblEmail 
            BackStyle       =   0  'Transparent
            Caption         =   "Language:"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1320
            Visible         =   0   'False
            Width           =   1935
         End
      End
      Begin VB.Frame HomeChannel 
         BackColor       =   &H00000000&
         Caption         =   "Home Channel"
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   2280
         TabIndex        =   19
         Top             =   960
         Width           =   2535
         Begin VB.TextBox txtHomeChannel 
            BackColor       =   &H00000040&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   120
            TabIndex        =   20
            Text            =   "Clan BNU"
            ToolTipText     =   "Your Home Channel"
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   2535
         Begin VB.TextBox aSDASD 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   2520
            TabIndex        =   8
            Top             =   1320
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label4 
            BackColor       =   &H00000000&
            Caption         =   "Battle.net Server"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "BNLS URL/Server:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2520
            TabIndex        =   9
            Top             =   1320
            Visible         =   0   'False
            Width           =   2295
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2055
         Begin VB.TextBox txtUsername 
            BackColor       =   &H00000040&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   285
            Left            =   120
            MaxLength       =   15
            TabIndex        =   6
            Text            =   "InvigNetUser"
            ToolTipText     =   "Input Battle.net Username Here"
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2055
         Begin VB.TextBox txtPassword 
            BackColor       =   &H00000040&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "º"
            TabIndex        =   4
            Text            =   "Password"
            ToolTipText     =   "Password"
            Top             =   240
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frmConfigBNET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
    Me.Visible = False
End Sub

Private Sub cmdCommit_Click()
On Error GoTo Error
    BNET.username = txtUsername.text
    BNET.Password = txtPassword.text
    txtCDKey.text = Replace(txtCDKey.text, "-", vbNullString)
    txtCDKey2.text = Replace(txtCDKey2.text, "-", vbNullString)
    BNET.CDKey = txtCDKey.text
    BNET.CDKey2 = txtCDKey2.text
    If txtProduct.text = "Starcraft" Then
        BNET.Product = "RATS"
    ElseIf txtProduct.text = "Brood War" Then
        BNET.Product = "PXES"
    ElseIf txtProduct.text = "Warcraft 2" Then
        BNET.Product = "NB2W"
    ElseIf txtProduct.text = "Warcraft 3" Then
        BNET.Product = "3RAW"
    ElseIf txtProduct.text = "Diablo 2" Then
        BNET.Product = "VD2D"
    ElseIf txtProduct.text = "D2:LOD" Then
        BNET.Product = "PX2D"
    ElseIf txtProduct.text = "War3:TFT" Then
        BNET.Product = "PX3W"
    End If
    BNET.BattlenetServer = txtBattlenet.text
    BNET.email = txtEmail.text
    BNET.HomeChannel = txtHomeChannel.text
    BNET.BotMaster = txtBotMaster.text
    'BNET.Trigger = txtTrigger.text
    If chkUDP.value = vbUnchecked Then
        BNET.UDP = 0
    ElseIf chkUDP.value = vbChecked Then
        BNET.UDP = 1
    End If
    If chkPing.value = vbUnchecked Then
        BNET.ShowPing = 0
    ElseIf chkPing.value = vbChecked Then
        BNET.ShowPing = 1
    End If
    If chkJoinNotify.value = vbUnchecked Then
        BNET.JoinNotify = 0
    ElseIf chkJoinNotify.value = vbChecked Then
        BNET.JoinNotify = 1
    End If
    'BNET.UDP = txtUDP.text
    'BNET.ShowPing = txtShowPing.text
    'BNET.JoinNotify = txtJoinNotify.text
    BNET.Trigger = txtTrigger.text
    BNET.BNLSServer = txtBNLSServer.text
    BNET.email = txtEmail.text
    BNET.Realm = txtBattlenet.text
    
    SaveConfig
    Me.Visible = True
Error:
  Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Error
    
    frmBattlenet.Visible = True
    
    txtUsername.text = BNET.username
    txtPassword.text = BNET.Password
    txtCDKey.text = BNET.CDKey
    txtCDKey2.text = BNET.CDKey2
    If BNET.Product = "RATS" Then
        txtProduct.text = "Starcraft"
    ElseIf BNET.Product = "PXES" Then
        txtProduct.text = "Brood War"
    ElseIf BNET.Product = "NB2W" Then
        txtProduct.text = "Warcraft 2"
    ElseIf BNET.Product = "3RAW" Then
        txtProduct.text = "Warcraft 3"
    ElseIf BNET.Product = "VD2D" Then
        txtProduct.text = "Diablo 2"
    ElseIf BNET.Product = "PX2D" Then
        txtProduct.text = "D2:LOD"
    ElseIf BNET.Product = "PX3W" Then
        txtProduct.text = "War3:TFT"
    End If
    txtBattlenet.text = BNET.BattlenetServer
    txtEmail.text = BNET.email
    txtHomeChannel.text = BNET.HomeChannel
    txtBotMaster.text = BNET.BotMaster
    If BNET.UDP = 0 Then
        chkUDP.value = vbUnchecked
    Else
        chkUDP.value = vbChecked
    End If
    If BNET.ShowPing = 0 Then
        chkPing.value = vbUnchecked
    Else
        chkPing.value = vbChecked
    End If
    If BNET.JoinNotify = 0 Then
        chkJoinNotify.value = vbUnchecked
    Else
        chkJoinNotify.value = vbChecked
    End If
    txtTrigger.text = BNET.Trigger
    txtBNLSServer.text = BNET.BNLSServer
    txtEmail.text = BNET.email
    txtBattlenet.text = BNET.Realm
Error:

End Sub



Private Sub txtProduct_Change()
    If BNET.Product = "RATS" Then
        txtProduct.text = "Starcraft"
    ElseIf BNET.Product = "PXES" Then
        txtProduct.text = "Brood War"
    ElseIf BNET.Product = "NB2W" Then
        txtProduct.text = "Warcraft 2"
    ElseIf BNET.Product = "3RAW" Then
        chkUDP.value = vbUnchecked
        txtProduct.text = "Warcraft 3"
    ElseIf BNET.Product = "PX3W" Then
        chkUDP.value = vbUnchecked
        txtProduct.text = "Warcraft 3"
    ElseIf BNET.Product = "VD2D" Then
        txtProduct.text = "Diablo 2"
    End If
End Sub
