VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfigBNET 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Battle.net Configuration"
   ClientHeight    =   3945
   ClientLeft      =   1875
   ClientTop       =   1935
   ClientWidth     =   5160
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
   ScaleHeight     =   3945
   ScaleWidth      =   5160
   Begin VB.ComboBox txtBattlenet 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   315
      ItemData        =   "frmConfigBNET.frx":08CA
      Left            =   2520
      List            =   "frmConfigBNET.frx":08DA
      TabIndex        =   2
      Text            =   "BNET Servers"
      Top             =   600
      Width           =   2175
   End
   Begin MSComctlLib.ListView lstchannels 
      Height          =   30
      Left            =   240
      TabIndex        =   17
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
      Left            =   3960
      TabIndex        =   7
      Top             =   3240
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
      Left            =   3000
      TabIndex        =   6
      Top             =   3240
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
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4935
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
         TabIndex        =   15
         Top             =   1680
         Width           =   4695
         Begin VB.TextBox txtBNLS 
            BackColor       =   &H00000040&
            ForeColor       =   &H000080FF&
            Height          =   285
            Left            =   3360
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   960
            Width           =   1215
         End
         Begin VB.Frame webbot 
            BackColor       =   &H80000007&
            Caption         =   "Webbot Control"
            ForeColor       =   &H0000FFFF&
            Height          =   1695
            Left            =   120
            TabIndex        =   20
            Top             =   120
            Width           =   2295
            Begin VB.TextBox txtWebPass 
               BackColor       =   &H00000040&
               ForeColor       =   &H000080FF&
               Height          =   285
               Left            =   240
               TabIndex        =   22
               Text            =   "Text2"
               Top             =   1320
               Width           =   1935
            End
            Begin VB.TextBox txtWebUser 
               BackColor       =   &H00000040&
               ForeColor       =   &H000080FF&
               Height          =   285
               Left            =   240
               TabIndex        =   21
               Text            =   "Text1"
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Webbot Password"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   1080
               Width           =   2055
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Webbot Username"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   360
               Width           =   1575
            End
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
            TabIndex        =   5
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
            ItemData        =   "frmConfigBNET.frx":0928
            Left            =   3600
            List            =   "frmConfigBNET.frx":0941
            TabIndex        =   4
            Text            =   "PXES"
            ToolTipText     =   "Product ID (Backwards)"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "BNLS"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   2760
            TabIndex        =   26
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "CDKey"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   2520
            TabIndex        =   18
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
            Left            =   2520
            TabIndex        =   16
            Top             =   600
            Width           =   2055
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
            TabIndex        =   3
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
         TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   14
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
            TabIndex        =   13
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
         TabIndex        =   10
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
            TabIndex        =   0
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
         TabIndex        =   9
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
            TabIndex        =   1
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
    BNET.WebUser = txtWebUser.text
    BNET.WebPass = txtWebPass.text
    txtCDKey.text = Replace(txtCDKey.text, "-", vbNullString)
    BNET.CDKey = txtCDKey.text
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
    End If
    BNET.BattlenetServer = txtBattlenet.text
    BNET.HomeChannel = txtHomeChannel.text
    BNET.BNLSServer = txtBNLS.text
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
    txtWebUser.text = BNET.WebUser
    txtWebPass.text = BNET.WebPass
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
    End If
    txtBattlenet.text = BNET.BattlenetServer
    txtHomeChannel.text = BNET.HomeChannel
    txtBNLS.text = BNET.BNLSServer

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
    ElseIf BNET.Product = "VD2D" Then
        txtProduct.text = "Diablo 2"
    End If
End Sub
