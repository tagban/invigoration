VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About "
   ClientHeight    =   4815
   ClientLeft      =   2190
   ClientTop       =   1425
   ClientWidth     =   5835
   ClipControls    =   0   'False
   FillColor       =   &H00FFC0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3323.4
   ScaleMode       =   0  'User
   ScaleWidth      =   5479.367
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4125
      TabIndex        =   0
      Top             =   4320
      Width           =   1500
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "Project Page@GoogleCode"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "DarkBlizz.org"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Myst of DarkBlizz for reviving my interest in this project and resolving a BNLS Issue I was having."
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Clan BNU's Website"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   7
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H80000007&
      Caption         =   "Version 1.0.2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   $"frmAbout.frx":000C
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   5415
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Invigoration sites/contributors:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   5655
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   112.686
      X2              =   5408.938
      Y1              =   2319.133
      Y2              =   2319.133
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Special Thanks to:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Written by BNU-Master (Aka Tagban) of BNET.cc and Clan BNU"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   450
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   4725
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "InvigWeb"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   720
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4485
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFC0&
      Index           =   0
      X1              =   98.6
      X2              =   5408.938
      Y1              =   662.609
      Y2              =   662.609
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "BNET.cc"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblVersion.Caption = botver
End Sub

Private Sub Label10_Click(Index As Integer)
    ShellExecute Me.hWnd, "Open", "http://www.clanbnu.net", 0&, 0&, 0&
End Sub


Private Sub Label3_Click()
    ShellExecute Me.hWnd, "Open", "http://www.bnet.cc", 0&, 0&, 0&
End Sub

Private Sub Label5_Click()
    ShellExecute Me.hWnd, "Open", "http://www.darkblizz.org", 0&, 0&, 0&
End Sub

Private Sub Label6_Click()
    ShellExecute Me.hWnd, "Open", "http://code.google.com/p/BNET.cc/", 0&, 0&, 0&
End Sub
