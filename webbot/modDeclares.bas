Attribute VB_Name = "modDeclares"
Public Const botver = "1.0.0[BnetWeb]"
Public Const vernum = "1.0.0"
''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function GetTickCount& Lib "KERNEL32" ()
'World-Accessable declares
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMessage As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal numbytes As Long)
Public Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Public constants

Public Const CRC32_POLYNOMIAL As Long = &HEDB88320
Public Const vbBage = &H80000000
Public Const vbGrey = &H808080
Public Const vbTeal = &HFFFF00
Public Const vbDGreen = &H8000&
Public Const CHNFLAG_PUB = &H1
Public Const CHNFLAG_MOD = &H2
Public Const CHNFLAG_STS = &H4
Public Const CHNFLAG_VOID = &H8
Public Const CHNFLAG_ADMIN = &H10
Public Const CHNFLAG_PROD = &H20
Public Const CHNFLAG_ALL = &H1000
' D2 Color Codes C/o BNU`Bot
Public Const D2White = &HFFFFFF
Public Const D2Red = &H3E3ECE
Public Const D2Green = &HCE00&
Public Const D2Blue = &H9C4044
Public Const D2Beige1 = &H6091A1
Public Const D2Gray = &H555555
Public Const D2Black = &H80808
Public Const D2Beige2 = &H659DA8
Public Const D2Orange = &H88CE&
Public Const D2LtYellow = &H51CECE
Public Const D2Purple = &HCE008D
Public Const D2Cyan = &HFFFF00
Public Const D2MedBlue = &HE8AC2C
Public Const D2LtBlue = &HC0C000
Public Const HEXPINK = &H9900FF

Public Type BotData
    username As String
    Password As String
    CDKey As String
    BattlenetServer As String
    BNLSServer As String
    HomeChannel As String
    Product As String
    TrueUsername As String
    NewPass As String
    CurrentChan As String
    FontSize As String
    WebUser As String
    WebPass As String
End Type
Public BNET As BotData
Public Type BotNetData
    username As String
    Password As String
    Database As String
    DatabasePassword As String
    Connected As Integer
    WebUser As String
    WebPass As String
End Type
Public BOTNET As BotNetData
Public PBuffer As New PacketBuffer

'Public nondependant variables

Public version As Long
Public CheckSum As Long
Public ExeInfo As String
Public Servers As Long
Public CdkeyHash As String
Public GTC As Long
Public HType As Long
Public CB As Long
Public SPass As Boolean
Public VerByte As Long
Public P1 As String
Public P2 As String
Public AttemptedC As Boolean
Public LRealm As Boolean
Public CRC32Table(0 To 255) As Long
Public hash(2) As String
Public Temporary As String
