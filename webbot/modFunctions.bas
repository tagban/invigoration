Attribute VB_Name = "modFunctions"
Public Declare Function SetForegroundWindow Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public nID As NOTIFYICONDATA
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
      
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_Message = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201 'Button down
Public Const WM_LBUTTONUP = &H202 'Button up
Public Const WM_LBUTTONDBLCLK = &H203 'Double-click
Public Const WM_RBUTTONDOWN = &H204 'Button down
Public Const WM_RBUTTONUP = &H205 'Button up
Public Const WM_RBUTTONDBLCLK = &H206 'Double-click
Public Function LoadConfig()
Dim FullText As String
    BNET.username = GetStuff("BNET", "Username")
    BNET.Password = GetStuff("BNET", "Password")
    BNET.CDKey = GetStuff("BNET", "CDKey")
    BNET.Product = GetStuff("BNET", "Product")
    BNET.BattlenetServer = GetStuff("BNET", "Server")
    BNET.HomeChannel = GetStuff("BNET", "Home")
    BNET.WebUser = GetStuff("BNET", "WebUser")
    BNET.WebPass = GetStuff("BNET", "WebPass")
    BNET.BNLSServer = GetStuff("BNLS", "Server")
End Function

Public Function SaveConfig()
Dim FullText As String
    WriteStuff "BNET", "Username", BNET.username
    WriteStuff "BNET", "Password", BNET.Password
    WriteStuff "BNET", "CDKey", BNET.CDKey
    WriteStuff "BNET", "Product", BNET.Product
    WriteStuff "BNET", "Server", BNET.BattlenetServer
    WriteStuff "BNET", "Home", BNET.HomeChannel
    WriteStuff "BNET", "WebUser", BNET.WebUser
    WriteStuff "BNET", "WebPass", BNET.WebPass
    WriteStuff "BNLS", "Server", BNET.BNLSServer
End Function

Public Function StrToHex(ByVal String1 As String) As String
    Dim strTemp As String, strReturn As String, i As Long
        For i = 1 To Len(String1)
            strTemp = Hex(Asc(Mid(String1, i, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
            strReturn = strReturn & strTemp
        Next i
    StrToHex = strReturn
End Function

Public Function ToHex(data As String) As String
Dim i As Integer
For i = 1 To Len(data)
    ToHex = ToHex & Right("00" & Hex(Asc(Mid(data, i, 1))), 2)
Next i
End Function

Public Function GetWORD(data As String) As Long
Dim lReturn As Long
    Call CopyMemory(lReturn, ByVal data, 2)
    GetWORD = lReturn
End Function
Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = CLng((300000 - 120000 + 1) * Rnd) + 120000
End Function
Public Function HexToStr(ByVal Hex1 As String) As String
    Dim strTemp As String, strReturn As String, i As Long
    Hex1 = Replace(Hex1, Space(1), vbNullString)
    If Len(Hex1) Mod 2 <> 0 Then Exit Function

    For i = 1 To Len(Hex1) Step 2
    strReturn = strReturn & Chr(Val("&H" & Mid(Hex1, i, 2)))
    Next i
    HexToStr = strReturn
End Function

Public Sub AddChat(ParamArray Message() As Variant)
On Error Resume Next

    Dim i As Integer
    With frmMain.rtbChat
        .SelStart = Len(.text)
        .SelLength = 0
        .SelColor = D2Gray
        .SelText = "|" & Format(Time, "HH:MM:SS") & "| "
    End With

   For i = LBound(Message) To UBound(Message) Step 2
       With frmMain.rtbChat
           .SelStart = Len(.text)
           .SelLength = 0
           .SelColor = Message(i)
           .SelText = Message(i + 1)
       End With
    Next i
    frmMain.rtbChat.SelText = vbCrLf
End Sub
Public Sub AddChat2(ParamArray Message() As Variant)
On Error Resume Next

    Dim i As Integer
   For i = LBound(Message) To UBound(Message) Step 2
       With frmMain.rtbChat
           .SelStart = Len(.text)
           .SelLength = 0
           .SelColor = Message(i)
           .SelText = Message(i + 1)
       End With
    Next i
    frmMain.rtbChat.SelText = vbCrLf
End Sub

Public Function KillNull(ByVal text As String) As String
    Dim i As Integer
    i = InStr(1, text, Chr(0))
    If i = 0 Then
        KillNull = text
        Exit Function
    End If
    KillNull = Left(text, i - 1)
End Function

Public Sub Send(ByVal strText As String, Socket As Winsock, Optional Extra As Boolean)
If Not frmMain.wsBnet.State = sckConnected Then Exit Sub
    With PBuffer
        .InsertNTString strText
        .SendPacket &HE
    End With
        AddChat D2White, ":: " & BNET.TrueUsername & " :: " & strText
End Sub

Public Function WriteStuff(appname As String, key As String, sString As String, Optional strIni As String) As Boolean
Dim sFile As String
Dim L As Long
WriteStuff = False
On Error GoTo WriteStuff_Error
If strIni = vbNullString Then
    sFile = App.Path & "\config.ini"
Else
    sFile = App.Path & "\" & strIni
End If
L = WritePrivateProfileString(appname, key, sString, sFile)
WriteStuff = True

WriteStuff_Error:
If Err.Number <> 0 Then
'MessageBox Err.Description
End If
End Function

Public Function GetStuff(appname As String, key As String, Optional strIni As String) As String
Dim sFile As String
Dim sDefault As String
Dim lSize As Integer
Dim L As Long
Dim sUser As String
sUser = Space$(128)
lSize = Len(sUser)
If strIni = vbNullString Then
    sFile = App.Path & "\config.ini"
Else
    sFile = strIni
End If
sDefault = vbNullString
L = GetPrivateProfileString(appname, key, sDefault, sUser, lSize, sFile)
sUser = Mid(sUser, 1, InStr(sUser, Chr(0)) - 1)
GetStuff = sUser
End Function

Public Function DoAddToSendList(text As String)

End Function

Public Function ClearBuffers()
    frmMain.rtbChat.text = vbNullString
    AddChat D2White, "Every buffer cleared."
End Function


Public Function MKL(X As String) As Long
    If Len(X) < 4 Then Exit Function
    CopyMemory MakeLong, ByVal X, 4
End Function

Public Function MKI(X As String) As Integer
    If Len(X) < 2 Then Exit Function
    CopyMemory MakeInt, ByVal X, 2
End Function
