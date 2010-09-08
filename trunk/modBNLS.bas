Attribute VB_Name = "modBNLS"
Public Success As Boolean

Public statstring As String

Public cookie As Long

Public versioncode As Long

Private Function GetDWORD(data As String) As Long
Dim lReturn As Long
    Call CopyMemory(lReturn, ByVal data, 4)
    GetDWORD = lReturn
End Function

Private Sub InitCRC32()
    Dim i As Long, J As Long, K As Long, XorVal As Long
    
    Static CRC32Initialized As Boolean
    If CRC32Initialized Then Exit Sub
    CRC32Initialized = True
    
    For i = 0 To 255
        K = i
        
        For J = 1 To 8
            If K And 1 Then XorVal = CRC32_POLYNOMIAL Else XorVal = 0
            If K < 0 Then K = ((K And &H7FFFFFFF) \ 2) Or &H40000000 Else K = K \ 2
            K = K Xor XorVal
        Next
        
        CRC32Table(i) = K
    Next
End Sub

Private Function CRC32(ByVal data As String) As Long
    Dim i As Long, J As Long
    
    Call InitCRC32
    
    CRC32 = &HFFFFFFFF
    
    For i = 1 To Len(data)
        J = CByte(Asc(Mid(data, i, 1))) Xor (CRC32 And &HFF&)
        If CRC32 < 0 Then CRC32 = ((CRC32 And &H7FFFFFFF) \ &H100&) Or &H800000 Else CRC32 = CRC32 \ &H100&
        CRC32 = CRC32 Xor CRC32Table(J)
    Next
    
    CRC32 = Not CRC32
End Function

Public Function BNLSChecksum(ByVal Password As String, ByVal ServerCode As Long) As Long
    BNLSChecksum = CRC32(Password & Right("0000000" & Hex(ServerCode), 8))
    If debugmode = 1 Then
        AddChat HEXPINK, "BNLS CheckSum: " & CRC32(Password & Right("0000000" & Hex(ServerCode), 8)) & " As Password.."
    End If
End Function

Public Function GetBNLSByte() As Long
Select Case BNET.Product
    Case "RATS"
        GetBNLSByte = &H1
    Case "PXES"
        GetBNLSByte = &H2
    Case "PX2D"
        GetBNLSByte = &H5
    Case "VD2D"
        GetBNLSByte = &H4
    Case "NB2W"
        GetBNLSByte = &H3
    Case "3RAW"
        GetBNLSByte = &H7
    Case "PX3W"
        GetBNLSByte = &H8
    Case Else
        AddChat "BNLS Can't find your game type."
End Select
    If debugmode = 1 Then
        AddChat HEXPINK, "BNLS Product Byte: " & GetBNLSByte
    End If
End Function

Public Sub ParseBNLS(ByVal data As String)
Select Case Asc(Mid(data, 3, 1))
'(WORD)      Message Length, including this header
'(BYTE)      Message ID
'(VOID)      Message Data
    Case &H1A
            Dim pB As New Buffer
         With pB
            .SetBuffer data
            .Skip 3
            Success = .GetBoolean
                'AddChat D2Purple, Success
            version = .GetDWORD
                'AddChat D2Purple, version
            CheckSum = .GetDWORD
                'AddChat D2Purple, CheckSum
            statstring = .GetSTRING
                'AddChat D2Purple, statstring
            cookie = .GetDWORD
                'AddChat D2Purple, cookie
            versioncode = .GetDWORD
                'AddChat D2Purple, versioncode
         End With
         
'(DWORD) Client Token
'(DWORD) EXE Version
'(DWORD) EXE Hash
'(DWORD) Number of CD-keys in this packet
'(BOOLEAN) Spawn CD-key

'For Each Key:'

   ' (DWORD) Key Length
   ' (DWORD) CD-key's product value
   ' (DWORD) CD-key's public value
   ' (DWORD) Unknown (0)
   ' (DWORD) [5] Hashed Key Data


'(STRING) Exe Information
'(STRING) CD-Key owner name

        If varproduct = "PX2D" Or varproduct = "PX3W" Then
            With PBuffer
                .InsertDWORD &H0
                .InsertBYTE &H2
                .InsertDWORD &H1
                .InsertDWORD Servers
                .InsertNTString frmConfigBNET.txtCDKey.text
                .InsertNTString frmConfigBNET.txtCDKey2.text
                .SendBNLSPacket &HC
            End With
        Else
            With PBuffer
               .InsertDWORD Servers
               .InsertNTString frmConfigBNET.txtCDKey.text
               .SendBNLSPacket &H1
            End With
        End If
         
    Case &H9
'(BOOL) Success*
'(DWORD) Version.
'(DWORD) Checksum.
'(STRING) Version check stat string.'
'(DWORD) Cookie.
'(DWORD) The latest version code for this product.

        
        If varproduct = "PX2D" Or varproduct = "PX3W" Then
            With PBuffer
                .InsertDWORD &H0
                .InsertBYTE &H2
                .InsertDWORD &H1
                .InsertDWORD Servers
                .InsertNTString frmConfigBNET.txtCDKey.text
                .InsertNTString frmConfigBNET.txtCDKey2.text
                .SendBNLSPacket &HC
            End With
        Else
            With PBuffer
               .InsertDWORD Servers
               .InsertNTString frmConfigBNET.txtCDKey.text
               .SendBNLSPacket &H1
            End With
        End If
    Case &H4
        With PBuffer
            .InsertNonNTString Mid$(data, 4)
            .InsertNTString BNET.username
            .SendPacket &H52
        End With
    Case &H2
       With PBuffer
            .InsertNonNTString Mid(data, 4)
            .InsertNTString BNET.username
            .SendPacket &H53
       End With
    Case &H3
        With PBuffer
            .InsertNonNTString Mid(data, 4)
            .SendPacket &H54
        End With
    Case &HC
        CdkeyHash = Mid(data, 18, 36)
        Cdkey2Hash = Mid(data, 58, 36)
        GTC = Val("&H" & StrToHex(StrReverse(Mid(data, 14, 4))))
        GTC = CLng(GTC)
        Send0x51
    Case &H1
        If debugmode = 1 Then
            AddChat HEXPINK, "BNLS PACKET 0x1"
        End If
        '(BOOLEAN) Result
        '(DWORD) Client Token
        '(DWORD[9]) CD key data for SID_AUTH_CHECK
            With pB
                .SetBuffer data
                .Skip 3
                
                If Not .GetBoolean Then
                    Exit Sub
                End If
                
                GTC = .GetDWORD
                CdkeyHash = .GetRaw(36)
                    If debugmode = 1 Then
                        AddChat HEXPINK, "CD Key HASH???"
                    End If
            End With
            'CdkeyHash = Mid(Data, 12)
            'GTC = Val("&H" & StrToHex(StrReverse(Mid(Data, 8, 4))))
            'GTC = CLng(GTC)
            
            With PBuffer
                .InsertDWORD GTC
                'AddChat D2Orange, "GTC", GTC
                .InsertDWORD version
                'AddChat D2Orange, "Version", version
                .InsertDWORD CheckSum
                'AddChat D2Orange, "CheckSum", CheckSum
                'If BNET.Product = "PX2D" Or BNET.Product = "PX3W" Then
                '    .InsertDWORD &H2
                'Else
                    .InsertDWORD &H1
                'End If
                .InsertDWORD &H0
                .InsertNonNTString CdkeyHash
                'AddChat D2Orange, "CD Key Hash", CdkeyHash
                'AddChat D2Orange, "CD Key Hash length", Len(CdkeyHash)
                'If BNET.Product = "PX2D" Or BNET.Product = "PX3W" Then
                '   .InsertNonNTString Cdkey2Hash
                'nd If
                .InsertNTString statstring
                .InsertNTString BNET.username
                .SendPacket &H51
            End With
           ' Send0x51
        Case &HE
            Dim key As Long, key2 As Long
            key2 = GetDWORD(Mid(data, 4, 4))
            key = BNLSChecksum("Invigoration", key2)
            With PBuffer
                .InsertDWORD key
                .SendBNLSPacket &HF
            End With
    Case &HB ' WHERE THE CONNECTION IS FREEZING?
        If debugmode = 1 Then
            AddChat HEXPINK, "BNLS Send 0xB START"
        End If
        If HType = 1 Then
        If debugmode = 1 Then
            AddChat HEXPINK, "HType = 1"
        End If
          CB = CB + 1
                If debugmode = 1 Then
                    AddChat HEXPINK, "CB: " & CB
                End If
            If CB = 1 Then
                If debugmode = 1 Then
                    AddChat HEXPINK, "CB = 1 "
                End If
                hash(0) = PBuffer.MakeDWORD(GTC)
                hash(1) = PBuffer.MakeDWORD(Servers)
                hash(2) = Mid(data, 4, Len(data) - 3)
                With PBuffer
                    .InsertDWORD &H1C
                    .InsertDWORD &H1
                    .InsertNonNTString hash(0) & hash(1) & hash(2)
                    .SendBNLSPacket &HB
                End With
            ElseIf CB = 2 Then
                If debugmode = 1 Then
                    AddChat HEXPINK, "CB = 2 "
                End If
                With PBuffer
                    If SPass = True Then
                        .InsertDWORD GTC
                        .InsertDWORD Servers
                        .InsertNonNTString Mid(data, 4, Len(data) - 3)
                        .InsertNTString BNET.username
                        .SendPacket &H3A
                        SPass = False
                        CB = 0
                    Else
                        .InsertDWORD GTC
                        .InsertNonNTString Mid(data, 4, Len(data) - 3)
                        .InsertNTString BNET.Realm
                        .SendPacket &H3E
                        CB = 0
                    End If
                End With
            End If
        ElseIf HType = 2 Then
            With PBuffer
                .InsertNonNTString Mid(data, 4, Len(data) - 3)
                .InsertNTString BNET.username
                .SendPacket &H2A
            End With
        ElseIf HType = 3 Then
            Static Hash2 As String
            CB = CB + 1
            If CB = 1 Then
                hash(0) = PBuffer.MakeDWORD(GTC)
                hash(1) = PBuffer.MakeDWORD(Servers)
                hash(2) = Mid(data, 4, Len(data) - 3)
                With PBuffer
                    .InsertDWORD &H1C
                    .InsertDWORD &H1
                    .InsertNonNTString hash(0) & hash(1) & hash(2)
                    .SendBNLSPacket &HB
                End With
            End If
            If CB = 2 Then
                Hash2 = Mid(data, 4, Len(data) - 3)
                With PBuffer
                    .InsertDWORD "&h" & Len(BNET.NewPass)
                    .InsertDWORD &H0
                    .InsertNonNTString BNET.NewPass
                    .SendBNLSPacket &HB
                End With
            End If
            If CB = 3 Then
                With PBuffer
                    .InsertDWORD GTC
                    .InsertDWORD Servers
                    .InsertNonNTString Hash2
                    .InsertNonNTString Mid(data, 4, Len(data) - 3)
                    .InsertNTString BNET.username
                    .SendPacket &H31
                End With
            End If
        End If
    Case &HF
        PBuffer.InsertDWORD GetBNLSByte()
        PBuffer.SendBNLSPacket &H10
        If BNET.Product = "3RAW" Then
            With PBuffer
                .InsertDWORD &H2
                .SendBNLSPacket &HD
            End With
        End If
    Case &H10
        VerByte = GetDWORD(Mid(data, 8, 4))
        frmMain.wsBnet.Close
        frmMain.wsBnet.Connect BNET.BattlenetServer, 6112
End Select
End Sub
