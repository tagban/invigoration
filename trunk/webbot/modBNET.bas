Attribute VB_Name = "modBNET"
Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
Private Server As String
Private Bleh2 As String
Private asplt() As String, z As Integer, tmpnews As Long, splti() As String
Private mpqn As Long, hash As String, MPQName As String
Private spltns() As String, spltn() As String, strss As String, Split0B() As String




Public Function MakeServer(data As String) As String
    MakeServer = CLng("&H" & ToHex(Mid(data, 1, 1))) & "." & CLng("&H" & ToHex(Mid(data, 2, 1))) & "." & CLng("&H" & ToHex(Mid(data, 3, 1))) & "." & CLng("&H" & ToHex(Mid(data, 4, 1)))
End Function

Public Sub Send0x51()
    With PBuffer
        .InsertDWORD GTC
        .InsertDWORD version
        .InsertDWORD CheckSum
        .InsertDWORD &H1
        .InsertDWORD &H0
        .InsertNonNTString CdkeyHash
        .InsertNTString ExeInfo
        .InsertNTString BNET.username
        .SendPacket &H51
    End With
End Sub
Public Sub JoinHome()
    With PBuffer
        .SendPacket &H10
        .InsertDWORD 2
        .InsertNTString BNET.HomeChannel
        .SendPacket &HC
    End With
End Sub
Public Sub Send0x50()

    With PBuffer
        .InsertDWORD &H0
        .InsertNonNTString "68XI" & BNET.Product
        .InsertDWORD VerByte
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertDWORD &H480
        .InsertDWORD &H1033
        .InsertDWORD &H1033
        .InsertNTString "USA"
        .InsertNTString "United States"
        .SendPacket &H50
            .InsertDWORD &H0
            .SendPacket &H25

    End With
End Sub
Public Sub Send0x07()
    With PBuffer
        .SendPacket &H7
        On Error GoTo Error:
    End With
If frmMain.wsBnet.State = sckConnected Then
       With PBuffer
        .SendPacket &H7
        On Error GoTo Error:
    End With
ElseIf frmMain.wsBnet.State = sckClosed Then

Else
Error:
End If
End Sub

Public Sub RequestBnetNews()
    With PBuffer
        .InsertDWORD &H0
        .SendPacket &H46
    End With
End Sub

Public Function GetChannelType(ByVal Flags As Long) As String
    Select Case Flags
        Case CHNFLAG_VOID
            GetChannelType = "DeVoid Channel"
        Case CHNFLAG_MOD
            GetChannelType = "Moderated channel"
        Case CHNFLAG_STS
            GetChannelType = "Product-Specific Channel"
        Case CHNFLAG_PUB
            GetChannelType = "Public Channel"
        Case Else
            GetChannelType = "Private Channel"
    End Select
End Function

Public Sub ParseBnet(data As String)
Dim PacketId As Byte, RP As Long, outb As String
Dim accountHash As String
Dim Product As String

    
    PacketId = Asc(Mid(data, 2, 1))

    Select Case PacketId
        Case &H0
        '' Liquid sex?
        Case &H34
            HType = 1
            With PBuffer
                .InsertDWORD &H8
                .InsertDWORD &H0
                .InsertNonNTString "password"
                .SendBNLSPacket &HB
            End With
        Case &H3E
                P1 = Mid(data, 5, 16)
                Server = Mid(data, 17, 8)
                Bleh2 = Mid(Server, 5, 4)
                AddChat D2Green, "Current realm server: " & MakeServer(Bleh2)
                P2 = Mid(data, 29, 48)
        Case &H51
            Select Case GetWORD(Mid(data, 5, 2))
                Case &H0 'ToDo: Remove BNLS from password/Username/CD Key equation.
                    AddChat D2Green, "Version Check + CDKey Check"
                    With PBuffer
                        If BNET.Product = "3RAW" Or BNET.Product = "PX3W" Then
                            .InsertNTString BNET.username
                            .InsertNTString BNET.Password
                            .SendBNLSPacket &H2
                        Else
                                .InsertNonNTString "bnet"
                            .SendPacket &H14
                            .SendPacket &H2D
                            If Cpass = False Then
                                HType = 1
                                .InsertDWORD Len(BNET.Password)
                                .InsertDWORD &H0
                                .InsertNonNTString BNET.Password
                                .SendBNLSPacket &HB
                                SPass = True
                            Else
                                Cpass = False
                                HType = 3
                                .InsertDWORD Len(BNET.Password)
                                .InsertDWORD &H0
                                .InsertNonNTString BNET.Password
                                .SendBNLSPacket &HB
                            End If
                        End If
                    End With
                Case &H100
                    AddChat D2Red, "Game version out of date."
                Case &H101
                    AddChat D2Red, "Invalid Game Version. What the FRAK?! Check BNLS/JBLS Server and try another one?"
                Case &H102
                    AddChat D2Red, "Game Version needs downgraded. WHAT THE FRAK!? Downgraded!??!?!? Fawk you!"
                Case &H200
                    AddChat D2Red, "Invalid CD Key."
                Case &H201
                    AddChat D2Red, "CDKey is in use."
                Case &H203
                    AddChat D2Red, "Incorrect cdkey for this product."
                    AddChat D2Red, "Please check your key and/or game."
                Case &H202
                    AddChat D2Red, "Current cdkey is banned from Battle.net."
            End Select
    Case &H25 'Negative Ping 0x25 No Response!
                        PBuffer.InsertDWORD &H0
                        PBuffer.SendPacket &H25
    Case &H66

    Case &H67

    Case &H68

    Case &H53
            If Asc(Mid$(data, 5, 4)) = &H1 Then
            'Account Doesn't Exist
                If BNET.Product = "3RAW" Then
                    AddChat D2Red, "Account doesn't exist, unable to create using Warcraft III."
                Else
                    AddChat D2Red, "Account doesn't exist, attempting creation."
                        With PBuffer
                            .InsertNTString BNET.username
                            .InsertNTString BNET.Password
                            .SendBNLSPacket &H4
                        End With
                End If
            Else
                With PBuffer
                    .InsertNonNTString Mid(data, 9, 64)
                    .SendBNLSPacket &H3
                End With
            End If
        Case &H52
            Select Case GetWORD(Mid(data, 5, 2))
                Case &H0
                    With PBuffer
                        .InsertNTString BNET.username
                        .InsertNTString BNET.Password
                        .SendBNLSPacket &H2
                    End With
                Case Else
            End Select
        Case &H54
            Select Case GetWORD(Mid(data, 5, 2))
            Case &H0 ' Originally set as &H0
                AddChat D2Green, "Battle.net Logon Passed!"
                ConnectTry = 0
                With PBuffer
                    .InsertNTString BNET.username
                    .InsertBYTE 0
                    .SendPacket &HA
                    .InsertNonNTString BNET.Product
                    .SendPacket &HB
                    .InsertDWORD 1
                    .InsertNTString "L"
                    .SendPacket &HC
                End With
            Case &H2
                AddChat D2Red, "Battle.net Logon Failed!"
                If AttemptedC = False Then
                With PBuffer
                    .InsertNTString BNET.username
                    .InsertNTString BNET.Password
                    .SendBNLSPacket &H4
                End With
                ElseIf AttemptedC = True Then
                    AddChat D2Orange, "Tried to make account!!"
                      '''''''''''''''''''' ACCOUNT CREATION ''''''''''''''''''''''
                    AddChat D2Orange, "Account created! :)"
                    AddChat D2Green, "Connecting with your NEW account!"
                    dctime = 0
                    AddChat D2Green, "Battle.net Login Server Connecting to " & BNET.BNLSServer & "..."
                    frmMain.wsBnls.Close
                    frmMain.wsBnls.Connect BNET.BNLSServer, 9367
                    frmMain.Caption = "BNET.cc - [ Connecting to: " & BNET.BNLSServer & " ]"
                    connectstatus = True
                End If
            Case &HE    'Email Address, currently ignored
                        'TODO: support email registration
                AddChat D2Green, "Battle.net Logon Passed!"
                AddChat D2Orange, "Battle.net requested an e-mail address. BNET.cc currently doesn't have the functionality to handle this and so it bypassed it."
                ConnectTry = 0
                With PBuffer
                    .InsertNTString BNET.username
                    .InsertBYTE 0
                    .SendPacket &HA
                    .InsertNonNTString BNET.Product
                    .SendPacket &HB
                    .InsertDWORD 1
                    .InsertNTString "L"
                    .SendPacket &HC
                End With
            End Select
        Case &H50 'Updated 9/12/2010 - Tagban
           Dim pB As New Buffer
           With pB
                .SetBuffer data
                .Skip 8
                
                Servers = .GetDWORD
                Dim UDP As Long
                UDP = .GetDWORD
                Dim timestamp As String
                timestamp = .GetFILETIME
                MPQName = .GetSTRING
                hash = .GetSTRING
           End With
            With PBuffer
                .InsertDWORD GetBNLSByte()
                .InsertDWORD &H0
                .InsertDWORD &H1 'cookie
                .InsertDWORD CLng(Split(timestamp, " ")(1)) 'timestamp
                .InsertDWORD CLng(Split(timestamp, " ")(0)) 'timestamp
                .InsertNTString MPQName
                .InsertNTString hash
                .SendBNLSPacket &H1A
            End With
                      
        Case &H46 ' get news reply .... Defunct??
                spltns() = Split(StrToHex(Mid(data, 22)), "0A")
                For tmpnews = 0 To UBound(spltns) - 1
                    AddChat D2Purple, "News: " & HexToStr(spltns(tmpnews))
                Next tmpnews
                Exit Sub
            Erase spltns()
            
        Case &H31 ' password change reply
            If InStr(data, Chr(&H1)) Then
                AddChat D2MedBlue, "Password changed, logging on."
            Else
                AddChat D2MedBlue, "Password not changed."
            End If
            
        Case &H3A ' account login reply
            Select Case Asc(Mid(data, 5, 1))
                Case &H1
                    AddChat D2Red, "Battle.net Logon failed!"
                    If AttemptedC = False Then
                        With PBuffer
                            HType = 2
                            .InsertDWORD Len(BNET.Password)
                            .InsertDWORD &H0
                            .InsertNonNTString BNET.Password
                            .SendBNLSPacket &HB
                        End With
                    ElseIf AttemptedC = True Then
                    End If
                Case &H2
                    AddChat D2Red, "Battle.net Logon failed, due to incorrect password."
                Case &H6 ' Added 9/14/2010 -- Tagban
                    AddChat D2Red, "YOUR ACCOUNT WAS CLOSED OR BANNED..."
                Case &H0
                    AddChat D2Green, "Battle.net Logon Passed!"
                    ConnectTry = 0
                    With PBuffer
                        If LRealm = True And BNET.Product = "VD2D" Then 'D2 Specialized Login
                            .InsertDWORD &H0
                            .InsertDWORD &H0
                            .InsertBYTE &H0
                            .SendPacket &H34
                        Else
                            .InsertNTString BNET.username
                            .InsertBYTE 0
                            .SendPacket &HA
                            .InsertNonNTString BNET.Product
                            .SendPacket &HB
                            .InsertDWORD 2
                            .InsertNTString BNET.HomeChannel
                            .SendPacket &HC
                        End If
                    End With
                Case Else
            End Select
            Exit Sub
        Case &H59
                AddChat D2MedBlue, "Please set email to this account via the actual game window. BNET.cc doesn't register accounts."
        Case &HF
            frmMain.ChatBot.DispatchMessage data
            Exit Sub
        Case &H2A
                AddChat D2Orange, "Account created! :)"
                AddChat D2Green, "Connecting with your NEW account!"
                frmMain.wsBnet.Close
                AddChat D2Green, "Battle.net Login Server Connecting to " & BNET.BNLSServer & "..."
                frmMain.wsBnls.Close
                frmMain.wsBnls.Connect BNET.BNLSServer, 9367
                frmMain.Caption = "BNET.cc - [ Connecting to: " & BNET.BNLSServer & " ]"
                'Case &H26
        
        Case &HA 'Properly name game codes for better user experience during login process.
                spltn() = Split(data, Chr(0), 5)
                BNET.TrueUsername = spltn(3)
                If BNET.Product = "RATS" Then
                    Product = "Starcraft"
                ElseIf BNET.Product = "PXES" Then
                    Product = "Starcraft: Brood War"
                ElseIf BNET.Product = "NB2W" Then
                    Product = "Warcraft 2 Battle.net Edition"
                ElseIf BNET.Product = "3RAW" Then
                    Product = "Warcraft 3"
                ElseIf BNET.Product = "VD2D" Then
                    Product = "Diablo 2"
                ElseIf BNET.Product = "XP2D" Then
                    Product = "D2:LOD"
                ElseIf BNET.Product = "XP3W" Then
                    Product = "War3:TFT"
                End If
                AddChat vbYellow, "Logged on as: " & BNET.TrueUsername & " using " & Product & "."
                frmMain.wsBnls.Close
            Erase spltn()
            Exit Sub
        Case &HB
                Split0B() = Split(Mid(data, 5, Len(data)), Chr(0))
' What is HB ?
            Erase Split0B()
            Exit Sub
        Case &H19
            AddChat D2White, "[&H19]: " & Replace(Mid(data, 9, Len(data)), Chr(0), vbNullString)
            Exit Sub
        Case &H15 ' Ad Change, not really needed, fuck it.
            
        Case &H2D
                splti() = Split(Mid(data, 1, Len(data) - 1), Chr(1), 2)
            Erase splti()
        Case &H75

        Case &H79 ' Currently, needs worked on. 9/12/2010 - Tagban
        Case &H66
        Case &H69
        Case &H78
        Case Else
            If Len(PacketId) = 1 Then
                AddChat D2White, "Unhandled Packet: 0x0" & Hex(PacketId)
            Else
                AddChat D2White, "Unhandled Packet: 0x" & Hex(PacketId)
            End If
            AddChat StrToHex(data)
    End Select
End Sub
