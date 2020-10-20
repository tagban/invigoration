Attribute VB_Name = "modCommands"
Global info As String
Global Canada As Integer
Global debugmode As Integer
Global idleMessage As String
Global idletimeset As Integer
Global fudd As Integer
Global uptimesec As Integer
Global uptimemin As Integer
Global uptimehour As Integer
Global uptimedays As Integer
Global uptimeweek As Integer
Global uptimemonth As Integer
Global moo As Integer
Global leetspeak As Integer
Global LastW As String
Global LastCW As String
Global LastM As String
Global LastSW As String
Global LastSM As String
Global BanCount As Integer
Global acceptinvites As Boolean
Global KickCount As Integer
Global JoinCount As Integer
Global beforetext As String
Global postpend As String
Global targetuser As String
Global targetusername As String
Public Sub ParseCommand(ByVal Message As String, username As String, Optional Inbot As Boolean = False)
    If LCase$(Message) = "?trigger" Then Message = BNET.Trigger & "trigger"
        'if the command doesn't have a trigger, its not a command, so exit
    If Left$(Message, 1) <> BNET.Trigger And Left$(Message, 1) <> "/" Then Exit Sub
        Message = Mid$(Message, 2) 'strip the trigger
    
    Dim Command As String, Rest As String
    If Len(Message) > 0 Then
        Command = Split(Message, Space(1))(0) 'get the first word to be the command
        'if there is more to the command, place it in the "rest" variable
    If Len(Message) > Len(Command) + 1 Then Rest = Trim(Mid$(Message, Len(Command) + 1))
        'until access system is setup, only parse commands from bot master
    If (LCase(username) <> LCase(BNET.BotMaster)) And Not Inbot Then Exit Sub


    End If
''''''''''''''''''''''''''''''''''''''' Commands ''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim i As Integer
    Dim b As Boolean
    Dim SayMessage As String
    Dim Extra As String

    Select Case LCase$(Command)
    Case "idle"
        Select Case Rest
            Case "off"
                frmMain.IdleTimer.Enabled = False
                IdleTime = 0
                Send LastCW & "Idle message turned off.", frmMain.wsBnet
                LastCW = vbNullString
            Case Else
                tmp = Split(Rest, Space(1), 2)
                idletimeset = tmp(0)
                If tmp(1) = "uptime" Then
                    Send LastCW & "Idle message set.[Uptime]", frmMain.wsBnet
                    LastCW = vbNullString
                    idleMessage = "/me has been online for: " & uptimemonth & " months, " & uptimeweek & " weeks, " & uptimedays & " days, " & uptimehour & " hours, " & uptimemin & " minutes, and " & uptimesec & " seconds."
                ElseIf tmp(1) = "ver" Then
                    Send LastCW & "Idle message set.[Version]", frmMain.wsBnet
                    LastCW = vbNullString
                    idleMessage = "/me is an Invigoration v" & botver & " by Tagban - http://www.BNET.cc"
                Else
                    idleMessage = tmp(1)
                    Send LastCW & "Idle message set.", frmMain.wsBnet
                    LastCW = vbNullString
                End If
                    frmMain.IdleTimer.Enabled = True
            End Select
    Case "disconnect", "disc"
    AddChat D2White, "Disconnecting..."
        frmMain.wsBnet.Close
        frmMain.wsBnls.Close
    Case "colors", "color"
    AddChat D2Orange, "Chat Colors Help:"
    AddChat D2White, "First use Alt+0160 infront and select a color from the list below:"
    AddChat D2White, "      r= rred w, w= wwhite, q= qgray w, g= ggreen w, y= yyellow w"
    AddChat D2White, "      b= bblue w, o= oorange w, c= ccyan(light blue) w, p= ppurple w"
    AddChat D2White, "      l= llight yellow w, e= ebeige w, k= kpink w"
    Case "reconnect"
    AddChat D2White, "Reconnecting, hold on tight!"
        With frmMain.lstChannel
            .ListItems.Clear
        End With
        frmMain.wsBnet.Close
        frmMain.wsBnls.Close
        frmMain.wsBnls.Close
        frmMain.wsBnls.Connect BNET.BNLSServer, 9367
    Case "hex", "h"
           Send Chr(163) & StrToHex(Rest), frmMain.wsBnet
           LastCW = vbNullString
    Case "invigencrypt", "encrypt", "ie", "i"
           Send Chr(149) & InvigEncrypt(Rest & "-"), frmMain.wsBnet
           LastCW = vbNullString
    Case "sysinfo"
        Send "Invigoration running on: " & OSVersion & ". -- Runtime: " & WindowsRunTime & " hours.", frmMain.wsBnet
        LastCW = vbNullString
    Case "ver"
        Send LastCW & "/me is an Invigoration v" & botver & " by Tagban - http://www.BNET.cc/", frmMain.wsBnet
        LastCW = vbNullString
    Case "uptime"
        Send LastCW & "/me has been online for: " & uptimemonth & " months, " & uptimeweek & " weeks, " & uptimedays & " days, " & uptimehour & " hours, " & uptimemin & " minutes, and " & uptimesec & " seconds.", frmMain.wsBnet
        LastCW = vbNullString
    Case "about"
        Send LastCW & "Invigoration was written in Visual Basic by Tagban Since 2004 -- http://www.bnet.cc", frmMain.wsBnet
        LastCW = vbNullString
    Case "say"
        Send Rest, frmMain.wsBnet
        LastCW = vbNullString
    Case "bancount"
        If BanCount = 0 Then
            Send LastCW & "Noone has been banned since I entered this channel..", frmMain.wsBnet
            LastCW = vbNullString
        ElseIf BanCount = 1 Then
            Send BanCount & " user has been banned since I joined this channel.", frmMain.wsBnet
            LastCW = vbNullString
        Else
            Send BanCount & " users have been banned since I joined this channel.", frmMain.wsBnet
            LastCW = vbNullString
        End If
    Case "kickcount"
        If KickCount = 0 Then
            Send LastCW & "Noone has been kicked since I entered this channel..", frmMain.wsBnet
            LastCW = vbNullString
        ElseIf KickCount = 1 Then
            Send KickCount & " user has been kicked since I joined this channel.", frmMain.wsBnet
            LastCW = vbNullString
        Else
            Send KickCount & " users have been kicked since I joined this channel.", frmMain.wsBnet
            LastCW = vbNullString
        End If
    Case "joincount"
        If JoinCount = 0 Then
            Send LastCW & "Noone has entered since I got here..", frmMain.wsBnet
            LastCW = vbNullString
        ElseIf JoinCount = 1 Then
            Send JoinCount & " user has joined the channel since I've been here.", frmMain.wsBnet
            LastCW = vbNullString
        Else
            Send JoinCount & " users have joined the channel since I've been here.", frmMain.wsBnet
            LastCW = vbNullString
        End If
    Case "ban"
        Send "/ban " & Rest, frmMain.wsBnet
    Case "join"
        With PBuffer
            .SendPacket &H10
            .InsertDWORD 2
            .InsertNTString Rest
            .SendPacket &HC
            'AddChat D2Orange, Rest
        End With
    Case "user"
        rtbsendbnet.SelColor = D2White
        targetuser = Rest & " : "
        targetusername = Rest
        
 '       AddChat D2White, "User focus set on: " & Rest & vbNullString, vbNewLine, D2Orange
    Case "useroff"
        targetuser = vbNullString
        'AddChat D2White, "User focus removed."
    Case "prepend", "pre"
        beforetext = Rest & vbNullString
        AddChat D2MedBlue, beforetext & " will be displayed before each send."
        AddChat D2MedBlue, "To deactivate, type: '/prepend' with nothig after it."
    Case "postpend", "post"
        AddChat D2MedBlue, Rest & " will be displayed after each send."
        AddChat D2MedBlue, "To deactivate, type: '/postpend' with nothig after it."
        postpend = Rest & vbNullString
    Case "join"
        Send "/join " & Rest, frmMain.wsBnet
    Case "setmaster"
        BNET.BotMaster = Rest
        Send LastCW & " Bot master changed!", frmMain.wsBnet
        SaveConfig
    Case "sethome"
        BNET.HomeChannel = Rest
        Send LastCW & " Home channel changed!", frmMain.wsBnet
        SaveConfig
    Case "setusername"
        BNET.username = Rest
        Send LastCW & " Login username changed!", frmMain.wsBnet
        SaveConfig
    Case "setpass"
        BNET.Password = Rest
        Send LastCW & " Password login changed!", frmMain.wsBnet
        SaveConfig
    Case "setserver"
        BNET.BattlenetServer = Rest
        Send LastCW & " Server changed!", frmMain.wsBnet
        SaveConfig
    Case "settrigger"
        BNET.Trigger = Rest
        Send LastCW & " Bot trigger changed!", frmMain.wsBnet
        SaveConfig
    Case "kick"
        Send "/kick " & Rest & " [InvigOp Alpha]", frmMain.wsBnet
    Case "trigger"
        Send LastCW & "The bot's trigger is: " & BNET.Trigger, frmMain.wsBnet
        LastCW = vbNullString
    Case "lastm", "lastMessage", "lastw", "last", "lrm", "lrw"
       Send LastCW & "Last recieved whisper From: " & LastW & " :: " & LastM, frmMain.wsBnet
       LastCW = vbNullString
    Case "lastsm", "lastsentMessage", "lastsw", "lastsend", "lsm", "lsw"
       Send LastCW & "Last sent whisper To: " & LastSW & " :: " & LastSM, frmMain.wsBnet
       LastCW = vbNullString
   '''''''''''''''''
    Case "quit", "unload"
        'Exits the Program, uses less ram to do so ^^
        Unload frmMain
        End
    Case "canada"
            If Canada = 0 Then
                Canada = 1
                Send LastCW & "Canada Mode enabled.", frmMain.wsBnet
                LastCW = vbNullString
            Else
                Canada = 0
                Send LastCW & "Canada Mode disabled.", frmMain.wsBnet
            End If
        'CanadaMode
    Case "accept"
            If acceptinvites Then
                acceptinvites = False
                Send LastCW & "Invite Auto-Accept Disabled.", frmMain.wsBnet
                LastCW = vbNullString
            Else
                acceptinvites = True
                Send LastCW & "Clan invitations will be accepted automatically.", frmMain.wsBnet
            End If
        'Accept Invites
        Case "debug"
            If debugmode = 0 Then
                debugmode = 1
                AddChat D2MedBlue, "Debug Mode enabled."
                LastCW = vbNullString
            Else
                debugmode = 0
                AddChat D2MedBlue, "Debug Mode disabled."
            End If
        'CanadaMode
    Case "say"
        Select Case Rest
            Case Message
                Send Message, frmMain.wsBnet
        End Select
        'CanadaMode
    Case "leetspeak"
            If leetspeak = 0 Then
                leetspeak = 1
                Send LastCW & "Leet Speak enabled.", frmMain.wsBnet
                LastCW = vbNullString
            Else
                leetspeak = 0
                Send LastCW & "Leet Speak off.", frmMain.wsBnet
                LastCW = vbNullString
            End If
        'Leet Speak Mode
    Case "fudd"
            If fudd = 0 Then
                fudd = 1
                Send LastCW & "Elmer Fudd mode enabled.", frmMain.wsBnet
                LastCW = vbNullString
            Else
                fudd = 0
                Send LastCW & "Elmer Fudd mode Off", frmMain.wsBnet
                LastCW = vbNullString
            End If
        'Elmer Fudd Mode
    Case "moo"
            If moo = 0 Then
                moo = 1
                Send LastCW & "Moooooooooooooooo mode engaged!", frmMain.wsBnet
                LastCW = vbNullString
            Else
                moo = 0
                Send LastCW & "Cows are off...", frmMain.wsBnet
                LastCW = vbNullString
            End If
        'Elmer Fudd Mode

    Case "home", "gohome", "homechan", "homechannel"
        PBuffer.SendPacket &H10
        PBuffer.InsertDWORD 2
        PBuffer.InsertNTString BNET.HomeChannel
        PBuffer.SendPacket &HC
    Case "rejoin"
        PBuffer.SendPacket &H10
        PBuffer.InsertDWORD 2
        PBuffer.InsertNTString BNET.CurrentChan
        PBuffer.SendPacket &HC
    Case "w", "m", "whisper", "unignore", "unsquelch", "Message", "help", "clan", "where", "c", "Message", "squelch", "ign,e", "f", "friend", "friends", "?", "help", "dnd", "options", "o", "emote", "me", "channel", "who", "whoami", "whois", "whereis", "beep", "designate", "mail", "time", "unban", "users", "stats", "set-email", "nobeep"
        If Inbot Then Send LastCW & "/" & Message, frmMain.wsBnet
        'Empty Command NOTHING NULL'
    Case Else
         'AddChat D2Red, "This is not a valid command, or is not currently functioning properly. If you feel this message is in error, please report it as a bug on Invigoration's website. http://invigoration.bnet.cc"
         '' The above only happens if the command is BLANK ''
    End Select
   
End Sub
