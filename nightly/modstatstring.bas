Attribute VB_Name = "modStatstring"
Option Explicit

Private Sub sprintf(ByRef Source As String, ByVal nText As String, _
    Optional ByVal a As Variant, _
    Optional ByVal b As Variant, _
    Optional ByVal c As Variant, _
    Optional ByVal d As Variant, _
    Optional ByVal E As Variant, _
    Optional ByVal F As Variant, _
    Optional ByVal G As Variant, _
    Optional ByVal H As Variant)
    
    nText = Replace(nText, "%S", "%s")
    Dim i As Byte
    i = 0
    Do While (InStr(1, nText, "%s") <> 0)
        Select Case i
            Case 0
                If IsEmpty(a) Then GoTo TheEnd
                nText = Replace(nText, "%s", a, 1, 1)
            Case 1
                If IsEmpty(b) Then GoTo TheEnd
                nText = Replace(nText, "%s", b, 1, 1)
            Case 2
                If IsEmpty(c) Then GoTo TheEnd
                nText = Replace(nText, "%s", c, 1, 1)
            Case 3
                If IsEmpty(d) Then GoTo TheEnd
                nText = Replace(nText, "%s", d, 1, 1)
            Case 4
                If IsEmpty(E) Then GoTo TheEnd
                nText = Replace(nText, "%s", E, 1, 1)
            Case 5
                If IsEmpty(F) Then GoTo TheEnd
                nText = Replace(nText, "%s", F, 1, 1)
            Case 6
                If IsEmpty(G) Then GoTo TheEnd
                nText = Replace(nText, "%s", G, 1, 1)
            Case 7
                If IsEmpty(H) Then GoTo TheEnd
                nText = Replace(nText, "%s", H, 1, 1)
        End Select
        i = i + 1
    Loop
TheEnd:
    Source = Source & nText
End Sub

Public Sub ParseStatString(ByVal statstring As String, ByRef outbuf As String)
On Error Resume Next
    Dim Values() As String
    Dim cType As String
    
    Select Case Left$(statstring, 4)
        Case "3RAW"
            sprintf outbuf, "WarCraft III: Reign of Chaos ("
            If Len(statstring) > 4 Then
                Exit Sub
            ElseIf Len(statstring) = 4 Then
                strcpy outbuf, "No stats available)"
                Exit Sub
            Else
                strcpy outbuf, "error: " & statstring & ")"
                Exit Sub
            End If
        Case "RHSS"
            Call strcpy(outbuf, "Starcraft Shareware.")
        Case "RATS"
            Values() = Split(Mid$(statstring, 6), Space(1))
            If UBound(Values) <> 8 Then
                Call sprintf(outbuf, "a Starcraft %sbot.", IIf((Values(3) = 1), " (spawn) ", vbNullString))
                Exit Sub
            End If
            If Values(0) > 0 Then
                Call sprintf(outbuf, "Starcraft%s: (%s wins, with a rating of %s on the ladder).", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2), Values(0))
            Else
                Call sprintf(outbuf, "Starcraft%s: (%s wins).", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2))
            End If
        Case "PXES"
            Values() = Split(Mid(statstring, 6), Space(1))
            If UBound(Values) <> 8 Then
                Call sprintf(outbuf, "a Starcraft Brood War %sbot.", IIf((Values(3) = 1), " (spawn) ", vbNullString))
                Exit Sub
            End If
            If Values(0) > 0 Then
                Call sprintf(outbuf, "Starcraft Brood War%s: (%s wins, with a rating of %s on the ladder).", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2), Values(0))
            Else
                Call sprintf(outbuf, "Starcraft Brood War%s: (%s wins).", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2))
            End If
        Case "RTSJ"
            Values() = Split(Mid(statstring, 6), Space(1))
            If UBound(Values) <> 8 Then
                Call sprintf(outbuf, "a Starcraft Japanese %sbot.", IIf((Values(3) = 1), " (spawn) ", vbNullString))
                Exit Sub
            End If
            If Values(0) > 0 Then
                Call sprintf(outbuf, "Starcraft Japanese%s: (%s wins, with a rating of %s on the ladder).", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2), Values(0))
            Else
                Call sprintf(outbuf, "Starcraft Japanese%s: (%s wins).", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2))
            End If
        Case "NB2W"
            Values() = Split(Mid$(statstring, 6), Space(1))
            If UBound(Values) <> 8 Then
                Call sprintf(outbuf, "a Warcraft II %sbot.", IIf((Values(3) = 1), " (spawn) ", vbNullString))
                Exit Sub
            End If
            If Values(0) > 0 Then
                Call sprintf(outbuf, "Warcraft II%s: (%s wins, with a rating of %s on the ladder).", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2), Values(0))
            Else
                Call sprintf(outbuf, "Warcraft II%s: (%s wins).", IIf((Values(3) = 1), " (spawn) ", vbNullString), Values(2))
            End If
        Case "RHSD"
            Values() = Split(Mid$(statstring, 6), Space(1))
            If UBound(Values) <> 8 Then
                Call strcpy(outbuf, "a Diablo shareware bot.")
                Exit Sub
            End If
            Select Case Values(2)
                Case 0: cType = "warrior"
                Case 1: cType = "rogue"
                Case 2: cType = "sorceror"
            End Select
            Call sprintf(outbuf, "Diablo shareware: (Level %s %s with %s dots, %s strength, %s magic, %s dexterity, %s vitality, and %s gold).", Values(0), cType, Values(1), Values(3), Values(4), Values(5), Values(6), Values(7))
        Case "LTRD"
            Values() = Split(Mid$(statstring, 6), Space(1))
            If UBound(Values) <> 8 Then
                Call strcpy(outbuf, "a Diablo bot.")
                Exit Sub
            End If
            Select Case Values(2)
                Case 0: cType = "warrior"
                Case 1: cType = "rogue"
                Case 2: cType = "sorceror"
            End Select
            Call sprintf(outbuf, "Diablo: (Level %s %s with %s dots, %s strength, %s magic, %s dexterity, %s vitality, and %s gold).", Values(0), cType, Values(1), Values(3), Values(4), Values(5), Values(6), Values(7))
        Case "PX2D"
            Call strcpy(outbuf, ParseD2Stats(statstring))
        Case "VD2D"
            Call strcpy(outbuf, ParseD2Stats(statstring))
        Case "TAHC"
            Call strcpy(outbuf, "a Chat bot.")
    End Select
End Sub

Public Function ParseD2Stats(ByVal stats As String)
    Dim d2classes(0 To 7) As String
        d2classes(0) = "amazon"
        d2classes(1) = "sorceress"
        d2classes(2) = "necromancer"
        d2classes(3) = "paladin"
        d2classes(4) = "barbarian"
        d2classes(5) = "druid"
        d2classes(6) = "assassin"
        d2classes(7) = "unknown class"
    Dim statbuf As String, p() As String, Server As String, Name As String
    If Len(stats) > 4 Then
        Dim sLen As Byte
        sLen = GetServer(stats, Server)
        sLen = GetCharacterName(stats, sLen, Name)
        Call MakeArray(Mid$(stats, sLen), p())
    End If
    If Left$(stats, 4) = "VD2D" Then
        Call strcpy(statbuf, "Diablo II: (")
    Else
        Call strcpy(statbuf, "Diablo II Lord of Destruction: (")
    End If
    If (Len(stats) = 4) Then
        Call strcpy(statbuf, "Open Character).")
    Else
        Dim version As Byte
            version = Asc(p(0)) - &H80
        Dim charclass As Byte
            charclass = Asc(p(13)) - 1
        If (charclass < 0) Or (charclass > 6) Then
            charclass = 7
        End If
        Dim female As Boolean
            female = False
        If (charclass = 0) Or (charclass = 1) Or (charclass = 6) Then
            female = True
        End If
        Dim charlevel As Byte
        charlevel = Asc(p(25))
        Dim hardcore As Byte
        hardcore = Asc(p(26)) And 4
        Dim expansion As Boolean
        expansion = False
        If Left$(stats, 4) = "PX2D" Then
            If (Asc(p(26)) And &H20) Then
                Select Case RShift((Asc(p(27)) And &H18), 3)
                    Case 1
                        If hardcore Then
                            Call strcpy(statbuf, "Destroyer ")
                        Else
                            Call strcpy(statbuf, "Slayer ")
                        End If
                    Case 2
                        If hardcore Then
                            Call strcpy(statbuf, "Conquerer ")
                        Else
                            Call strcpy(statbuf, "Champion ")
                        End If
                    Case 3
                        If hardcore Then
                            Call strcpy(statbuf, "Guardian ")
                        Else
                            If Not female Then
                                Call strcpy(statbuf, "Patriarch ")
                            Else
                                Call strcpy(statbuf, "Matriarch ")
                            End If
                        End If
                End Select
                expansion = True
            End If
        End If
        If Not expansion Then
            Select Case RShift((Asc(p(27)) And &H18), 3)
                Case 1
                    If female = False Then
                        If hardcore Then
                            Call strcpy(statbuf, "Count ")
                        Else
                            Call strcpy(statbuf, "Sir ")
                        End If
                    Else
                        If hardcore Then
                            Call strcpy(statbuf, "Countess ")
                        Else
                            Call strcpy(statbuf, "Dame ")
                        End If
                    End If
                Case 2
                    If female = False Then
                        If hardcore Then
                            Call strcpy(statbuf, "Duke ")
                        Else
                            Call strcpy(statbuf, "Lord ")
                        End If
                    Else
                        If hardcore Then
                            Call strcpy(statbuf, "Duchess ")
                        Else
                            Call strcpy(statbuf, "Lady ")
                        End If
                    End If
                Case 3
                    If female = False Then
                        If hardcore Then
                            Call strcpy(statbuf, "King ")
                        Else
                            Call strcpy(statbuf, "Baron ")
                        End If
                    Else
                        If hardcore Then
                            Call strcpy(statbuf, "Queen ")
                        Else
                            Call strcpy(statbuf, "Baroness ")
                        End If
                    End If
            End Select
        End If
        Call sprintf(statbuf, "%s a ", Name)
        If hardcore Then
            If (Asc(p(26)) And &H8) Then
                Call strcpy(statbuf, "dead ")
            End If
            Call sprintf(statbuf, "hardcore level %s ", charlevel)
        Else
            Call sprintf(statbuf, "level %s ", charlevel)
        End If
        Call sprintf(statbuf, "%s on realm %s).", d2classes(charclass), Server)
    End If
    ParseD2Stats = statbuf
End Function

Private Function GetServer(ByVal statstring As String, ByRef Server As String) As Byte
    'returns the begining of the character name
    Server = Mid$(statstring, 5, InStr(5, statstring, ",") - 5)
    GetServer = InStr(5, statstring, ",") + 1
End Function

Private Function GetCharacterName(ByVal statstring As String, ByVal start As Byte, ByRef cName As String) As Byte
    cName = Mid$(statstring, start, InStr(start, statstring, ",") - start)
    GetCharacterName = InStr(start, statstring, ",") + 1
End Function

Private Sub MakeArray(ByVal text As String, ByRef nArray() As String)
    Dim i As Long
    ReDim nArray(0)
    For i = 0 To Len(text)
        nArray(i) = Mid$(text, i + 1, 1)
        If i <> Len(text) Then
            ReDim Preserve nArray(0 To UBound(nArray) + 1)
        End If
    Next i
End Sub
Public Function RShift(ByVal pnValue As Long, ByVal pnShift As Long) As Double
    ' Equivilant to C's Bitwise >> operator
    RShift = CDbl(pnValue \ (2 ^ pnShift))
End Function
Public Sub strcpy(ByRef Source As String, ByVal nText As String)
    Source = Source & nText
End Sub



