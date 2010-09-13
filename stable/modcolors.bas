Attribute VB_Name = "modColors"
Public Function fBotColors()
Dim sOne As String, iOne As Long
    With frmMain.rtbChat
Do Until InStr(.text, " ") = 0
    iOne = InStr(.text, " "): .SelStart = iOne: .SelLength = 1: sOne = .SelText: .SelStart = iOne - 1: .SelLength = Len(.text)
 Select Case sOne
            Case "r": .SelColor = vbRed
            Case "w": .SelColor = D2White
            Case "q": .SelColor = vbGrey
            Case "g": .SelColor = D2Green
            Case "y": .SelColor = vbYellow
            Case "b": .SelColor = D2MedBlue
            Case "o": .SelColor = D2Orange
            Case "c": .SelColor = D2LtBlue
            Case "p": .SelColor = D2Purple
            Case "l": .SelColor = D2LtYellow
            Case "e": .SelColor = D2Beige2
            Case "k": .SelColor = HEXPINK
End Select
    .SelStart = iOne - 1: .SelLength = 2: .SelText = ""
Loop
End With
End Function
