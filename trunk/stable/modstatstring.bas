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


Private Function GetServer(ByVal statstring As String, ByRef Server As String) As Byte
    'returns the begining of the character name
    Server = Mid$(statstring, 5, InStr(5, statstring, ",") - 5)
    GetServer = InStr(5, statstring, ",") + 1
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



