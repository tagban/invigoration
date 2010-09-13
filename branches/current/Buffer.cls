VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Buffer     As String
Private Position   As Long
Private InNOut     As Boolean
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                                                     ByRef Source As Any, _
                                                                     ByVal numbytes As Long)
                                                                     
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Sub Class_Initialize()

    On Error Resume Next
    Position = 1
    On Error GoTo 0

End Sub

Public Sub Clear()

    On Error Resume Next
    Buffer = vbNullString
    Position = 1
    On Error GoTo 0

End Sub

Public Function GetBoolean() As Boolean

Dim lngTemp As Long

    On Error Resume Next
    If LenBuffer() - (Position + 1) >= 4 Then
        CopyMemory lngTemp, ByVal Mid$(Buffer, Position, 4), 4
        Position = Position + 4
    Else
        Exit Function
    End If
    If lngTemp <> 0 Then
        GetBoolean = True
    End If
    On Error GoTo 0
End Function

Public Function GetBuffer() As String

    On Error Resume Next
    GetBuffer = Buffer
    On Error GoTo 0

End Function

Public Function GetByte() As Byte

Dim bytTemp As Byte

    On Error Resume Next
        CopyMemory bytTemp, ByVal Mid$(Buffer, Position, 1), 1
        Position = Position + 1
        GetByte = bytTemp
    On Error GoTo 0
End Function

Public Function GetDWORD() As Long

Dim lngTemp As Long

    On Error Resume Next
    CopyMemory lngTemp, ByVal Mid$(Buffer, Position, 4), 4
    Position = Position + 4
    GetDWORD = lngTemp
    On Error GoTo 0

End Function

Public Function GetDWORDArray(ByVal Amount As Long) As Long()

Dim RtnAry() As Long
Dim i        As Long

    On Error Resume Next
    If Not Amount < 1 Then
        If LenBuffer() - (Position + 1) >= Amount * 4 Then
            ReDim RtnAry(Amount) As Long
            For i = 0 To Amount - 1
                RtnAry(i) = GetDWORD()
            Next i
            GetDWORDArray = RtnAry
        End If
        On Error GoTo 0
    End If

End Function

Public Function GetFixedString(ByVal Length As Long) As String

    On Error Resume Next
    If LenBuffer() - (Position + 1) >= Length Then
        GetFixedString = Mid$(Buffer, Position, Length)
        Position = Position + Length
    Else
        Position = Position + Length
    End If
    On Error GoTo 0

End Function

Public Function GetMode() As Boolean

    On Error Resume Next
    GetMode = InNOut
    On Error GoTo 0

End Function

Public Function GetNonNTString(ByVal Length As Long) As String

    On Error Resume Next
    If LenBuffer() - (Position + 1) >= Length Then
        GetNonNTString = Mid$(Buffer, Position, Length)
        Position = Position + Length
    End If
    On Error GoTo 0

End Function

Public Function GetPosition() As Long

    On Error Resume Next
    GetPosition = Position
    On Error GoTo 0

End Function

Public Function GetRaw(ByVal Length As Long) As String

    On Error Resume Next
    GetRaw = Mid$(Buffer, Position, Length)
    Position = Position + Length
    On Error GoTo 0

End Function

Public Function GetSTRING() As String

Dim Pos As Long
    On Error Resume Next
    GetSTRING = vbNullString
    Pos = InStr(Position, Buffer, Chr$(&H0))
    If Pos = 0 Then
        Exit Function
    End If
    GetSTRING = Mid$(Buffer, Position, Pos - Position)
    Position = Pos + 1
    On Error GoTo 0

End Function

Public Function GetStringArray(ByVal Amount As Long) As String()

Dim RtnAry() As String
Dim i        As Long

    On Error Resume Next
    If Not Amount < 1 Then
    ReDim RtnAry(Amount) As String
        For i = 0 To Amount - 1
            RtnAry(i) = GetSTRING()
        Next i
        GetStringArray = RtnAry
        On Error GoTo 0
    End If

End Function

Public Function GetTermString(ByVal Terminator As Integer) As String
Dim Pos As Long
    On Error Resume Next
    GetTermString = vbNullString
    Pos = InStr(Position, Buffer, Chr$(Terminator))
    If Pos = 0 Then
        Exit Function
    End If
    GetTermString = Mid$(Buffer, Position, Pos - Position)
    Position = Pos + 1
    On Error GoTo 0

End Function

Public Function GetWORD() As Long

Dim lngTemp As Long

    On Error Resume Next
    If LenBuffer() - (Position + 1) >= 2 Then
        CopyMemory lngTemp, ByVal Mid$(Buffer, Position, 2), 2
        Position = Position + 2
        GetWORD = lngTemp
    End If
    On Error GoTo 0

End Function

Public Function GetWORDArray(ByVal Amount As Long) As Long()

Dim RtnAry() As Long
Dim i        As Long
    If Not Amount < 1 Then
        If LenBuffer() - (Position + 1) >= Amount * 2 Then
            ReDim RtnAry(Amount) As Long
            For i = 0 To Amount - 1
                RtnAry(i) = GetWORD()
            Next i
            GetWORDArray = RtnAry
        Else
        End If
    End If

End Function

Public Function GetFILETIME() As String
    Dim FT As FILETIME
     Dim TEMP As String * 8
     TEMP = Mid$(Buffer, Position, 8)
     CopyMemory FT, ByVal TEMP, 8
     Position = Position + 8
     GetFILETIME = FT.dwHighDateTime & Space$(1) & FT.dwLowDateTime
 End Function

Public Sub InsertByte(ByVal Data As Byte)

    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    Buffer = Buffer & MakeByte(Data)
    On Error GoTo 0

End Sub

Public Sub InsertDWORD(ByVal Data As Long)

    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    Buffer = Buffer & MakeDWORD(Data)
    On Error GoTo 0

End Sub

Public Sub InsertDWORDArray(Data() As Long)

Dim i As Long

    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    If UBound(Data) = 0 Then
        If Data(0) = 0 Then
            Exit Sub
        End If

    End If
    For i = LBound(Data) To UBound(Data)
        InsertDWORD Data(i)
    Next i
    On Error GoTo 0

End Sub

Public Sub InsertNonNTString(ByVal Data As String)

    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    Buffer = Buffer & Data
    On Error GoTo 0

End Sub

Public Sub InsertString(ByVal Data As String)

    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    Buffer = Buffer & Data & Chr$(&H0)
    On Error GoTo 0

End Sub

Public Sub InsertStringArray(Data() As String)

Dim i As Long

    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    If UBound(Data) = 0 Then
        If Len(Data(0)) = 0 Then
            Exit Sub
        End If
    End If
    For i = LBound(Data) To UBound(Data)
        InsertString Data(i)
    Next i
    On Error GoTo 0

End Sub

Public Sub InsertStringList(Data() As String)

Dim i As Long

    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    If UBound(Data) = 0 Then
        If Len(Data(0)) = 0 Then
            Exit Sub
        End If
    End If
    For i = LBound(Data) To UBound(Data)
        InsertString Data(i)
    Next i
    InsertByte &H0
    On Error GoTo 0

End Sub

Public Sub InsertWORD(ByVal Data As Integer)

    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    Buffer = Buffer & MakeWORD(Data)
    On Error GoTo 0

End Sub

Public Sub InsertWORDArray(Data() As Integer)

Dim i As Long

    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    If UBound(Data) = 0 Then
        If Data(0) = 0 Then
            Exit Sub
        End If
    End If
    For i = LBound(Data) To UBound(Data)
        InsertWORD Data(i)
    Next i
    On Error GoTo 0

End Sub

Public Function LenBuffer() As Long

    On Error Resume Next
    LenBuffer = Len(Buffer)
    On Error GoTo 0

End Function

Public Function MakeByte(ByVal Data As Byte) As String

Dim Result As String * 1

    On Error Resume Next
    CopyMemory ByVal Result, Data, 1
    MakeByte = Result
    On Error GoTo 0

End Function

Public Function MakeDWORD(ByVal Data As Long) As String

Dim Result As String * 4

    On Error Resume Next
    CopyMemory ByVal Result, Data, 4
    MakeDWORD = Result
    On Error GoTo 0

End Function

Public Function MakeWORD(ByVal Data As Integer) As String

Dim Result As String * 2

    On Error Resume Next
    CopyMemory ByVal Result, Data, 2
    MakeWORD = Result
    On Error GoTo 0

End Function

Public Sub SendMCPPacket(Socket As Winsock, _
                         PacketID As Byte)

    On Error Resume Next
    If Not InNOut Then
        Socket.SendData MakeWORD(Len(Buffer) + 3) & Chr$(PacketID) & Buffer
        Clear
        On Error GoTo 0
    End If
End Sub


Public Sub SendRaw(Socket As Winsock)

    On Error Resume Next
    If Not InNOut Then
        Socket.SendData Buffer
        Clear
        On Error GoTo 0
    End If

End Sub


Public Sub SetBuffer(ByVal Data As String)
    On Error Resume Next
    Buffer = Data
    InNOut = True
    On Error GoTo 0

End Sub

Public Sub SetInNOut(ByVal Bool As Boolean)

    On Error Resume Next
    InNOut = Bool
    On Error GoTo 0

End Sub

Public Sub SetPosition(ByVal Pos As Long)

    On Error Resume Next
    Position = Pos
    On Error GoTo 0

End Sub

Public Sub Skip(ByVal Bytes As Long)

    On Error Resume Next
    Position = Position + Bytes
    On Error GoTo 0

End Sub



