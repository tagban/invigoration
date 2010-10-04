VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PacketBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Buffer As String

Public Function InsertDWORDArray(Data() As Long)
    Dim i As Integer
    For i = LBound(Data) To UBound(Data) Step 1
        Buffer = Buffer & MakeDWORD(Data(i))
    Next i
End Function

Public Function InsertDWORD(Data As Long)
    Buffer = Buffer & MakeDWORD(Data)
    If debugmode = 1 Then
        AddChat HEXPINK, "InsertDWORD: " & Data
    End If
End Function

Public Function InsertData(Data As String)
    Buffer = Buffer & Data
    If debugmode = 1 Then
        AddChat HEXPINK, "Data: " & Data
    End If
End Function

Public Function InsertWORD(Data As Integer)
    Buffer = Buffer & MakeWORD(Data)
End Function

Public Function InsertByte(Data As Integer)
    Buffer = Buffer & Chr(Data)
    If debugmode = 1 Then
        AddChat HEXPINK, "BYTE: " & Data
    End If

End Function

Public Sub InsertBytes(Data As String)
    Dim i As Long
    Dim enqueueer As String
    
    For i = 1 To Len(Data) Step 3
        enqueueer = enqueueer & Chr(Val("&h0" & Mid(Data, i, 2)))
    If debugmode = 1 Then
        AddChat HEXPINK, "Bytes: " & Data
    End If
    Next i
    Buffer = Buffer & enqueueer
End Sub

Public Function InsertNTString(Data As String)
    Buffer = Buffer & Data & Chr(0)
    If debugmode = 1 Then
        AddChat HEXPINK, "NTString: " & Data
    End If
End Function
Public Function InsertATString(Data As String)
    Buffer = Buffer & Data & Chr(&HA)
    If debugmode = 1 Then
        AddChat HEXPINK, "ATString: " & Data
    End If
End Function
Public Function InsertNonNTString(Data As String)
    Buffer = Buffer & Data
    If debugmode = 1 Then
        AddChat HEXPINK, "NonNTString: " & Data
    End If
End Function
Public Function InsertNonNTStringArray(Data() As String)
Dim i As Integer
For i = LBound(Data) To UBound(Data) Step 1
    Buffer = Buffer & Data(i)
    If debugmode = 1 Then
        'AddChat HEXPINK, "NonNTStringARRAY: " & data
    End If
    Next
End Function
Public Function MakeDWORD(value As Long) As String
    Dim Result As String * 4
    CopyMemory ByVal Result, value, 4
    MakeDWORD = Result
End Function

Function MakeWORD(value As Integer) As String
    Dim Result As String * 2
    CopyMemory ByVal Result, value, 2
    MakeWORD = Result
End Function

Public Function Clear()
    Buffer = vbNullString
End Function

Public Function SendPacket(PacketID As Byte)
    On Error GoTo Err:
If frmMain.wsBnet.State <> sckConnected Then: Exit Function
'' Runtime Error 6 ??
    frmMain.wsBnet.SendData Chr(&HFF) & Chr(PacketID) & MakeWORD(Len(Buffer) + 4) & Buffer
    If debugmode = 1 Then
        AddChat HEXPINK, "Sent Packet: 0x" & Hex(PacketID)
    End If
    Clear
Err:
    'AddChat D2Red, "ERROR: ", HEXPINK, "Packet send error: 0x" & Hex(PacketId)
    ''' Error Handler!
    'info = "<font color=red>ERROR: </font> <font color=pink>Packet send error: 0x" & Hex(PacketId) & "</font><br>"
    'frmMain.txtTesting.text = info
End Function

Public Function SendBNLSPacket(PacketID As Byte)
On Error GoTo Err2:
If frmMain.wsBnls.State <> sckConnected Then: Exit Function
    frmMain.wsBnls.SendData MakeWORD(Len(Buffer) + 3) & Chr(PacketID) & Buffer
    If debugmode = 1 Then
        AddChat HEXPINK, "Sent BNLS Packet: 0x" & Hex(PacketID)
    End If
    Clear
Err2:
End Function

