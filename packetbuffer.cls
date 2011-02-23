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
Private buffer As String

Public Function InsertDWORDArray(data() As Long)
    Dim i As Integer
    For i = LBound(data) To UBound(data) Step 1
        buffer = buffer & MakeDWORD(data(i))
    Next i
End Function

Public Function InsertDWORD(data As Long)
    buffer = buffer & MakeDWORD(data)
    If debugmode = 1 Then
        AddChat HEXPINK, "InsertDWORD: " & data
    End If
End Function

Public Function InsertData(data As String)
    buffer = buffer & data
    If debugmode = 1 Then
        AddChat HEXPINK, "Data: " & data
    End If
End Function

Public Function InsertWORD(data As Integer)
    buffer = buffer & MakeWORD(data)
End Function

Public Function InsertBYTE(data As Integer)
    buffer = buffer & Chr(data)
    If debugmode = 1 Then
        AddChat HEXPINK, "BYTE: " & data
    End If

End Function

Public Sub InsertBytes(data As String)
    Dim i As Long
    Dim enqueueer As String
    
    For i = 1 To Len(data) Step 3
        enqueueer = enqueueer & Chr(Val("&h0" & Mid(data, i, 2)))
    If debugmode = 1 Then
        AddChat HEXPINK, "Bytes: " & data
    End If
    Next i
    buffer = buffer & enqueueer
End Sub

Public Function InsertNTString(data As String)
    buffer = buffer & data & Chr(0)
    If debugmode = 1 Then
        AddChat HEXPINK, "NTString: " & data
    End If
End Function
Public Function InsertATString(data As String)
    buffer = buffer & data & Chr(&HA)
    If debugmode = 1 Then
        AddChat HEXPINK, "ATString: " & data
    End If
End Function
Public Function InsertNonNTString(data As String)
    buffer = buffer & data
    If debugmode = 1 Then
        AddChat HEXPINK, "NonNTString: " & data
    End If
End Function
Public Function InsertNonNTStringArray(data() As String)
Dim i As Integer
For i = LBound(data) To UBound(data) Step 1
    buffer = buffer & data(i)
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
    buffer = vbNullString
End Function

Public Function SendPacket(PacketId As Byte)
    On Error GoTo Err:
If frmMain.wsBnet.State <> sckConnected Then: Exit Function
'' Runtime Error 6 ??
    frmMain.wsBnet.SendData Chr(&HFF) & Chr(PacketId) & MakeWORD(Len(buffer) + 4) & buffer
    If debugmode = 1 Then
        AddChat HEXPINK, "Sent Packet: 0x" & Hex(PacketId)
    End If
    Clear
Err:
    'AddChat D2Red, "ERROR: ", HEXPINK, "Packet send error: 0x" & Hex(PacketId)
    ''' Error Handler!
    'info = "<font color=red>ERROR: </font> <font color=pink>Packet send error: 0x" & Hex(PacketId) & "</font><br>"
    'frmMain.txtTesting.text = info
End Function
Public Function SendRPacket(PacketId As Byte)
On Error GoTo Err3:
If frmMain.wsRealm.State <> sckConnected Then: Exit Function
    frmMain.wsRealm.SendData MakeWORD(Len(buffer) + 3) & Chr(PacketId) & buffer
    If debugmode = 1 Then
        AddChat HEXPINK, "Sent Realm Packet: 0x" & Hex(PacketId)
    End If
    Clear
Err3:
End Function
Public Function SendBNLSPacket(PacketId As Byte)
On Error GoTo Err2:
If frmMain.wsBnls.State <> sckConnected Then: Exit Function
    frmMain.wsBnls.SendData MakeWORD(Len(buffer) + 3) & Chr(PacketId) & buffer
    If debugmode = 1 Then
        AddChat HEXPINK, "Sent BNLS Packet: 0x" & Hex(PacketId)
    End If
    Clear
Err2:
End Function

