Attribute VB_Name = "Bot_Media"
Option Explicit 'Everythings Declared                                                                                                                                                                                                                                                                                                                                                                                                          _
                                                                                                                                                                                                                                                                                         _


Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" ( _
ByVal lpstrCommand As String, _
ByVal lpstrReturnString As String, _
ByVal uReturnLength As Long, _
ByVal hwndCallback As Long) As Long
'lpstrCommand  = The String Command To Send To Windows ("play file")
'lpstrReturnString = A String With A Buffer (Spaces, Null Characters) Of uReturnlength
'uReturnLength = Number Of Characters To Return In ^
'hwndCallBack = Handle Of The Application Thats Sending The Messages (Can Use 0 If You Want)
Public Declare Function FloodFill Lib "gdi32" ( _
ByVal hdc As Long, _
ByVal x As Long, _
ByVal y As Long, _
ByVal crColor As Long) As Long
'hdc = The Handle To A Device Context
'x = The X Position Of Where To Start Our FloodFill
'y = The Y Position Of Where To Start Our FloodFill
'crColor = The Color To Fill With
Type ID3Tag 'Our ID3 Data Type
Name As String * 30 'Holds The Songname (Max Is 30 Char)
Artist As String * 30 'Holds The Artist Name (Max Is 30 Char)
Album As String * 30 'Holds The Album Name (Max Is 30 Char)
Date As String * 4 'Holds The Date (Max Is 4 Char) Usually 2001,2002,2003...etc
Comments As String * 30 'Any Comments On The Songs (Max Is 30 Char)
Style As Byte 'Holds The Style-This Is A Number Between 0-255 Anything Else Will Result In An Error
End Type
Global Info As ID3Tag, FileLoaded As String 'Holds The Id3's Info, Holds The File That Is Currently Loaded
Sub GetTag(FileName As String, Info As ID3Tag)
    'Gets An Mp3 ID3 Tag
    Dim TagPrint As String * 3 'To Hold "TAG" If It Exists
    Dim File As Integer 'To Hold Our FreeFile Value
    Dim i As ID3Tag
    Info.Style = 255 'Set The Stile To 255 (Unknown) Just To Prevent Errors
    Info = i
    File = FreeFile
    On Error Resume Next
    Open FileName For Binary As File
    Get File, FileLen(FileName) - 127, TagPrint 'Check To See If A Tag Exists
    If TagPrint = "TAG" Then 'If There Is A Tag...
    Get File, FileLen(FileName) - 124, Info '..Put It Into The "Info" ID3 Data Type
End If
Close File
End Sub
Function SaveTag(FileName As String, Info As ID3Tag) As Boolean
    'Saves A Mp3 ID3 Tag
    'Returns True If Save Was Successful
    Dim File As Integer 'To Hold Our FreeFile Value
    Dim TagPrint As String * 3 'To Hold "TAG"
    File = FreeFile
    On Error GoTo ErrHand
    Open FileName For Binary As File
    Get File, LOF(File) - Len(Info) - 2, TagPrint 'First We Check To See If There Is A Tag
    If TagPrint = "TAG" Then 'If There Is A Tag..
    Put File, , Info '...Simply Put The New Info In
Else 'Otherwise We...
    Put File, LOF(File), "TAG" '...Write "TAG" To The File Stating There Is A Tag
    Put File, , Info 'And We Write The ID3 Information In
End If
Close File
SaveTag = True
Exit Function
ErrHand:
SaveTag = False 'If Theres An Error It Returns False
End Function
Function FileExt(File As String) As String
    'This Function Returns The File Extension
    Dim x As Long
    x = InStrRev(File, ".") 'Search For "." Backwards
    FileExt = LCase(Mid(File, x + 1)) 'Return Only The Ext
End Function
Function SecsToMin(Seconds As Long, MilliSeconds As Boolean) As String
    'Makes 130 Seconds To 2:10
    'If Its 130000 Milliseconds Just Set MilliSeconds To True
    'JavaScriptVersion: function SecsToMin(Seconds) { var Times=""; var x =0; var Y=0; var z=""; if (Seconds > 59){x = Math.floor(Seconds / 60);Y = Math.abs(Seconds - (x * 60));if (Y < 10) { z = "0" + String(Y);} else {z = String(Y)};Times = String(x) + ":" + String(z);}else{if(Seconds >= 10) {Times = "0:" + String(Seconds)} else {Times = "0:0" + String(Seconds)}}return Times;}
    'This Is An Excerpt From My Module
    Dim Times, x As Long, y As Long, z As String
    If MilliSeconds Then 'If Seconds Is Actually Milliseconds...
    Seconds = Seconds / 1000 '..We Divide It By 1000 And Convert It To Seconds
End If
If Seconds > 59 Then 'If The Total Seconds Is Greater Then 59...
x = Int(Seconds / 60) '...We Get The Integer Value Of Seconds Divided By 60 (Even Though Its Declared As Long Which By Default Can Only Be An Integer No Decimals Allowed)
y = Abs(Seconds - (x * 60)) 'We Now Get The Absolute Value Of Seconds Minus (X Mulitiplied By 60 )
If y < 10 Then z = "0" & y Else z = y 'If Y (Look Up One Line) Is Less Than Ten We Put The 0 (Like In 09, 08,...)
Times = x & ":" & z
ElseIf Seconds >= 10 Then
Times = "0:" & Seconds
Else
Times = "0:0" & Seconds
End If
SecsToMin = Times
End Function
Function StyleId(sStyle As String) As Long
    'For Mp3's
    'If Anyone Knows A Better Way To List This (No Listboxes Only Code)
    'Please Email Me
    Select Case sStyle
        Case "Blues"
            StyleId = 0
        Case "Classic Rock"
            StyleId = 1
        Case "Country"
            StyleId = 2
        Case "Dance"
            StyleId = 3
        Case "Disco"
            StyleId = 4
        Case "Funk"
            StyleId = 5
        Case "Grunge"
            StyleId = 6
        Case "Hip-Hop"
            StyleId = 7
        Case "Jazz"
            StyleId = 8
        Case "Metal"
            StyleId = 9
        Case "New Age"
            StyleId = 10
        Case "Oldies"
            StyleId = 11
        Case "Other"
            StyleId = 12
        Case "Pop"
            StyleId = 13
        Case "R&B"
            StyleId = 14
        Case "Rap"
            StyleId = 15
        Case "Reggae"
            StyleId = 16
        Case "Rock"
            StyleId = 17
        Case "Techno"
            StyleId = 18
        Case "Industrial"
            StyleId = 19
        Case "Alternative"
            StyleId = 20
        Case "Ska"
            StyleId = 21
        Case "Death Metal"
            StyleId = 22
        Case "Pranks"
            StyleId = 23
        Case "Soundtrack"
            StyleId = 24
        Case "Euro-Techno"
            StyleId = 25
        Case "Ambient"
            StyleId = 26
        Case "Trip-Hop"
            StyleId = 27
        Case "Vocal"
            StyleId = 28
        Case "Jazz+Funk"
            StyleId = 29
        Case "Fusion"
            StyleId = 30
        Case "Trance"
            StyleId = 31
        Case "Classical"
            StyleId = 32
        Case "Instrumental"
            StyleId = 33
        Case "Acid"
            StyleId = 34
        Case "House"
            StyleId = 35
        Case "Game"
            StyleId = 36
        Case "Sound Clip"
            StyleId = 37
        Case "Gospel"
            StyleId = 38
        Case "Noise"
            StyleId = 39
        Case "AlternRock"
            StyleId = 40
        Case "Bass"
            StyleId = 41
        Case "Soul"
            StyleId = 42
        Case "Punk"
            StyleId = 43
        Case "Space"
            StyleId = 44
        Case "Meditative"
            StyleId = 45
        Case "Instrumental Pop"
            StyleId = 46
        Case "Instrumental Rock"
            StyleId = 47
        Case "Ethnic"
            StyleId = 48
        Case "Gothic"
            StyleId = 49
        Case "Darkwave"
            StyleId = 50
        Case "Techno-Industrial"
            StyleId = 51
        Case "Electronic"
            StyleId = 52
        Case "Pop-Folk"
            StyleId = 53
        Case "Eurodance"
            StyleId = 54
        Case "Dream"
            StyleId = 55
        Case "Southern Rock"
            StyleId = 56
        Case "Comedy"
            StyleId = 57
        Case "Cult"
            StyleId = 58
        Case "Gangsta"
            StyleId = 59
        Case "Top 40"
            StyleId = 60
        Case "Christian Rap"
            StyleId = 61
        Case "Pop/Funk"
            StyleId = 62
        Case "Jungle"
            StyleId = 63
        Case "Native American"
            StyleId = 64
        Case "Cabaret"
            StyleId = 65
        Case "New Wave"
            StyleId = 66
        Case "Psychadelic"
            StyleId = 67
        Case "Rave"
            StyleId = 68
        Case "Showtunes"
            StyleId = 69
        Case "Trailer"
            StyleId = 70
        Case "Lo-Fi"
            StyleId = 71
        Case "Tribal"
            StyleId = 72
        Case "Acid Punk"
            StyleId = 73
        Case "Acid Jazz"
            StyleId = 74
        Case "Polka"
            StyleId = 75
        Case "Retro"
            StyleId = 76
        Case "Musical"
            StyleId = 77
        Case "Rock & Roll"
            StyleId = 78
        Case "Hard Rock"
            StyleId = 79
        Case "Folk"
            StyleId = 80
        Case "Folk-Rock"
            StyleId = 81
        Case "National Folk"
            StyleId = 82
        Case "Swing"
            StyleId = 83
        Case "Fast Fusion"
            StyleId = 84
        Case "Bebob"
            StyleId = 85
        Case "Latin"
            StyleId = 86
        Case "Revival"
            StyleId = 87
        Case "Celtic"
            StyleId = 88
        Case "Bluegrass"
            StyleId = 89
        Case "Avantgarde"
            StyleId = 90
        Case "Gothic Rock"
            StyleId = 91
        Case "Progressive Rock"
            StyleId = 92
        Case "Psychedelic Rock"
            StyleId = 93
        Case "Symphonic Rock"
            StyleId = 94
        Case "Slow Rock"
            StyleId = 95
        Case "Big Band"
            StyleId = 96
        Case "Chorus"
            StyleId = 97
        Case "Easy Listening"
            StyleId = 98
        Case "Acoustic"
            StyleId = 99
        Case "Humour"
            StyleId = 100
        Case "Speech"
            StyleId = 101
        Case "Chanson"
            StyleId = 102
        Case "Opera"
            StyleId = 103
        Case "Chamber Music"
            StyleId = 104
        Case "Sonata"
            StyleId = 105
        Case "Symphony"
            StyleId = 106
        Case "Booty Bass"
            StyleId = 107
        Case "Primus"
            StyleId = 108
        Case "Porn Groove"
            StyleId = 109
        Case "Satire"
            StyleId = 110
        Case "Slow Jam"
            StyleId = 111
        Case "Club"
            StyleId = 112
        Case "Tango"
            StyleId = 113
        Case "Samba"
            StyleId = 114
        Case "Folklore"
            StyleId = 115
        Case "Ballad"
            StyleId = 116
        Case "Power Ballad"
            StyleId = 117
        Case "Rhythmic Soul"
            StyleId = 118
        Case "Freestyle"
            StyleId = 119
        Case "Duet"
            StyleId = 120
        Case "Punk Rock"
            StyleId = 121
        Case "Drum Solo"
            StyleId = 122
        Case "Acapella"
            StyleId = 123
        Case "Euro-House"
            StyleId = 124
        Case "Dance Hall"
            StyleId = 125
        Case "Unknown"
            StyleId = 255
    End Select
End Function
Function Style(Id As Long) As String
    'For MP3's
    Select Case Id
        Case 0
            Style = "Blues"
        Case 1
            Style = "Classic Rock"
        Case 2
            Style = "Country"
        Case 3
            Style = "Dance"
        Case 4
            Style = "Disco"
        Case 5
            Style = "Funk"
        Case 6
            Style = "Grunge"
        Case 7
            Style = "Hip-Hop"
        Case 8
            Style = "Jazz"
        Case 9
            Style = "Metal"
        Case 10
            Style = "New Age"
        Case 11
            Style = "Oldies"
        Case 12
            Style = "Other"
        Case 13
            Style = "Pop"
        Case 14
            Style = "R&B"
        Case 15
            Style = "Rap"
        Case 16
            Style = "Reggae"
        Case 17
            Style = "Rock"
        Case 18
            Style = "Techno"
        Case 19
            Style = "Industrial"
        Case 20
            Style = "Alternative"
        Case 21
            Style = "Ska"
        Case 22
            Style = "Death Metal"
        Case 23
            Style = "Pranks"
        Case 24
            Style = "Soundtrack"
        Case 25
            Style = "Euro-Techno"
        Case 26
            Style = "Ambient"
        Case 27
            Style = "Trip-Hop"
        Case 28
            Style = "Vocal"
        Case 29
            Style = "Jazz+Funk"
        Case 30
            Style = "Fusion"
        Case 31
            Style = "Trance"
        Case 32
            Style = "Classical"
        Case 33
            Style = "Instrumental"
        Case 34
            Style = "Acid"
        Case 35
            Style = "House"
        Case 36
            Style = "Game"
        Case 37
            Style = "Sound Clip"
        Case 38
            Style = "Gospel"
        Case 39
            Style = "Noise"
        Case 40
            Style = "AlternRock"
        Case 41
            Style = "Bass"
        Case 42
            Style = "Soul"
        Case 43
            Style = "Punk"
        Case 44
            Style = "Space"
        Case 45
            Style = "Meditative"
        Case 46
            Style = "Instrumental Pop"
        Case 47
            Style = "Instrumental Rock"
        Case 48
            Style = "Ethnic"
        Case 49
            Style = "Gothic"
        Case 50
            Style = "Darkwave"
        Case 51
            Style = "Techno-Industrial"
        Case 52
            Style = "Electronic"
        Case 53
            Style = "Pop-Folk"
        Case 54
            Style = "Eurodance"
        Case 55
            Style = "Dream"
        Case 56
            Style = "Southern Rock"
        Case 57
            Style = "Comedy"
        Case 58
            Style = "Cult"
        Case 59
            Style = "Gangsta"
        Case 60
            Style = "Top 40"
        Case 61
            Style = "Christian Rap"
        Case 62
            Style = "Pop/Funk"
        Case 63
            Style = "Jungle"
        Case 64
            Style = "Native American"
        Case 65
            Style = "Cabaret"
        Case 66
            Style = "New Wave"
        Case 67
            Style = "Psychadelic"
        Case 68
            Style = "Rave"
        Case 69
            Style = "Showtunes"
        Case 70
            Style = "Trailer"
        Case 71
            Style = "Lo-Fi"
        Case 72
            Style = "Tribal"
        Case 73
            Style = "Acid Punk"
        Case 74
            Style = "Acid Jazz"
        Case 75
            Style = "Polka"
        Case 76
            Style = "Retro"
        Case 77
            Style = "Musical"
        Case 78
            Style = "Rock & Roll"
        Case 79
            Style = "Hard Rock"
        Case 80
            Style = "Folk"
        Case 81
            Style = "Folk-Rock"
        Case 82
            Style = "National Folk"
        Case 83
            Style = "Swing"
        Case 84
            Style = "Fast Fusion"
        Case 85
            Style = "Bebob"
        Case 86
            Style = "Latin"
        Case 87
            Style = "Revival"
        Case 88
            Style = "Celtic"
        Case 89
            Style = "Bluegrass"
        Case 90
            Style = "Avantgarde"
        Case 91
            Style = "Gothic Rock"
        Case 92
            Style = "Progressive Rock"
        Case 93
            Style = "Psychedelic Rock"
        Case 94
            Style = "Symphonic Rock"
        Case 95
            Style = "Slow Rock"
        Case 96
            Style = "Big Band"
        Case 97
            Style = "Chorus"
        Case 98
            Style = "Easy Listening"
        Case 99
            Style = "Acoustic"
        Case 100
            Style = "Humour"
        Case 101
            Style = "Speech"
        Case 102
            Style = "Chanson"
        Case 103
            Style = "Opera"
        Case 104
            Style = "Chamber Music"
        Case 105
            Style = "Sonata"
        Case 106
            Style = "Symphony"
        Case 107
            Style = "Booty Bass"
        Case 108
            Style = "Primus"
        Case 109
            Style = "Porn Groove"
        Case 110
            Style = "Satire"
        Case 111
            Style = "Slow Jam"
        Case 112
            Style = "Club"
        Case 113
            Style = "Tango"
        Case 114
            Style = "Samba"
        Case 115
            Style = "Folklore"
        Case 116
            Style = "Ballad"
        Case 117
            Style = "Power Ballad"
        Case 118
            Style = "Rhythmic Soul"
        Case 119
            Style = "Freestyle"
        Case 120
            Style = "Duet"
        Case 121
            Style = "Punk Rock"
        Case 122
            Style = "Drum Solo"
        Case 123
            Style = "Acapella"
        Case 124
            Style = "Euro-House"
        Case 125
            Style = "Dance Hall"
        Case Else
            Style = "Unknown"
    End Select
End Function


