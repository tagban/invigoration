Attribute VB_Name = "inet"
Option Explicit

Const INTERNET_OPEN_TYPE_PRECONFIG = 0

Const INTERNET_FLAG_RELOAD = &H80000000


Private Declare Function InternetOpenUrl Lib _
    "wininet.dll" Alias "InternetOpenUrlA" _
    (ByVal hInternetSession As Long, _
    ByVal lpszUrl As String, _
    ByVal lpszHeaders As String, _
    ByVal dwHeadersLength As Long, _
    ByVal dwFlags As Long, ByVal dwContext As Long) _
    As Long

Private Declare Function InternetOpen Lib "wininet.dll" _
    Alias "InternetOpenA" (ByVal sAgent As String, _
    ByVal lAccessType As Long, _
    ByVal sProxyName As String, _
    ByVal sProxyBypass As String, ByVal lFlags As Long) _
    As Long

Private Declare Function InternetReadFile Lib _
    "wininet.dll" (ByVal hFile As Long, _
    ByVal sBuffer As String, _
    ByVal lNumBytesToRead As Long, _
    lNumberOfBytesRead As Long) As Integer

Private Declare Function InternetCloseHandle Lib _
    "wininet.dll" (ByVal hInet As Long) As Integer

Public Function GoWinInet(sURL$) As String
    Dim sBuffer As String * 4096, sReturn As String
    Dim lNumBytes As Long, lSession As Long, lFile As Long
    Dim bReadOK As Boolean
    lSession = InternetOpen("GSInternet", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    lFile = InternetOpenUrl(lSession, sURL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
    If lFile Then
        Do
            bReadOK = InternetReadFile(lFile, sBuffer, Len(sBuffer), lNumBytes)
            If lNumBytes Then
            sReturn = sReturn & Left$(sBuffer, lNumBytes)
            End If
        Loop While bReadOK And lNumBytes > 0
        InternetCloseHandle (lFile)
        GoWinInet = sReturn
    End If
End Function

