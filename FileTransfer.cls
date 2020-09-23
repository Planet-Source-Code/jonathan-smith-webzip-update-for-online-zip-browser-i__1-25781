VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'File transfer class written in about ten minutes for downloading the entire zip

Private WithEvents dl As Winsock
Attribute dl.VB_VarHelpID = -1

Private Enum wzDownloadFlags
    dlIdle = 0
    dlConnecting = 1
    dlCheckingSize = 2
    dlDownloading = 3
End Enum

Private dlflags As wzDownloadFlags
Private m_szFile As New CString
Private m_szServer As New CString
Private m_nFileSize As Long
Private m_nBytes As Long
Private m_szData As New CString
Private m_szLocalFile As New CString


Public Event Status(ByVal szStatus As String)
Public Event Progress(ByVal nPercent As Long, ByVal nReceivedTotal As Long)

Public Property Get FileSize() As Long
    FileSize = m_nFileSize
End Property


Public Sub Download(ByVal szURL As String, ByVal szLocalFile As String)
    Dim szServer As New CString
    szURL = Replace(szURL, "http://", "")
    m_szServer = Trim(Left(szURL, InStr(szURL, "/") - 1))
    m_szFile = Trim(Right(szURL, Len(szURL) - Len(m_szServer)))
    m_szLocalFile = szLocalFile
    
    If dlflags <> dlIdle Then
        RaiseEvent Status("Socket currently in use")
        Exit Sub
    End If
    
    RaiseEvent Status("Connecting to server")
    dlflags = dlConnecting
    
    dl.Connect m_szServer, 80
End Sub

Private Sub Class_Initialize()
    Set dl = New Winsock
End Sub

Private Sub Class_Terminate()
    dl.Close
    
End Sub

Private Sub dl_Connect()
    Dim szRequest As New CString
    
    If dlflags = dlConnecting Then
        RaiseEvent Status("Checking file size")
        dlflags = dlCheckingSize
    ElseIf dlflags = dlCheckingSize Then
        RaiseEvent Status("Downloading archive")
        dlflags = dlDownloading
    End If
    
    szRequest = szRequest & "GET " & m_szFile & " HTTP/1.1" & vbCrLf
    szRequest = szRequest & "Accept: *.*, */*" & vbCrLf
    szRequest = szRequest & "User-Agent: WebZip" & vbCrLf
    szRequest = szRequest & "Referer: " & dl.LocalIP & vbCrLf
    szRequest = szRequest & "Host: " & dl.LocalIP & vbCrLf & vbCrLf
    dl.SendData szRequest

End Sub

Private Sub SaveFile()
    Dim f As Integer
    f = FreeFile
    Open m_szLocalFile For Binary As #f
        Put #f, , m_szData.Value
    Close #f
    
End Sub

Private Sub dl_DataArrival(ByVal bytesTotal As Long)
    Dim szTempData As String
    Dim szData As New CString
    dl.GetData szTempData, vbString
    
    szData = szTempData
    
    Select Case dlflags
    Case dlCheckingSize
        Dim szHeader As New CString
        dl.Close
        szHeader = LCase(Left(szData, InStr(szData, vbCrLf & vbCrLf)))
        If InStr(szHeader, "content-length") > 0 Then
            szHeader = Mid(szHeader, InStr(szHeader, "content-length"))
            szHeader = Left(szHeader, InStr(szHeader, Chr(13)) - 1)
            szHeader = Mid(szHeader, InStr(szHeader, ":") + 1)
        Else
            szHeader = "0"
        End If
        m_nFileSize = CLng(szHeader)
        dlflags = dlDownloading
        dl.Connect m_szServer, 80
    Case dlDownloading
        If InStr(LCase(szData), "content-type:") Then
            szData.Append Mid(szData, InStr(szData, vbCrLf & vbCrLf) + 4)
        End If
        m_szData.Append szData
        RaiseEvent Status("Downloading archive - " & m_szData.Length & " / " & m_nFileSize & " (" & CInt((m_szData.Length / m_nFileSize) * 100) & "%)")
        If m_szData.Length >= m_nFileSize Then
            dl.Close
            RaiseEvent Progress(100, m_nFileSize)
            RaiseEvent Status("Download complete")
            dl.Close
            SaveFile
            dlflags = dlIdle
        Else
            RaiseEvent Progress((m_szData.Length / m_nFileSize) * 100, m_szData.Length)
        End If
    End Select
End Sub

