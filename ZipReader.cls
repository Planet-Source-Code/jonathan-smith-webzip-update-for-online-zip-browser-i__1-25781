VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZipReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' #VBIDEUtils#************************************************************
' * Programmer Name  : Roger Johansson
' * Web Site         : http://www.sourcecode4free.com
' * E-Mail           : rang3r@sourcecode4free.com
' * Date             : 05-01-2001
' * Time             : 20:30
' * Module Name      : ZipReader
' * Module Filename  : ZipReader.cls
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************


Dim WithEvents theSock  As Winsock
Attribute theSock.VB_VarHelpID = -1
Dim mCol                As Collection
Dim Data                As String
Public ReadFrom            As Currency
Public ReadBytes           As Currency
Dim Ready               As Boolean
Dim OnlyHeader          As Boolean
Public Filename         As String
Public Server           As String

Public Event Inform(ByVal text As String)
Public Event ProgressChange(ByVal nProg As Integer)

Public Function GetChunk(ByVal StartPos As Currency, ByVal EndPos As Currency) As String
    ReadFrom = StartPos
    ReadBytes = EndPos - StartPos
    RaiseEvent Inform("Starting to download")
    Download
    RaiseEvent Inform("Downloading chunk " & StartPos & " - " & EndPos)
    GetChunk = Data
    RaiseEvent Inform("Download complete")
End Function

Private Sub Class_Initialize()
   Set theSock = New Winsock
   Set mCol = New Collection
End Sub

Private Sub theSock_Connect()

   Dim cmd              As String
   cmd = cmd & "GET " & Me.Filename & " HTTP/1.1" & vbCrLf
   cmd = cmd & "Accept: *.*, */*" & vbCrLf
   If ReadFrom > -1 Then
      cmd = cmd & "Range: bytes=" & ReadFrom & "-" & ReadFrom + ReadBytes & vbCrLf
   End If
   cmd = cmd & "User-Agent: WebZip" & vbCrLf
   cmd = cmd & "Referer: " & theSock.LocalIP & vbCrLf
   cmd = cmd & "Host: " & theSock.LocalIP & vbCrLf & vbCrLf
   theSock.SendData cmd
End Sub


Private Sub theSock_DataArrival(ByVal bytesTotal As Long)
   Dim strData          As String
   theSock.GetData strData, vbString

   If InStr(LCase(strData), "content-type:") Then    'remove header
      If OnlyHeader = True Then
         theSock.Close
         Data = Left(strData, InStr(strData, vbCrLf & vbCrLf))
         Ready = True
         Exit Sub
      End If

      strData = Mid(strData, InStr(strData, vbCrLf & vbCrLf) + 4)
   End If
   Data = Data & strData
   
   'WebZip specific to the form
    RaiseEvent Inform("Downloading package - " & Len(Data) & " / " & ReadBytes & " (" & CInt((Len(Data) / ReadBytes) * 100) & " %)")
    RaiseEvent ProgressChange(CInt((Len(Data) / ReadBytes) * 100))
   
   If Len(Data) >= ReadBytes Then
      'close the connection now...
      theSock.Close
      Data = Left(Data, ReadBytes)
      RaiseEvent ProgressChange(0)
      Ready = True
   End If

End Sub

Public Sub DLStop()
    theSock.Close
    Data = ""
    RaiseEvent Inform("Download cancelled")
    Ready = True
End Sub


Private Sub Download()
   Ready = False
   theSock.Close
   Do
      DoEvents
   Loop While theSock.State <> sckClosed

   Data = ""
   theSock.Connect Me.Server, 80
   Do While Not Ready
      DoEvents
   Loop
End Sub

Public Function GetFiles() As ZipFile
   Dim FileCount        As Long
   Dim ArchiveSize      As Long
   Dim LibSize          As Long
   Dim FileSize         As Long
   Dim LibPackage       As String
   Dim FileNameLen      As Long
   Dim i                As Long
   Dim FilePos          As Long
   Dim file             As file
   Dim ZipFile          As New ZipFile

   
   FileSize = SizeOf("")
   ReadFrom = FileSize - 22
   ReadBytes = 22
   OnlyHeader = False
   RaiseEvent Inform("Checking file size")
   Download
   RaiseEvent Inform("File size found")
   LibPackage = Data
   FileCount = WordVal(9)
   LibSize = LongVal(13)
   ReadFrom = FileSize - 22 - LibSize
   ReadBytes = LibSize
   RaiseEvent Inform("Downloading package")
   Download
   RaiseEvent Inform("Download complete")
   RaiseEvent Inform("Extracting filenames")
   For i = 1 To FileCount
      FileNameLen = WordVal(29 + FilePos)

      Set file = New file
      file.LibData = Mid(Data, FilePos + 1, 46 + FileNameLen)
      file.Filename = Mid(Data, 47 + FilePos, FileNameLen)
      file.Index = i
      file.CRC32 = LongVal(16 + FilePos)
      'File.DataPosition = 123
      file.PacketPosition = LongVal(43 + FilePos)
      
      file.PackedSize = LongVal(21 + FilePos)
      file.RealSize = LongVal(25 + FilePos)
      
      
      mCol.Add file
      FilePos = FilePos + 46 + FileNameLen
      Set file.Parent = Me
   Next
   Set ZipFile.Files = mCol
   ZipFile.ArchiveSize = ArchiveSize
   RaiseEvent Inform("Done")
   Set GetFiles = ZipFile
End Function

Private Function LongVal(BytNo As Long) As Currency
    On Error Resume Next
   Dim bt1              As Currency
   Dim bt2              As Currency
   Dim bt3              As Currency
   Dim bt4              As Currency

   bt1 = Asc(Mid(Data, BytNo + 0, 1))
   bt2 = Asc(Mid(Data, BytNo + 1, 1))
   bt3 = Asc(Mid(Data, BytNo + 2, 1))
   bt4 = Asc(Mid(Data, BytNo + 3, 1))

   bt1 = bt1
   bt2 = bt2 * &H100
   bt3 = bt3 * &H10000
   bt4 = bt4 * &H1000000

   LongVal = bt1 + bt2 + bt3 + bt4
End Function

Private Function WordVal(BytNo As Long) As Currency
   On Error Resume Next
   Dim bt1              As Currency
   Dim bt2              As Currency

   bt1 = Asc(Mid(Data, BytNo + 0, 1))
   bt2 = Asc(Mid(Data, BytNo + 1, 1))

   bt1 = bt1
   bt2 = bt2 * &H100

   WordVal = bt1 + bt2
End Function

Public Function SizeOf(ByVal file As String) As Long
   Dim header           As String
   Dim ContentRow       As String
   ReadFrom = -1
   OnlyHeader = True
   Download
   header = LCase(Data)

   If InStr(header, "content-length") > 0 Then
      header = Mid(header, InStr(header, "content-length"))
      header = Left(header, InStr(header, Chr(13)) - 1)
      header = Mid(header, InStr(header, ":") + 1)
   Else
      header = "0"
   End If
   SizeOf = CLng(header)

End Function
