VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form frmWebZip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WebZip 1.0"
   ClientHeight    =   8025
   ClientLeft      =   3600
   ClientTop       =   3375
   ClientWidth     =   12165
   Icon            =   "frmWebZip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   535
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   811
   Begin MSComDlg.CommonDialog dlgZip 
      Left            =   11520
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      CancelError     =   -1  'True
      DialogTitle     =   "Save Zip File"
      Filter          =   "Zip files|*.zip"
   End
   Begin MSComctlLib.ImageList ilstTbr 
      Left            =   10920
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWebZip.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWebZip.frx":075C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbr 
      Align           =   2  'Align Bottom
      Height          =   780
      Left            =   0
      TabIndex        =   9
      Top             =   6990
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   1376
      ButtonWidth     =   1879
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ilstTbr"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Download Zip"
            Object.ToolTipText     =   "Download entire ZIP archive"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   195
      Left            =   7920
      TabIndex        =   8
      Top             =   7815
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin InetCtlsObjects.Inet inCheck 
      Left            =   11520
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   811
      TabIndex        =   2
      Top             =   0
      Width           =   12165
      Begin VB.PictureBox pic16 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.ComboBox cboURL 
         Height          =   315
         Left            =   600
         TabIndex        =   3
         Text            =   "http://"
         Top             =   120
         Width           =   10215
      End
      Begin WebZip.ButtonEx cmdGet 
         Height          =   420
         Left            =   10920
         TabIndex        =   4
         Top             =   0
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   635
         Caption         =   "Get it!"
         CaptionOffsetY  =   -2
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   8454016
         SkinUp          =   "frmWebZip.frx":0EAE
      End
      Begin MSComctlLib.ImageList iml16 
         Left            =   480
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin WebZip.ButtonEx cmdStop 
         Height          =   420
         Left            =   10920
         TabIndex        =   7
         Top             =   360
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   635
         Caption         =   "Stop"
         CaptionOffsetY  =   -2
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   8454016
         SkinUp          =   "frmWebZip.frx":19C0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Feel free to use as you like, and remember, RANG3R FOUND THE ZIP ALGORITHMS! =)"
         Height          =   195
         Left            =   2520
         TabIndex        =   10
         Top             =   480
         Width           =   6390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "URL:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7770
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Status: Idle"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   6075
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   10716
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Filename"
         Text            =   "File Name"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "Index"
         Text            =   "Index"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "Size"
         Text            =   "Size"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   "Packed"
         Text            =   "Packed size"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   "CRC"
         Text            =   "CRC"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Byte Position"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmWebZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Icon Sizes in pixels
Private Const LARGE_ICON As Integer = 32
Private Const SMALL_ICON As Integer = 16
Private Const MAX_PATH = 260

Private Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Private Const SHGFI_LARGEICON = &H0       'Large icon
Private Const SHGFI_SMALLICON = &H1       'Small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type


'----------------------------------------------------------
'Functions & Procedures
'----------------------------------------------------------
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal i&, ByVal hdcDest&, _
    ByVal x&, ByVal Y&, ByVal Flags&) As Long


'----------------------------------------------------------
'Private variables
'----------------------------------------------------------
Private ShInfo As SHFILEINFO

Dim WithEvents zipReader As zipReader
Attribute zipReader.VB_VarHelpID = -1
Dim ZipFile As ZipFile
Dim file As file

Dim WithEvents ft As FileTransfer
Attribute ft.VB_VarHelpID = -1
        


Dim nLen As Long
Dim nPos As Long
Dim szThisURL As String
Dim szServer As String
Dim icmpPing As ICMP_ECHO_REPLY



   
Private Function GetIcon(Filename As String) As Long
    '---------------------------------------------------------------------
    'Extract an individual icon
    '---------------------------------------------------------------------
    Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
    Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
    Dim r As Long
    
    On Error Resume Next
    
    Dim szFile As String
    szFile = Right(Filename, Len(Filename) - (InStr(Filename, "/")))
    'MsgBox szFile
    'szFile = Filename
    Open szFile For Append As FreeFile: Close
    
    'Get a handle to the small icon
    hSIcon = SHGetFileInfo(szFile, 0&, ShInfo, Len(ShInfo), _
             BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    
    If FileLen(szFile) = 0 Then Kill szFile
    
    'MsgBox hSIcon
    
    'If the handle(s) exists, load it into the picture box(es)
    If hSIcon <> 0 Then
      With pic16
        Set .Picture = LoadPicture("")
        .AutoRedraw = True
        r = ImageList_Draw(hSIcon, ShInfo.iIcon, pic16.hdc, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      
    End If
End Function

Private Sub cmdAbout_Click()
    Load frmAbout
    frmAbout.Show
    
End Sub

Private Sub cmdDownload_Click()
    
End Sub

Private Sub cmdGet_Click()
    
    Dim szURL As String
    Dim szHeader As String
    
    szThisURL = cboURL.text
        
    szURL = Replace(LCase(cboURL.text), "http://", "")
    szServer = Trim(Left(szURL, InStr(szURL, "/") - 1))
    
    'Ping the server
    icmpPing.Data = String(32, Int(Rnd * 255) + 1)
    sbStatus.SimpleText = "Status: Reply from " & szServer & " -> " & Ping(szServer, icmpPing) & " ms"
    
    'Check to make sure file exists
    inCheck.Execute "http://" & szURL
Wait1:
    While inCheck.StillExecuting
        DoEvents
    Wend
    If inCheck.StillExecuting Then GoTo Wait1
    
    inCheck.GetChunk 1, icByteArray
    szHeader = inCheck.GetHeader
    If InStr(szHeader, "404 object not found") Then
        sbStatus.SimpleText = "Status: File not found"
        Exit Sub
    End If
    
    'Check to make sure resuming is supported
    'HTTP 1.1
    If Val(Mid(szHeader, 6, 3)) < 1.1 Then
        sbStatus.SimpleText = "Status: Cannot retrieve ZIP information (Cannot find server, or server does not support resuming)"
        Exit Sub
    End If
        
    'Continue
    Set lvFiles.SmallIcons = Nothing
    lvFiles.ListItems.Clear
    iml16.ListImages.Clear
    
    zipReader.Server = Trim(Left(szURL, InStr(szURL, "/") - 1))
    zipReader.Filename = Trim(Right(szURL, Len(szURL) - Len(zipReader.Server)))
        
    If Not ZipFile Is Nothing Then
        Do Until ZipFile.Files.Count = 0
            ZipFile.Files.Remove 1
            DoEvents
        Loop
    End If
        
    Set ZipFile = zipReader.GetFiles()
    
    For Each file In ZipFile.Files
        GetIcon file.Filename
        iml16.ListImages.Add file.Index, , pic16.Image
    Next
    
    Set lvFiles.SmallIcons = iml16
    
    For Each file In ZipFile.Files
        With lvFiles
            .ListItems.Add , , file.Filename
            .ListItems(.ListItems.Count).SubItems(1) = file.Index
            .ListItems(.ListItems.Count).SmallIcon = file.Index
            .ListItems(.ListItems.Count).SubItems(3) = IIf(CLng(file.PackedSize / 1024) >= 1, CLng(file.PackedSize / 1024) & " KB", "<1 KB")
            .ListItems(.ListItems.Count).SubItems(2) = IIf(CLng(file.RealSize / 1024) >= 1, CLng(file.RealSize / 1024) & " KB", "<1 KB")
            .ListItems(.ListItems.Count).SubItems(4) = file.CRC32
            .ListItems(.ListItems.Count).SubItems(5) = file.PacketPosition
        End With
    Next
    
    
End Sub

Private Sub cmdStop_Click()
    zipReader.DLStop
    
End Sub

Private Sub Form_Load()
    Set zipReader = New zipReader
    Set ft = New FileTransfer
    pic16.BackColor = lvFiles.BackColor
    Dim szVer As String
    SocketsInitialize szVer
    sbStatus.SimpleText = "Status: Idle (Winsock version " & szVer & ")"
    
    dlgZip.Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNReadOnly Or cdlOFNFileMustExist
    
End Sub

Private Sub ft_Progress(ByVal nPercent As Long, ByVal nReceivedTotal As Long)
    pbProgress.Value = nPercent
    
End Sub

Private Sub ft_Status(ByVal szStatus As String)
    sbStatus.SimpleText = "Status: " & szStatus
    If LCase(szStatus) = "download complete" Then pbProgress.Value = 0
    
    
End Sub

Private Sub lvFiles_DblClick()
    Dim file As file
    Set file = ZipFile.Files(lvFiles.SelectedItem.Index)
    szTemp = Right(lvFiles.SelectedItem.text, Len(lvFiles.SelectedItem.text) - (InStr(lvFiles.SelectedItem.text, "/")))
    If Dir(App.Path & "\" & szTemp & ".zip") <> "" Then
        Select Case MsgBox("The file """ & App.Path & "\" & szTemp & ".zip"" already exists. Do you want to overwrite it?", vbQuestion + vbYesNo, "WebZip")
        Case vbYes
            file.SaveAs App.Path & "\" & szTemp & ".zip"
            MsgBox "File saved as """ & App.Path & "\" & szTemp & ".zip""", vbInformation, "WebZip 1.0"
        End Select
    Else
        file.SaveAs App.Path & "\" & szTemp & ".zip"
        MsgBox "File saved as """ & App.Path & "\" & szTemp & ".zip""", vbInformation, "WebZip 1.0"
    End If
    
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Caption = "Download Zip" Then
        If szThisURL = "" Then
            MsgBox "No ZIP file chosen.", vbExclamation, "WebZip 1.0"
            Exit Sub
        End If
        
        dlgZip.ShowSave
        If Err = 0 Then
            ft.Download szThisURL, dlgZip.Filename
        End If
    ElseIf Button.Caption = "About" Then
        Load frmAbout
        frmAbout.Show
        
    End If
    
End Sub

Private Sub ZIPGet_DLComplete()
    sbStatus.SimpleText = "Status: Download complete"
    pbProgress.Value = 0
    
End Sub

Private Sub ZIPGet_Percent(lPercent As Long)
    pbProgress.Value = lPercent
    
End Sub

Private Sub ZIPGet_StatusChange(lpStatus As String)
    sbStatus.SimpleText = "Status: " & lpStatus
End Sub

Private Sub ZIPGet_TimeLeft(lpTime As String)
    sbStatus.SimpleText = "Status: Downloading ZIP package " & szThisURL & " (" & lpTime & " remaining)"
End Sub

Private Sub timPing_Timer()
    
End Sub

Private Sub zipReader_Inform(ByVal text As String)
    sbStatus.SimpleText = "Status: " & text
    
End Sub

Private Sub zipReader_ProgressChange(ByVal nProg As Integer)
    On Error Resume Next
    pbProgress.Value = nProg
    
End Sub
