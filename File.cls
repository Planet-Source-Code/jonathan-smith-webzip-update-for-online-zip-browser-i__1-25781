VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarFileName    As String 'local copy
Private mvarPackedSize  As Currency 'local copy
Private mvarRealSize    As Currency 'local copy
Private mvarIndex       As Long 'local copy
Private mvarCRC32       As Currency 'local copy
Private mvarPacketPosition As Currency 'local copy
Private mvarDataPosition As Currency 'local copy
'local variable(s) to hold property value(s)
Private mvarParent As zipReader 'local copy
'local variable(s) to hold property value(s)
Private mvarLibData As String 'local copy

Friend Property Let LibData(ByVal vData As String)
    On Error Resume Next
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LibData = 5
    mvarLibData = vData
    'since we always only extract ONE file .. we can always set the packet position = &h 0000 0000
    Mid(mvarLibData, 43, 4) = Chr(0) & Chr(0) & Chr(0) & Chr(0)
    '(and YES you can use MID to assign values in teh middle of a string)
End Property


Friend Property Get LibData() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LibData
    LibData = mvarLibData
End Property



Friend Property Set Parent(ByVal vData As zipReader)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Parent = Form1
    Set mvarParent = vData
End Property


Friend Property Get Parent() As zipReader
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Parent
    Set Parent = mvarParent
End Property



Public Property Let DataPosition(ByVal vData As Currency)
   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.DataPosition = 5
   mvarDataPosition = vData
End Property

Public Property Get DataPosition() As Currency
   'used when retrieving value of a property, on the right side of an assignment.
   'Syntax: Debug.Print X.DataPosition
   DataPosition = mvarDataPosition
End Property


Public Property Let PacketPosition(ByVal vData As Currency)
   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.PacketPosition = 5
   mvarPacketPosition = vData
End Property

Public Property Get PacketPosition() As Currency
   'used when retrieving value of a property, on the right side of an assignment.
   'Syntax: Debug.Print X.PacketPosition
   PacketPosition = mvarPacketPosition
End Property


Public Property Let CRC32(ByVal vData As Currency)
   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.CRC32 = 5
   mvarCRC32 = vData
End Property

Public Property Get CRC32() As Currency
   'used when retrieving value of a property, on the right side of an assignment.
   'Syntax: Debug.Print X.CRC32
   CRC32 = mvarCRC32
End Property


Public Property Let Index(ByVal vData As Long)
  'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.Index = 5
   mvarIndex = vData
End Property

Public Property Get Index() As Long
   'used when retrieving value of a property, on the right side of an assignment.
   'Syntax: Debug.Print X.Index
   Index = mvarIndex
End Property


Public Property Let RealSize(ByVal vData As Currency)
   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.RealSize = 5
   mvarRealSize = vData
End Property

Public Property Get RealSize() As Currency
   'used when retrieving value of a property, on the right side of an assignment.
   'Syntax: Debug.Print X.RealSize
   RealSize = mvarRealSize
End Property


Public Property Let PackedSize(ByVal vData As Currency)
   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.PackedSize = 5
   mvarPackedSize = vData
End Property

Public Property Get PackedSize() As Currency
   'used when retrieving value of a property, on the right side of an assignment.
   'Syntax: Debug.Print X.PackedSize
   PackedSize = mvarPackedSize
End Property


Public Property Let Filename(ByVal vData As String)
   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.FileName = 5
   mvarFileName = vData
End Property

Public Property Get Filename() As String
   'used when retrieving value of a property, on the right side of an assignment.
   'Syntax: Debug.Print X.FileName
   Filename = mvarFileName
End Property


Public Sub SaveAs(ByVal Filename As String)
    Dim file As String
    'Dim strLastPackage As String
    Dim bytLastPackage(21) As Byte
    Dim i As Long
    file = Parent.GetChunk(Me.PacketPosition, Me.PacketPosition + Me.PackedSize + 30 + Len(Me.Filename))
    
    'LastPAckage = "50 4B 05 06 00 00 00 00 01 00 01 00 3A 00 00 00 67 00 00 00 00 00"
    file = file & Me.LibData
    bytLastPackage(0) = &H50    'header for the last package
    bytLastPackage(1) = &H4B
    bytLastPackage(2) = &H5
    bytLastPackage(3) = &H6
    
    bytLastPackage(8) = &H1         ' only one file in the lib
    bytLastPackage(10) = &H1
    
    bytLastPackage(12) = Len(Me.LibData)    'size of lib
    
    bytLastPackage(16) = &H3A   'pos of the lib (?? not sure what this baby does(ripped from a zip file))
    
    For i = 0 To 21
        file = file & Chr(bytLastPackage(i))
    Next
    
    Filename = Replace(Filename, "/", "\")
    
    On Error Resume Next
    Kill Filename
    Err.Clear
    On Error GoTo 0
    Open Filename For Binary As #1
    Put #1, , file
    Close #1
    
End Sub
