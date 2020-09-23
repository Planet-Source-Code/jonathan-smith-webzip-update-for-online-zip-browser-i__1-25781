VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZipFile"
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
' * Time             : 20:34
' * Module Name      : ZipFile
' * Module Filename  : ZipFile.cls
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Private mvarArchiveSize As Currency 'local copy
Private mvarFiles       As Collection 'local copy
Friend Property Set Files(ByVal vData As Collection)
   'used when assigning an Object to the property, on the left side of a Set statement.
   'Syntax: Set x.Files = Form1
   Set mvarFiles = vData
End Property

Public Property Get Files() As Collection
   'used when retrieving value of a property, on the right side of an assignment.
   'Syntax: Debug.Print X.Files
   Set Files = mvarFiles
End Property

Public Property Get FileCount() As Long
   'used when retrieving value of a property, on the right side of an assignment.
   'Syntax: Debug.Print X.FileCount
   FileCount = Files.Count
End Property


Friend Property Let ArchiveSize(ByVal vData As Currency)
   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.ArchiveSize = 5
   mvarArchiveSize = vData
End Property

Public Property Get ArchiveSize() As Currency
   'used when retrieving value of a property, on the right side of an assignment.
   'Syntax: Debug.Print X.ArchiveSize
   ArchiveSize = mvarArchiveSize
End Property
