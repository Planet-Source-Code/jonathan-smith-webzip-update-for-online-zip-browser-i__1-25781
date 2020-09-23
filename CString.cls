VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const NONE = 0
Const STRINGTYPE = 1
Const INTEGERTYPE = 2
Const LONGTYPE = 3
Const FLOATTYPE = 4
Const CHARPERCENT = 5

Private m_szValue As String


Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    Value = m_szValue
    
End Property

Public Property Let Value(ByVal vData As String)
    m_szValue = CStr(vData)
    
End Property

Public Property Get Length() As Long
    Length = Len(m_szValue)
End Property

Public Sub Append(ByVal szString As String)
    If TypeName(szString) <> "CString" And TypeName(szString) <> "String" Then
        Err.Raise 93, "CString", "Appending object must be a String or CString."
        Exit Sub
    End If

    m_szValue = m_szValue & szString
    
End Sub

Public Function CountSubstring(ByVal strFind As String) As Long
    Dim strInput As String
    strInput = m_szValue
    CountSubstring = (Len(strInput) - Len(InterfaceReplace(strInput, strFind, ""))) / Len(strFind)
    
End Function

Public Function Compare(ByVal szString As String)
    
    Dim nValue As Long
    Dim nValue2 As Long
    Dim i
    
    If TypeName(szString) <> "CString" And TypeName(szString) <> "String" Then
        Err.Raise 93, "CString", "Comparing object must be a String or CString."
        Exit Function
    End If
    
    If Not Len(szString) = 0 And Not Len(m_szValue) = 0 Then
        For i = 1 To Len(m_szValue)
            nValue = nValue + CLng(Asc(VBA.Mid(m_szValue, i, 1)))
        Next
        For i = 1 To Len(szString)
            nValue2 = nValue2 + CLng(Asc(VBA.Mid(szString, i, 1)))
        Next
    End If
    
    If nValue - nValue2 < 0 Then
        Compare = -1
    ElseIf nValue - nValue2 > 0 Then
        Compare = 1
    Else
        Compare = 0
    End If
    
End Function

Public Function CompareNoCase(ByVal szString As String)
    
    Dim szMax As String
    Dim szMax2 As String
    Dim i
    
    Dim nValue As Long
    Dim nValue2 As Long
        
    If TypeName(szString) <> "CString" And TypeName(szString) <> "String" Then
        Err.Raise 93, "CString", "Comparing object must be a String or CString."
        Exit Function
    End If
    
    szMax = UCase(m_szValue)
    szMax2 = UCase(szString)
    
    
    If Not Len(szString) = 0 And Not Len(m_szValue) = 0 Then
        For i = 1 To Len(m_szValue)
            nValue = nValue + CLng(Asc(VBA.Mid(szMax, i, 1)))
        Next
        For i = 1 To Len(szString)
            nValue2 = nValue2 + CLng(Asc(VBA.Mid(szMax2, i, 1)))
        Next
    End If
    
    If nValue - nValue2 < 0 Then
        CompareNoCase = -1
    ElseIf nValue - nValue2 > 0 Then
        CompareNoCase = 1
    Else
        CompareNoCase = 0
    End If
    
End Function

Public Function Equals(ByVal szString As Variant) As Boolean
    If TypeName(szString) <> "CString" And TypeName(szString) <> "String" Then
        Err.Raise 93, "CString", "Equalizing object must be a String or CString."
        Exit Function
    End If
    
    Equals = (szString = m_szValue)
    
End Function

Public Function EqualsNoCase(ByVal szString As Variant) As Boolean
    If TypeName(szString) <> "CString" And TypeName(szString) <> "String" Then
        Err.Raise 93, "CString", "Equalizing object must be a String or CString."
        Exit Function
    End If
    
    EqualsNoCase = (LCase(szString) = LCase(m_szValue))
    
End Function


Public Function GetAt(ByVal nWhere As Long) As String
    If nWhere > Len(m_szValue) Then GetAt = ""
    GetAt = VBA.Mid(m_szValue, nWhere, 1)
        
End Function

Public Sub SetAt(ByVal nWhere As Long, ByVal sChar As String)
    sChar = VBA.Left(sChar, 1)
    
    Dim szTemp As String
    Dim i As Integer

    If nWhere > Len(m_szValue) Then
        szTemp = Space(nWhere)
        Mid(szTemp, i, Len(m_szValue)) = m_szValue
    Else
        szTemp = m_szValue
    End If
    Mid(szTemp, nWhere, 1) = sChar
            
End Sub

Public Function IsEmpty() As Boolean
    If Len(m_szValue) = 0 Then IsEmpty = True
    
End Function

Public Sub MakeEmpty()
    m_szValue = ""
End Sub

Public Function Mid(ByVal nFirst As Long, Optional ByVal nCount As Long) As String
    Mid = CStr(VBA.Mid(m_szValue, nFirst, nCount))
End Function

Public Function Left(ByVal nCount As Long) As String
    Left = CStr(VBA.Left(m_szValue, nCount))
End Function

Public Function Right(ByVal nCount As Long) As String
    Right = CStr(VBA.Right(m_szValue, nCount))
End Function

Public Function SpanIncluding(ByVal szCharSet As String) As String
    Dim szRet As String
    Dim i
    
    
    If Not Len(m_szValue) > 0 Or Not Len(szCharSet) > 0 Then
        Exit Function
    End If
    
    For i = 1 To Len(m_szValue)
        If InStr(szCharSet, VBA.Mid(m_szValue, i, 1)) <> 0 Then
            szRet = szRet & VBA.Mid(m_szValue, i, 1)
        End If
    Next
    
    SpanIncluding = szRet
    
End Function

Public Function SpanExcluding(ByVal szCharSet As String) As String
    Dim szRet As String
    Dim i
    
    If Not Len(m_szValue) > 0 Or Not Len(szCharSet) > 0 Then
        Exit Function
    End If
    
    For i = 1 To Len(m_szValue)
        If InStr(szCharSet, VBA.Mid(m_szValue, i, 1)) = 0 Then
            szRet = szRet & VBA.Mid(m_szValue, i, 1)
        End If
    Next
    
    SpanExcluding = szRet
    
End Function

Public Sub MakeUpper()
    m_szValue = UCase(m_szValue)
End Sub

Public Sub MakeLower()
    m_szValue = LCase(m_szValue)
End Sub

Public Sub MakeReverse()
    Dim szTemp As String
    Dim i
    
    If Len(m_szValue) = 0 Then Exit Sub
    
    For i = Len(m_szValue) To 1 Step -1
        szTemp = szTemp & VBA.Mid(m_szValue, i, 1)
    Next
    
    m_szValue = szTemp
    
    
End Sub

Public Sub Replace(strFind As String, strReplace As String)
    m_szValue = InterfaceReplace(m_szValue, strFind, strReplace)
End Sub

Public Sub Remove(ByVal szChar As String)
    If Len(szChar) > 1 Then szChar = VBA.Left(szChar, 1)
    
    Replace szChar, ""
    
End Sub

Public Sub Insert(ByVal nIndex As Long, ByVal szStr As String)
    Dim szLeft As String
    Dim szRight As String
    
    If nIndex > 1 And Len(m_szValue) > 0 Then
        szLeft = Left(nIndex - 1)
    Else
        szLeft = ""
    End If
    
    szRight = Right(Len(m_szValue) - nIndex + 1)
    
    m_szValue = szLeft & szStr & szRight
            
End Sub

Public Sub Delete(ByVal nIndex As Long, Optional nCount As Long = 1)
    Dim sLeft As String, sRight As String
    Dim nLen As Integer
    
    nLen = Len(m_szValue)


    If nIndex >= 0 And nIndex <= nLen Then


        If nIndex > 1 And nLen > 0 Then
            sLeft = Left(nIndex - 1)
        Else
            sLeft = ""
        End If

        If (nIndex + nCount) <= nLen Then
            sRight = VBA.Mid(m_szValue, nIndex + nCount)
        Else
            sRight = ""
        End If

        m_szValue = sLeft & sRight
        
    End If

End Sub

Public Sub TrimLeft()
    LTrim m_szValue
End Sub

Public Sub TrimRight()
    RTrim m_szValue
End Sub

Public Function Find(Optional ByVal nStart As Long = 1, Optional ByVal szSubstr As String) As Long
    Find = InStr(nStart, m_szValue, szSubstr)
End Function

Public Function ReverseFind(ByVal szSubstr As String) As Long
    ReverseFind = InterfaceInStrRev(m_szValue, szSubstr)
End Function

Public Function FindOneOf(ByVal szCharSet As String) As Long
    If Not Len(m_szValue) > 0 Or Not Len(szCharSet) > 0 Then
        Exit Function
    End If
    
    Dim iPos As Long
    Dim i As Integer
    
    For i = 1 To Len(m_szValue)
        iPos = InStr(szCharSet, VBA.Mid(m_szValue, i, 1))
        If iPos <> 0 Then
            FindOneOf = iPos
            Exit Function
        End If
    Next
    
End Function

Public Sub AllocSysString(ByRef szString As String)
    szString = Space$(Length)
    
End Sub

Public Sub SetSysString(ByRef szString As String)
    szString = m_szValue
End Sub

Public Function Split(ByRef vBuf() As Variant, szDelim As String) As Long
    Split = InterfaceSplit(vBuf, m_szValue, szDelim)
End Function

Public Function GetParameter(ByVal szFormat As String, ByVal nRef As Integer) As String
    'This function doesn't fully work. If you want to fix it and post it, go ahead. Just remember
    'who wrote it originally =)
    
    Dim szTemp As String
    Dim nPos As Integer
    Dim thing As Variant
    Dim szBuf()
    nPos = 1
    szTemp = m_szValue
    
    If Not m_szValue Like szFormat Then Exit Function
    
    If VBA.Left(szFormat, 1) = "*" Then szFormat = " " & szFormat
    
    InterfaceSplit szBuf, szFormat, "*"
    
    For Each thing In szBuf
        szTemp = InterfaceReplace(szTemp, (thing), Chr(255) & Chr(1))
    Next
    
    InterfaceSplit szBuf, szTemp, Chr(255) & Chr(1)
    
    If nRef - 1 < LBound(szBuf) Or nRef - 1 > UBound(szBuf) Then Exit Function
    
    GetParameter = szBuf(nRef - 1)
    
    
End Function


Public Sub Sprintf(DefString As String, ParamArray TheVals() As Variant)


Dim DefLen As Integer, DefIdx As Integer
Dim CurIdx As Integer, WorkString As String
Dim CurVal As Integer, MaxVal As Integer
Dim CurFormat As String, ValCount As Integer
Dim xIndex As Integer, FoundV As Boolean, vType As Integer
Dim CurParm As String

    DefLen = Len(DefString)
    DefIdx = 1
    CurVal = 0
    MaxVal = UBound(TheVals) + 1
    ValCount = 0

    ' Check for equal number of 'flags' as values, raise an error if inequal

    Do
        CurIdx = InStr(DefIdx, DefString, "%")
        If CurIdx > 0 Then
            If VBA.Mid$(DefString, CurIdx + 1, 1) <> "%" Then ' don't count %%, will be converted to % later
                ValCount = ValCount + 1
                DefIdx = CurIdx + 1
            Else
                DefIdx = CurIdx + 2
            End If
        Else
            Exit Do
        End If
    Loop
    
    If ValCount <> MaxVal Then Err.Raise 450, , "Mismatch of parameters for string " & DefString & ".  Expected " & ValCount & " but received " & MaxVal & "."
       
    DefIdx = 1
    CurVal = 0
    ValCount = 0
    
    WorkString = ""
    
    Do
        CurIdx = InStr(DefIdx, DefString, "%")
        If CurIdx <> 0 Then
        ' First, get the variable identifier.  Scan from Defidx (the %) to EOL looking for the
        ' first occurance of s, d, l, or f
            FoundV = False
            vType = NONE
            xIndex = CurIdx + 1
            Do While FoundV = False
                If Not FoundV Then
                    CurParm = VBA.Mid$(DefString, xIndex, 1)
                    Select Case VBA.Mid$(DefString, xIndex, 1)
                        Case "%"
                            vType = CHARPERCENT
                            FoundV = True
                            CurIdx = CurIdx + 1
                            CurVal = xIndex + 2
                        Case "s"
                            vType = STRINGTYPE
                            FoundV = True
                            CurVal = xIndex + 1
                        Case "d"
                            vType = INTEGERTYPE
                            FoundV = True
                            CurVal = xIndex + 1
                        Case "l"
                            If VBA.Mid$(DefString, xIndex + 1, 1) = "d" Then
                                vType = LONGTYPE
                                FoundV = True
                                CurVal = xIndex + 2
                            Else
                                Err.Raise 93, , "Unrecognized pattern " & VBA.Mid$(DefString, xIndex - 1, 3) & " in " & DefString
                            End If
                        Case "f"
                            vType = FLOATTYPE
                            FoundV = True
                            CurVal = xIndex + 1
                    End Select
                End If
                If Not FoundV Then xIndex = xIndex + 1
            Loop
            If Not FoundV Then Err.Raise 93, , "Invalid % format in " & DefString
            CurParm = VBA.Mid$(DefString, CurIdx, CurVal - CurIdx) ' For debugging purposes
            
            If vType = CHARPERCENT Then
                WorkString = WorkString & VBA.Mid$(DefString, DefIdx, CurIdx - DefIdx)
                CurVal = CurVal - 1
            Else
                CurFormat = BuildFormat(CurParm, vType)
                WorkString = WorkString & VBA.Mid$(DefString, DefIdx, CurIdx - DefIdx) & Format$(TheVals(ValCount), CurFormat)
                ValCount = ValCount + 1
            End If
            DefIdx = CurVal
        Else
            WorkString = WorkString & VBA.Right$(DefString, Len(DefString) - DefIdx + 1)
            Exit Do
        End If
    Loop
    m_szValue = TreatBackSlash(WorkString)
End Sub


'********************************************************
'Utility Functions
'********************************************************

Private Function BuildFormat(Parm As String, DataType As Integer) As String
    Dim Prefix As String, TmpFmt As String

    If DataType = LONGTYPE Then Prefix = VBA.Mid$(Parm, 2, Len(Parm) - 3) Else Prefix = VBA.Mid$(Parm, 2, Len(Parm) - 2)

    Select Case InStr(Prefix, ".")
        Case 0, Len(Prefix)
            If VBA.Left$(Prefix, 1) = "0" Then TmpFmt = String(CInt(Prefix), "0") Else TmpFmt = "#"
        Case 1
            If VBA.Mid$(Prefix, 2, 1) = "0" Then TmpFmt = "#." & String(CInt(VBA.Right$(Prefix, 2)), "0") Else TmpFmt = "#.#"
        Case Else
            If VBA.Left$(Prefix, 1) = "0" Then TmpFmt = String(CInt(VBA.Left$(Prefix, InStr(Prefix, "."))), "0") & "." Else TmpFmt = "#."
            If VBA.Mid$(Prefix, InStr(Prefix, ".") + 1, 1) = "0" Then TmpFmt = TmpFmt & String(CInt(VBA.Right$(Prefix, InStr(Prefix, ".") - 1)), "0") Else TmpFmt = TmpFmt & "#"
    End Select

    BuildFormat = TreatBackSlash(TmpFmt)
End Function

Private Function TreatBackSlash(sLine As String) As String
      TreatBackSlash = sLine
      TreatBackSlash = InterfaceReplace(TreatBackSlash, "\n", vbCrLf)
      TreatBackSlash = InterfaceReplace(TreatBackSlash, "\r", vbCr)
      TreatBackSlash = InterfaceReplace(TreatBackSlash, "\t", vbTab)
      TreatBackSlash = InterfaceReplace(TreatBackSlash, "\b", vbBack)
      TreatBackSlash = InterfaceReplace(TreatBackSlash, "\0", vbNullString)
      TreatBackSlash = InterfaceReplace(TreatBackSlash, "\\", "\")
         
End Function

'If you have VB6, the proceeding functions are unnecessary
Private Function InterfaceReplace(ByVal strMain As String, strFind As String, strReplace As String) As String
    Dim lngSpot As Long, lngNewSpot As Long, strLeft As String
    Dim strRight As String, strNew As String
    lngSpot& = InStr(LCase(strMain$), LCase(strFind$))
    lngNewSpot& = lngSpot&
    Do
        If lngNewSpot& > 0& Then
            strLeft$ = VBA.Left(strMain$, lngNewSpot& - 1)
            If lngSpot& + Len(strFind$) <= Len(strMain$) Then
                strRight$ = VBA.Right(strMain$, Len(strMain$) - lngNewSpot& - Len(strFind$) + 1)
            Else
                strRight = ""
            End If
            strNew$ = strLeft$ & strReplace$ & strRight$
            strMain$ = strNew$
        Else
            strNew$ = strMain$
        End If
        lngSpot& = lngNewSpot& + Len(strReplace$)
        If lngSpot& > 0 Then
            lngNewSpot& = InStr(lngSpot&, LCase(strMain$), LCase(strFind$))
        End If
    Loop Until lngNewSpot& < 1
    InterfaceReplace$ = strNew$
End Function

Private Function InterfaceInStrRev(Optional Start, Optional String1, Optional String2)

Dim lngLastPos As Long, lngPos As Long, lngStartChar As Long
Dim strString As String
Dim strSearchString$


  'check to see if String2 is missing. If yes, then
  'the start argument wasn't given so automatically
  'give it the value of the length of String1.
  If IsMissing(String2) Then
    lngStartChar& = Len(Start)
    strString$ = CStr(Start)
    strSearchString$ = CStr(String1)
  Else
    lngStartChar& = CLng(Start)
    strString$ = CStr(String1)
    strSearchString$ = CStr(String2)
  End If

'if the string can't be found then exit
If InStr(strString$, strSearchString$) = 0 Then Exit Function

'loop through the text until lngPos is bigger than Start or equal to 0.
'then return the character position prior to that.

 Do
   DoEvents
   lngPos& = InStr(lngLastPos& + 1, strString$, strSearchString$)
   If lngPos& > lngStartChar& Or lngPos& = 0 Then Exit Do
   lngLastPos& = lngPos&
 Loop

InterfaceInStrRev = lngLastPos&

End Function

Private Function InterfaceSplit(vBuf() As Variant, sIn As String, sDel As String) As Long
    Dim i As Integer, x As Integer, s As Integer, t As Integer
    i = 0: s = 1: t = 1: x = 0
    ReDim tArr(0) As Variant


    If InStr(1, sIn, sDel) <> 0 Then
        Do
            ReDim Preserve tArr(0 To x) As Variant
            tArr(i) = VBA.Mid(sIn, t, InStr(s, sIn, sDel) - t)
            t = InStr(s, sIn, sDel) + Len(sDel)
            s = t
            If tArr(i) <> "" Then i = i + 1
            x = x + 1
        Loop Until InStr(s, sIn, sDel) = 0
        ReDim Preserve tArr(0 To x) As Variant
        tArr(i) = VBA.Mid(sIn, t, Len(sIn) - t + 1)
    Else
        tArr(0) = sIn
    End If
    For i = LBound(tArr) To UBound(tArr)
        ReDim Preserve vBuf(0 To UBound(tArr)) As Variant
        vBuf(i) = tArr(i)
    Next
    InterfaceSplit = UBound(tArr)
End Function
