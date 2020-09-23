Attribute VB_Name = "modVB6"
'###################################################################
'
' File:         modVB6.bas
'
' Function:     Provides VB6 functionality to VB5
'
' Description:  Adds extra methods to Visual Basic environment
'
' Author:       ULTiMaTuM (js)
'
' Environment:  Visual Basic version 5, Pentium II 400MHz 64mb RAM
'               Windows 98 SE 4.10.2222 A
'
' Notes:        Source code found on Planet Source Code, ASplit
'               function modified by js
'
' Revisions:    1.00  10/30/00 (js) First release
'
'###################################################################




Public Function InStrRev(Optional Start, Optional String1, Optional String2)

Dim lngLastPos As Long, lngPos As Long, lngStartChar As Long
Dim strString As String


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

InStrRev = lngLastPos&

End Function

Public Function ASplit(vBuf() As Variant, sIn As String, sDel As String) As Long
    Dim i As Integer, x As Integer, s As Integer, t As Integer
    i = 0: s = 1: t = 1: x = 0
    ReDim tArr(0) As Variant


    If InStr(1, sIn, sDel) <> 0 Then
        Do
            ReDim Preserve tArr(0 To x) As Variant
            tArr(i) = Mid(sIn, t, InStr(s, sIn, sDel) - t)
            t = InStr(s, sIn, sDel) + Len(sDel)
            s = t
            If tArr(i) <> "" Then i = i + 1
            x = x + 1
        Loop Until InStr(s, sIn, sDel) = 0
        ReDim Preserve tArr(0 To x) As Variant
        tArr(i) = Mid(sIn, t, Len(sIn) - t + 1)
    Else
        tArr(0) = sIn
    End If
    For i = LBound(tArr) To UBound(tArr)
        ReDim Preserve vBuf(0 To UBound(tArr)) As Variant
        vBuf(i) = tArr(i)
    Next
    ASplit = UBound(tArr)
End Function

Public Function Replace(ByVal strMain As String, strFind As String, strReplace As String) As String
    Dim lngSpot As Long, lngNewSpot As Long, strLeft As String
    Dim strRight As String, strNew As String
    lngSpot& = InStr(LCase(strMain$), LCase(strFind$))
    lngNewSpot& = lngSpot&
    Do
        If lngNewSpot& > 0& Then
            strLeft$ = Left(strMain$, lngNewSpot& - 1)
            If lngSpot& + Len(strFind$) <= Len(strMain$) Then
                strRight$ = Right(strMain$, Len(strMain$) - lngNewSpot& - Len(strFind$) + 1)
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
    Replace$ = strNew$
End Function
