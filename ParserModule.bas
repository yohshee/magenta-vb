Attribute VB_Name = "Parser"
Option Explicit

' Splits a string up based on a delimiter and places the contents
' into a string array. Note that the delimiter can be more than one character.
Public Sub SplitString(ByVal szString As String, szArray() As String, Optional ByVal szDelimiter As String = " ", Optional ByVal lCount As Long = 0)
Dim i As Long
Dim Pos As Long
Dim StartPos As Long
Dim Length As Long

ReDim szArray(1 To 1)

Pos = InStr(1, szString, szDelimiter)
If Pos = 0 Then
    ' Ok, we don't need to do anything to this string;
    ' simply put the string into the array and leave it alone.
    ReDim szArray(1 To 1) As String
    szArray(1) = szString
    Exit Sub
End If

StartPos = 1
Length = Pos - 1
i = 1
Do
    'Fix the array
    ReDim Preserve szArray(1 To i)
    If i = lCount And lCount <> 0 Then
        ' Prematurely exit and don't parse the rest of it.
        szArray(i) = Trim$(Mid$(szString, StartPos))
        Exit Do
    End If
    Length = Pos - StartPos + IIf(Pos = Len(szString), 1, 0)
    szArray(i) = Trim$(Mid$(szString, StartPos, Length))
    StartPos = Pos + Len(szDelimiter)
    'If StartPos = Len(szString) Then StartPos = StartPos + 1
    Pos = InStr(StartPos, szString, szDelimiter)
    Pos = IIf(Pos = 0, Len(szString), Pos)
    i = i + 1
Loop Until StartPos = Len(szString) + 1
End Sub

Public Function InStrLast(ByVal Start As Long, ByVal Source As String, ByVal Search As String, _
    Optional Cmp As VbCompareMethod) As Long
    Start = Start - 1
    Do
        Start = InStr(Start + 1, Source, Search, Cmp)
        If Start = 0 Then Exit Function
        InStrLast = Start
    Loop
End Function

Public Function SearchAndReplace(ByVal Start As Long, ByVal Source As String, ByVal Search As String, _
    Replace As String, Optional Cmp As VbCompareMethod) As String
    Dim lTemp As Long
    
    Start = Start - 1
    lTemp = Len(Search)
    Do
        Start = InStr(Start + 1, Source, Search, Cmp)
        If Start = 0 Then Exit Do
        If lTemp > 1 Then
            Source = Mid$(Source, 1, Start - 1) & Replace & Mid$(Source, Start + Len(Replace))
        Else
            Source = Mid$(Source, 1, Start - 1) & Replace & Mid$(Source, Start + Len(Replace) - 1)
        End If
    Loop
SearchAndReplace = Source
End Function

