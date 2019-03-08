Attribute VB_Name = "ShortcutsRanges"
Public Function NextRow(refersTo As String) As String
    Dim i As Integer, rowNum As Integer, newStr As String
    For i = 1 To Len(refersTo)
        If InStr(1, Mid(refersTo, i, Len(refersTo) - i), "$") = 0 Then
            rowNum = CInt(Mid(refersTo, i, Len(refersTo) - 1)) + 1
            Exit For
        Else
            newStr = newStr & Mid(refersTo, i, 1)
        End If
    Next i
    NextRow = newStr & CStr(rowNum)
End Function
Public Function LessRow(refersTo As String) As String
    Dim i As Integer, rowNum As Integer, newStr As String
    For i = 1 To Len(refersTo)
        If InStr(1, Mid(refersTo, i, Len(refersTo) - i), "$") = 0 Then
            rowNum = CInt(Mid(refersTo, i, Len(refersTo) - 1)) - 1
            Exit For
        Else
            newStr = newStr & Mid(refersTo, i, 1)
        End If
    Next i
    LessRow = newStr & CStr(rowNum)
End Function
Public Function AddTo(thisRng As Range, paramAry() As Variant) As Range
    Dim c As Integer
    Names(thisRng.Name.Name).refersTo = ShortcutsRanges.NextRow(Names(thisRng.Name.Name).refersTo)
    For c = 0 To UBound(paramAry, 2)
        thisRng(UBound(thisRng.value), c) = paramAry(0, c)
    Next c
End Function
Public Function RemoveFrom(thisRange As Range, index As Integer) As Range
    Dim c As Integer, r As Integer
    For r = index To UBound(thisRng.value, 1)
        For c = 0 To UBound(thisRng.value, 2)
            If Not (r + 1) > UBound(thisRng.value, 1) Then: thisRng(r, c) = thisRng(r + 1, c)
        Next c
    Next r
    Names(thisRng.Name.Name).refersTo = ShortcutsRanges.LessRow(Names(thisRng.Name.Name).refersTo)
End Function
Public Function ChangeRow(thisRange As Range, index As Integer, paramAry() As Variant) As Range
    Dim c As Integer
    For c = 0 To UBound(paramAry, 2)
        thisRng.value(index, c) = paramAry(index, c)
    Next c
End Function
