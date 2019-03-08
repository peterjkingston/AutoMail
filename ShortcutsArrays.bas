Attribute VB_Name = "ShortCutsArrays"
Public Function Append2DArray(arrayA As Variant, arrayB As Variant, columncount As Integer) As Variant
    Dim newArray As Variant, r As Integer, c As Integer, totalRows As Integer
    newArray = Array()
    totalRows = UBound(arrayA) + UBound(arrayB) + 1
    ReDim newArray(totalRows, columncount) As Variant
    For r = 0 To UBound(arrayA)
        For c = 0 To columncount
            newArray(r, c) = arrayA(r, c)
        Next c
    Next r
    For r = UBound(arrayA) + 1 To totalRows
        For c = 0 To columncount
            newArray(r, c) = arrayB(r - 1 - UBound(arrayA), c)
        Next c
    Next r
    Append2DArray = newArray
End Function
Public Function RemoveFrom2DArray(thisArray As Variant, index As Integer, columncount As Integer) As Variant
    Dim newArray As Variant, r As Integer, c As Integer, i As Integer
    newArray = Array()
    ReDim newArray(UBound(thisArray) - 1, columncount - 1) As Variant
    For r = 0 To UBound(thisArray)
        For c = 0 To columncount - 1
            If r <> index Then
                newArray(i, c) = thisArray(r, c)
            End If
        Next c
        If r <> index Then: i = i + 1
    Next r
    RemoveFrom2DArray = newArray
End Function
Public Function Replace2DArray(mainArray As Variant, singleRowArray As Variant, index As Integer, aryUbound As Integer) As Variant
    Dim c As Integer
    For c = 0 To aryUbound
        mainArray(index, c) = singleRowArray(0, c)
    Next c
    Replace2DArray = mainArray
End Function
Public Function ShellSort2D(ByRef arr As Variant, Optional lastEl As Variant, _
    Optional descending As Boolean) As Variant
    Dim value() As Variant
    Dim index As Long, index2 As Long
    Dim firstEl As Long
    Dim distance As Long
    Dim numEls As Long
    Dim r As Integer, c As Integer
    ' account for optional arguments
    If IsMissing(lastEl) Then lastEl = UBound(arr)
    firstEl = LBound(arr)
    ReDim value(0, UBound(arr, 2)) As Variant
    numEls = lastEl - firstEl + 1
    ' find the best value for distance
    Do
        distance = distance * 3 + 1
    Loop Until distance > numEls

    Do
        distance = distance \ 3
        For index = distance + firstEl To lastEl
            For c = 0 To UBound(arr, 2)
                value(0, c) = arr(index, c)
            Next c
            index2 = index
            Do While (arr(index2 - distance, 0) > value(0, 0)) Xor descending
                arr(index2, 0) = arr(index2 - distance, 0)
                index2 = index2 - distance
                If index2 - distance < firstEl Then Exit Do
            Loop
            For c = 0 To UBound(arr, 2)
                arr(index2, c) = value(0, c)
            Next c
        Next
    Loop Until distance = 1
    
    ShellSort2D = arr
End Function

