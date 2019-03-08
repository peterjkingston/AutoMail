Attribute VB_Name = "hTimedTest"
Private Sub TimeThis()
    Dim thisTime As Double, PDDoc As AcroPDDoc, AVDoc As AcroAVDoc, App As AcroApp
    Dim i As Integer
    Set App = CreateObject("AcroExch.App"): Set AVDoc = CreateObject("AcroExch.AVDoc")
    App.Hide
    If AVDoc.Open("X:\Accounting\Collections\Invoices\247332 INVOICE.pdf", "OpenDoc") Then
        Set PDDoc = AVDoc.GetPDDoc
        thisTime = Timer
        For i = 0 To 100 ''To 1000
            If Not "I NVO I CE" = GetRectStr(PDDoc, AVDoc, 780, 759, 520, 600) Then: MsgBox "Failure to complete rectangle. Iteration " & i
        Next i
        Debug.Print "Time with rectangles: " & CStr(Timer - thisTime)
        thisTime = Timer
        For i = 0 To 100 ''To 1000
            If Not AVDoc.findText("I NVO I CE", 0, 0, 0) Then: MsgBox "Failure to find word. Iteration " & i
        Next i
        Debug.Print "Time with AVDoc.findText: " & CStr(Timer - thisTime)
    End If
    App.CloseAllDocs
End Sub
Private Function GetRectStr(PDDoc, AVDoc, top, bottom, left, right) As String
    Dim textSelect As AcroPDTextSelect, i As Integer, pdText As String
    Set textSelect = PDDoc.CreateTextSelect(0, GetRect(top, bottom, left, right))
    If textSelect Is Nothing Then
        GetRectStr = ""
    Else
        Call AVDoc.SetTextSelection(textSelect)
        For i = 0 To textSelect.GetNumText() - 1
            pdText = pdText & textSelect.GetText(i)
        Next
        GetRectStr = pdText
    End If
End Function
Private Function GetRect(top, bottom, left, right) As AcroRect
    Dim PDRect As AcroRect
    Set PDRect = CreateObject("AcroExch.Rect")
    PDRect.top = top
    PDRect.bottom = bottom
    PDRect.left = left
    PDRect.right = right
    Set GetRect = PDRect
End Function
