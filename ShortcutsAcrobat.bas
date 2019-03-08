Attribute VB_Name = "ShortcutsAcrobat"
Option Explicit
''TODO Proof this.
Public Function GetRectStr(PDDoc As AcroPDDoc, AVDoc As AcroAVDoc, top As Integer, bottom As Integer, left As Integer, right As Integer)
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
Public Function GetRect(top As Integer, bottom As Integer, left As Integer, right As Integer) As AcroRect
    Dim PDRect As AcroRect
    Set PDRect = CreateObject("AcroExch.Rect")
    PDRect.top = top
    PDRect.bottom = bottom
    PDRect.left = left
    PDRect.right = right
    Set GetRect = PDRect
End Function
Public Function GetAVDoc(ARDoc As ARDocument, App As AcroApp) As AcroAVDoc
    Dim i As Integer
    For i = 0 To App.GetNumAVDocs - 1
        Set GetAVDoc = App.GetAVDoc(i)
        If InStr(1, GetAVDoc.GetTitle, ARDoc.SO) > 0 Then Exit For
    Next i
End Function
