Attribute VB_Name = "Calibration"
Option Explicit
Private m_App As AcroApp, m_AVDoc As AcroAVDoc, m_PDDoc As AcroPDDoc
Public Sub FindString(FileDir As String, top As Integer, bottom As Integer, left As Integer, right As Integer)
    Dim myStr As String, ms As Integer, textSelect As AcroPDTextSelect, myRect As AcroRect
    Set m_App = CreateObject("AcroExch.App")
    Set m_AVDoc = CreateObject("AcroExch.AVDoc")
    If m_AVDoc.Open(FileDir, "Finding SO") Then
        m_App.Hide
        Do Until ms = vbYes
            myStr = ShortcutsAcrobat.GetRectStr(m_AVDoc.GetPDDoc, m_AVDoc, top, bottom, left, right)
            ms = MsgBox(myStr & vbCrLf & "Is this correct?", vbYesNo)
            If ms = vbNo Then
                Exit Do
            ElseIf ms = vbYes Then
                Set m_PDDoc = m_AVDoc.GetPDDoc
                Set textSelect = m_PDDoc.CreateTextSelect(0, ShortcutsAcrobat.GetRect(top, bottom, left, right))
                Set myRect = textSelect.GetBoundingRect
                MsgBox DisplayRectCoords(myRect)
                With PDFCoordinatesAddWindow
                    .TextBoxTop = CStr(myRect.top)
                    .TextBoxBottom = CStr(myRect.bottom)
                    .TextBoxLeft = CStr(myRect.left)
                    .TextBoxRight = CStr(myRect.right)
                End With
            End If
        Loop
    End If
    m_App.CloseAllDocs
End Sub
Private Function DisplayRectCoords(myRect As AcroRect) As String
    DisplayRectCoords = "Top: " & CStr(myRect.top) & vbCrLf & _
                       "Bottom: " & CStr(myRect.bottom) & vbCrLf & _
                       "Left: " & CStr(myRect.left) & vbCrLf & _
                       "Right: " & CStr(myRect.right)
End Function
