VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PDFCoordinatesAddWindow 
   Caption         =   "Add Coordinates"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   OleObjectBlob   =   "PDFCoordinatesAddWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PDFCoordinatesAddWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BrowseButton_Click()
    PDFCoordinatesAddWindow.FileTextBox = hFileBrowse.FileBrowse_Main(1, 7)
End Sub
Private Sub ButtonCalibrate_Click()
    With PDFCoordinatesAddWindow
        If IsFilledOut Then
            Call Calibration.FindString(.FileTextBox, CInt(Trim(.TextBoxTop)), CInt(Trim(.TextBoxBottom)), _
                 CInt(Trim(.TextBoxLeft)), CInt(Trim(.TextBoxRight)))
        End If
    End With
End Sub
Private Sub ButtonSubmit_Click()
    Dim newArray(0, 5) As String, myRange As Range
    If IsFilledOut And IsValid Then
        With PDFCoordinatesAddWindow
            newArray(0, 0) = .TextBoxDataType
            newArray(0, 1) = .TextBoxTop
            newArray(0, 2) = .TextBoxBottom
            newArray(0, 3) = .TextBoxLeft
            newArray(0, 4) = .TextBoxRight
            newArray(0, 5) = False
        End With
        With PDFCoordinatesWindow.ListBoxCoordinates
            .List = ShortCutsArrays.Append2DArray(.List, newArray, .columncount)
            Names("Coordinates").refersTo = ShortcutsRanges.NextRow(Names("Coordinates").refersTo)
            Range("Coordinates").value = .List
            .ListIndex = UBound(.List)
        End With
        PDFCoordinatesAddWindow.Hide
    Else
        MsgBox "Missing data, please fill out all fields.", vbOKOnly, "AutoMail"
    End If
End Sub
Private Sub UserForm_Initialize()
    With PDFCoordinatesAddWindow
        .FileTextBox = ThisWorkbook.Worksheets("Rules").Cells(1, 7)
    End With
End Sub
Private Function IsFilledOut() As Boolean
    Dim myBool As Boolean
    With PDFCoordinatesAddWindow
        If IsValid Then
            Select Case myBool
                Case Trim(.TextBoxTop) <> "": myBool = True
                Case CBool(Not IsError(CInt(Trim(.TextBoxTop)))): myBool = True
                Case Trim(.TextBoxBottom) <> "": myBool = True
                Case CBool(Not IsError(CInt(Trim(.TextBoxBottom)))): myBool = True
                Case Trim(.TextBoxLeft) <> "": myBool = True
                Case CBool(Not IsError(CInt(Trim(.TextBoxLeft)))): myBool = True
                Case Trim(.TextBoxRight) <> "": myBool = True
                Case CBool(Not IsError(CInt(Trim(.TextBoxRight)))): myBool = True
            End Select
        End If
    End With
    IsFilledOut = Not myBool
End Function
Private Function IsValid() As Boolean
    Dim myBool As Boolean
    With PDFCoordinatesAddWindow
        Select Case myBool
            Case CBool(CInt(.TextBoxTop) > CInt(.TextBoxBottom)): myBool = True
            Case CBool(CInt(.TextBoxRight) > CInt(.TextBoxLeft)): myBool = True
        End Select
    End With
    IsValid = Not myBool
End Function
