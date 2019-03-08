VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PDFCoordinatesWindow 
   Caption         =   "PDF Coordinates"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "PDFCoordinatesWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PDFCoordinatesWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ButtonAccept_Click()
    PDFCoordinatesWindow.Hide
End Sub
Private Sub ButtonAdd_Click()
    With PDFCoordinatesAddWindow
        .TextBoxDataType = ""
        .TextBoxTop = ""
        .TextBoxBottom = ""
        .TextBoxLeft = ""
        .TextBoxRight = ""
        .Show
    End With
End Sub
Private Sub ButtonEdit_Click()
    With PDFCoordinatesEditWindow
        .TextBoxDataType = GetFromMyList(PDFCoordinatesWindow.ListBoxCoordinates.ListIndex, 0)
        .TextBoxTop = GetFromMyList(PDFCoordinatesWindow.ListBoxCoordinates.ListIndex, 1)
        .TextBoxBottom = GetFromMyList(PDFCoordinatesWindow.ListBoxCoordinates.ListIndex, 2)
        .TextBoxLeft = GetFromMyList(PDFCoordinatesWindow.ListBoxCoordinates.ListIndex, 3)
        .TextBoxRight = GetFromMyList(PDFCoordinatesWindow.ListBoxCoordinates.ListIndex, 4)
        .Show
    End With
End Sub
Private Function GetFromMyList(index As Integer, column As Integer) As String
    GetFromMyList = PDFCoordinatesWindow.ListBoxCoordinates.List(index, column)
End Function
Private Sub ButtonRemove_Click()
    Dim index As Integer
    With PDFCoordinatesWindow.ListBoxCoordinates
        index = .ListIndex
        .List = RemoveFrom2DArray(.List, index, .columncount)
        Range("Coordinates").value = .List
        Names("Coordinates").refersTo = ShortcutsRanges.LessRow(Names("Coordinates").refersTo)
        If .ListIndex = UBound(.List) Then
            .ListIndex = index - 1
        Else
            .ListIndex = UBound(.List)
        End If
    End With
End Sub
Private Sub UserForm_Initialize()
    With PDFCoordinatesWindow
        .ListBoxCoordinates.List = GetCoordinatesList
    End With
End Sub
Private Function GetCoordinatesList() As Variant
    Dim thisArray() As Variant, myRng As Variant, r As Integer, c As Integer
    myRng = Range("Coordinates")
    With PDFCoordinatesWindow.ListBoxCoordinates
        ReDim thisArray(UBound(myRng) - 1, .columncount + 1) As Variant
        For r = 0 To UBound(myRng)
            For c = 0 To .columncount - 1
                thisArray(r, c) = myRng(r + 1, c + 1)
            Next c
            If r + 1 = UBound(myRng) Then Exit For
        Next r
    End With
    GetCoordinatesList = thisArray
End Function
