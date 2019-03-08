Attribute VB_Name = "hFileBrowse"
Option Explicit
Private wb As Workbook, ws As Worksheet, myCell As Object, fileDirStr As String, fileDia As Object
Public Function FileBrowse_Main(recordRow As Integer, recordColumn As Integer) As String
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Rules")
    Set myCell = ws.Cells(recordRow, recordColumn)
    Set fileDia = Application.FileDialog(msoFileDialogFilePicker)
    With fileDia
        .AllowMultiSelect = False
        .title = "Open Directory"
        .Filters.Add "Excel", "*.pdf"
        .Filters.Add "All", "*.*"
        If .Show = True Then
            fileDirStr = .SelectedItems(1)
        End If
    End With
    myCell.value = fileDirStr
    FileBrowse_Main = fileDirStr
End Function
