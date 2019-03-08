VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutoMailWindow 
   Caption         =   "AutoMail"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12450
   OleObjectBlob   =   "AutoMailWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AutoMailWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_ws As Worksheet
Private Sub AddButton_Click()
    AddRuleWindow.Show
End Sub
Private Sub BrowseButton_Click()
    AutoMailWindow.FileTextBox = hFileBrowse.FileBrowse_Main(1, 1)
End Sub
Private Sub ButtonAdmin_Click()
    AdminWindow.Show
End Sub

Private Sub ButtonFeedback_Click()
    
End Sub

Private Sub DownButton_Click()
    Dim index As Integer
    With AutoMailWindow.RuleListBox
        index = .ListIndex
        If .ListIndex < UBound(.List) Then
            .List = Swap2DArray(.List, index, .ListIndex + 1, .columncount)
            .ListIndex = index + 1
        End If
    End With
End Sub
Private Sub EditButton_Click()
    EditRuleWindow.Show
End Sub

Private Sub Help_Click()

End Sub

Private Sub RemoveButton_Click()
    Dim index As Integer
    With AutoMailWindow.RuleListBox
        index = .ListIndex
        .List = ShortCutsArrays.RemoveFrom2DArray(.List, index, .columncount)
        Range("RuleList").value = .List
        Names("RuleList").refersTo = ShortcutsRanges.LessRow(Names("RuleList").refersTo)
        If .ListIndex = UBound(.List) Then
            If .ListIndex > 1 Then
                .ListIndex = index - 1
            Else
                .ListIndex = index
            End If
        Else
            .ListIndex = UBound(.List)
        End If
    End With
End Sub
Private Sub RuleListBox_Change()
    With AutoMailWindow.RuleListBox
        AutoMailWindow.NotesLabel = .List(.ListIndex, 4)
    End With
End Sub
Private Sub RuleListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    EditRuleWindow.Show
End Sub

Private Sub RunButton_Click()
    With AutoMailWindow
        Call hOOPAutoMail.Main
    End With
End Sub

Private Sub UpButton_Click()
    Dim index As Integer
    With AutoMailWindow.RuleListBox
        index = .ListIndex
        If .ListIndex > 0 Then
            .List = Swap2DArray(.List, index, .ListIndex - 1, .columncount)
            .ListIndex = index - 1
        End If
    End With
End Sub
Private Function Swap2DArray(thisArray As Variant, indexA As Integer, indexB As Integer, columncount As Integer) As Variant
    Dim r As Integer, c As Integer, tempArray As Variant
    tempArray = Array()
    ReDim tempArray(1, columncount) As Variant
    For c = 0 To 4
        tempArray(0, c) = thisArray(indexA, c)
        tempArray(1, c) = thisArray(indexB, c)
    Next c
    For c = 0 To 4
        thisArray(indexA, c) = tempArray(1, c)
        thisArray(indexB, c) = tempArray(0, c)
    Next c
    Swap2DArray = thisArray
End Function
Private Sub UserForm_Initialize()
    'Stop
    'hMenuBar.ApplyMenuBar
    g_HostName = VBA.Environ$("computername")
    With Me
        .RuleListBox.List = GetRuleList
        .FileTextBox = GetFileDirectory
        Select Case g_HostName
            Case "KERRTERM2": Call MockButtons
            Case "KERRTERM3": Call MockButtons
            Case Else
                MsgBox "Server not supported. Automail closing..."
                Me.Hide
                Exit Sub
        End Select
        '.ButtonTest = False
    End With
    If Not Workbooks.Count > 1 Then
        Application.Visible = False
    End If
End Sub
Private Sub MockButtons()
    AutoMailWindow.UpButton.Caption = VBA.ChrW(9650)
    AutoMailWindow.DownButton.Caption = VBA.ChrW(9660)
End Sub
Private Function GetFileDirectory() As String
    GetFileDirectory = ThisWorkbook.Worksheets("Rules").Cells(1, 1)
End Function
Private Function GetWorksheet(wsName As String) As Worksheet
    Set GetWorksheet = ThisWorkbook.Worksheets(wsName)
End Function
Private Function GetRuleList() As Variant
    Dim thisArray As Variant, r As Integer, c As Integer, myRng() As Variant
    thisArray = Array()
    myRng = Range("RuleList")
    ReDim thisArray(UBound(myRng) - 1, 4) As Variant
    For r = 0 To UBound(myRng) - 1
        For c = 0 To 4
            thisArray(r, c) = myRng(r + 1, c + 1)
        Next c
    Next r
    GetRuleList = thisArray
End Function
Private Sub UserForm_Terminate()
    With AutoMailWindow.RuleListBox
        Range("RuleList") = .List
    End With
    ThisWorkbook.Save
    ThisWorkbook.Saved = True
    If Workbooks.Count > 1 Then
      ThisWorkbook.Close
    Else
        Application.Quit
    End If
End Sub
