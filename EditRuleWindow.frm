VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditRuleWindow 
   Caption         =   "Edit Rule"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   OleObjectBlob   =   "EditRuleWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EditRuleWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SubmitButton_Click()
    Dim newArray(0, 4) As String
    If IsFilledOut Then
        With EditRuleWindow
            newArray(0, 0) = .DataTypeComboBox
            newArray(0, 1) = .ConditionTextBox
            newArray(0, 2) = .ActionComboBox
            newArray(0, 3) = .AccessorTextBox
            newArray(0, 4) = .NotesTextBox
            .Hide
            Call UserForm_Initialize
        End With
        With AutoMailWindow.RuleListBox
            .List = ShortCutsArrays.Replace2DArray(.List, newArray, .ListIndex, .columncount)
            Range("RuleList").value = .List
        End With
    Else
        MsgBox "Missing data, please fill out all fields.", vbOKOnly, "AutoMail"
    End If
End Sub
Private Sub UserForm_Initialize()
    Dim listArray() As Variant, index As Integer
    listArray = AutoMailWindow.RuleListBox.List
    index = AutoMailWindow.RuleListBox.ListIndex
    With EditRuleWindow
        .Width = 408
        .DataTypeComboBox.List = Array("<Data Type>", "Document Type", "SO#", "PO#", "Customer ID", "Broker", "EmailAddress", "StreetAddress", "Find Text")
        .ActionComboBox.List = Array("<Action>", "Do not Email", "Do not Print", "Email", "CC", "Print", "Notify me", "Inspect it", "Do Nothing")
        .DataTypeComboBox = listArray(index, 0)
        .ConditionTextBox = listArray(index, 1)
        .ActionComboBox = listArray(index, 2)
        .AccessorTextBox = listArray(index, 3)
        .NotesTextBox = listArray(index, 4)
    End With
    Call ActionComboBox_Change
End Sub
Private Sub ActionComboBox_Change()
    With EditRuleWindow
        Select Case .ActionComboBox.value
            Case "Email"
                .Width = 594
                .AccessorWordLabel.Visible = True
                .AccessorWordLabel.Caption = "to:"
                .AccessorTextBox.Visible = True
                .SubmitButton.left = 498
                .NotesTextBox.Width = 450
            Case "CC"
                .Width = 594
                .AccessorWordLabel.Visible = True
                .AccessorWordLabel.Caption = "to:"
                .AccessorTextBox.Visible = True
                .SubmitButton.left = 498
                .NotesTextBox.Width = 450
            Case Else
                .Width = 408
                .AccessorWordLabel.Visible = False
                .AccessorTextBox.Visible = False
                .SubmitButton.left = 312
                .NotesTextBox.Width = 264
        End Select
    End With
End Sub
Private Function IsFilledOut() As Boolean
    Dim myBool As Boolean
    With EditRuleWindow
        myBool = True
        Select Case myBool
            Case (Mid(.DataTypeComboBox, 1, 1) = "<"): myBool = False
            Case (Mid(.ConditionTextBox, 1, 1) = "<"): myBool = False
            Case (Trim(Mid(.ConditionTextBox, 1, 1)) = ""): myBool = False
            Case (Mid(.ActionComboBox, 1, 1) = "<"): myBool = False
        End Select
        If .ActionComboBox = "Email To:" And Mid(.AccessorTextBox, 1, 1) = "<" Then
            myBool = False
        End If
        IsFilledOut = myBool
    End With
End Function
