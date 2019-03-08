VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddRuleWindow 
   Caption         =   "Add Rule"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   OleObjectBlob   =   "AddRuleWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddRuleWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActionComboBox_Change()
    With AddRuleWindow
        Select Case .ActionComboBox.value
            Case "Email"
                .Width = 594
                .AccessorWordLabel.Visible = True
                .AccessorWordLabel.Caption = "to:"
                .AccessorTextBox.Visible = True
                .AddRuleButton.left = 498
                .NotesTextBox.Width = 450
            Case "CC"
                .Width = 594
                .AccessorWordLabel.Visible = True
                .AccessorWordLabel.Caption = "to:"
                .AccessorTextBox.Visible = True
                .AddRuleButton.left = 498
                .NotesTextBox.Width = 450
            Case Else
                .Width = 408
                .AccessorWordLabel.Visible = False
                .AccessorTextBox.Visible = False
                .AddRuleButton.left = 312
                .NotesTextBox.Width = 264
        End Select
    End With
End Sub
Private Sub AddRuleButton_Click()
    Dim newArray(0, 4) As String, myRange As Range
    If IsFilledOut Then
        With AddRuleWindow
            newArray(0, 0) = .DataTypeComboBox
            newArray(0, 1) = .ConditionTextBox
            newArray(0, 2) = .ActionComboBox
            If left(.AccessorTextBox, 1) = "<" Then: .AccessorTextBox = "N/A"
            newArray(0, 3) = .AccessorTextBox
            If Trim(.NotesTextBox) = "<Type notes here>" Or Trim(.NotesTextBox = "") Then: .NotesTextBox = "No notes."
            newArray(0, 4) = .NotesTextBox
        End With
        With AutoMailWindow.RuleListBox
            .List = ShortCutsArrays.Append2DArray(.List, newArray, .columncount)
            Names("RuleList").refersTo = ShortcutsRanges.NextRow(Names("RuleList").refersTo)
            Range("RuleList").value = .List
            .ListIndex = UBound(.List)
        End With
        AddRuleWindow.Hide
    Else
        MsgBox "Missing data, please fill out all fields.", vbOKOnly, "AutoMail"
    End If
End Sub
Private Function IsFilledOut() As Boolean
    Dim myBool As Boolean
    With AddRuleWindow
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
Private Sub UserForm_Initialize()
    With AddRuleWindow
        .Width = 408
        .DataTypeComboBox.List = Array("<Data Type>", "Document Type", "SO#", "PO#", "Customer ID", "Broker", "EmailAddress", "StreetAddress", "Find Text")
        .ActionComboBox.List = Array("<Action>", "Do not Email", "Do not Print", "Email", "CC", "Print", "Notify me", "Inspect it", "Do Nothing")
    End With
End Sub
