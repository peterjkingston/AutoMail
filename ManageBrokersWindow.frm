VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ManageBrokersWindow 
   Caption         =   "Manage Brokers"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   OleObjectBlob   =   "ManageBrokersWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ManageBrokersWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonAdd_Click()
    With BrokersAddOrEditWindow
        .Caption = "Add Broker"
        .Func = "0"
        .TextBoxName = ""
        .TextBoxEmails = ""
        .Show
    End With
End Sub

Private Sub ButtonEdit_Click()
    Dim myLB As Object
    Set myLB = ManageBrokersWindow.ListBox1
    With BrokersAddOrEditWindow
        .Caption = "Edit Broker"
        .Func = "1"
        .TextBoxName = ManageBrokersWindow.ListBox1.List(myLB.ListIndex, 0)
        .TextBoxEmails = ManageBrokersWindow.ListBox1.List(myLB.ListIndex, 1)
        .Show
    End With
End Sub

Private Sub ButtonRemove_Click()
    Dim myLB As Object, myRng As Range
    Set myLB = ManageBrokersWindow.ListBox1
    Set myRng = Range("Brokers")
    myLB.List = ShortCutsArrays.RemoveFrom2DArray(myLB.List, myLB.ListIndex, 2)
    Names("Brokers").refersTo = ShortcutsRanges.LessRow(Names("Brokers").refersTo)
    myRng.value = myLB.List
End Sub

Private Sub ButtonSort_Click()
    Dim myLB As Object, myRng As Range
    Set myLB = ManageBrokersWindow.ListBox1
    Set myRng = Range("Brokers")
    MsgBox "Function disabled.", vbOKOnly, "AutoMail"
    'If UBound(myLB.List) >= 1 Then
    '    myLB.List = ShortCutsArrays.ShellSort2D(myLB.List)
    '    myRng.value = myLB.List
    'End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim myLB As Object
    Set myLB = ManageBrokersWindow.ListBox1
    With BrokersAddOrEditWindow
        .Caption = "Edit Broker"
        .Func = "1"
        .TextBoxName = ManageBrokersWindow.ListBox1.List(myLB.ListIndex, 0)
        .TextBoxEmails = ManageBrokersWindow.ListBox1.List(myLB.ListIndex, 1)
        .Show
    End With
End Sub

Private Sub UserForm_Initialize()
    ListBox1.List = Range("Brokers").value
End Sub
