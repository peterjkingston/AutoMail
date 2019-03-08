VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BrokersAddOrEditWindow 
   Caption         =   "Add Broker"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   OleObjectBlob   =   "BrokersAddOrEditWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BrokersAddOrEditWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Dim Func As String
    Func = BrokersAddOrEditWindow.Func
    Select Case Func
        Case "0": AddIntoList
        Case "1": AppendList
    End Select
    BrokersAddOrEditWindow.Hide
End Sub
Private Sub AddIntoList()
    Dim myLB As Object, myRng As Range, newRecord(0, 1) As Variant
    Set myLB = ManageBrokersWindow.ListBox1
    Set myRng = Range("Brokers")
    With BrokersAddOrEditWindow
        Stop
        newRecord(0, 0) = UCase(Trim(.TextBoxName))
        newRecord(0, 1) = Trim(.TextBoxEmails)
        myLB.List = ShortCutsArrays.Append2DArray(myLB.List, newRecord, 1)
        'ShortCutsArrays.ShellSort (myLB.List)
        ThisWorkbook.Names("Brokers").refersTo = ShortcutsRanges.NextRow(ThisWorkbook.Names("Brokers").refersTo)
        Range("Brokers").value = myLB.List
    End With
End Sub
Private Sub AppendList()
    Dim myLB As Object, myRng As Range, newRecord(0, 1) As Variant
    Set myLB = ManageBrokersWindow.ListBox1
    Set myRng = Range("Brokers")
    With BrokersAddOrEditWindow
        newRecord(0, 0) = UCase(Trim(.TextBoxName))
        newRecord(0, 1) = Trim(.TextBoxEmails)
        myLB.List = ShortCutsArrays.Replace2DArray(myLB.List, newRecord, myLB.ListIndex, 1)
        'ShortCutsArrays.ShellSort (myLB.List)
        Range("Brokers").value = myLB.List
    End With
End Sub

Private Sub UserForm_Click()

End Sub
