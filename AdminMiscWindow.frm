VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdminMiscWindow 
   Caption         =   "Miscellaneous Settings"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2790
   OleObjectBlob   =   "AdminMiscWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AdminMiscWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBoxDisableEmail_Change()
    Dim ws As Worksheet, wsIsEmailDisabled As Boolean
    Set ws = ThisWorkbook.Worksheets("Rules")
    wsIsEmailDisabled = AdminMiscWindow.CheckBoxDisableEmail.value
    ws.Cells(2, 19) = wsIsEmailDisabled
End Sub

Private Sub CheckBoxDisablePrinting_Change()
    Dim ws As Worksheet, wsIsPrintingDisabled As Boolean
    Set ws = ThisWorkbook.Worksheets("Rules")
    wsIsPrintingDisabled = AdminMiscWindow.CheckBoxDisablePrinting.value
    ws.Cells(1, 19) = wsIsPrintingDisabled
End Sub

Private Sub UserForm_Activate()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Rules")
    CheckBoxDisablePrinting = ws.Cells(1, 19)
    CheckBoxDisableEmail = ws.Cells(2, 19)
End Sub

