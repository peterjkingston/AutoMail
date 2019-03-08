VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdminWindow 
   Caption         =   "Administrator"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2790
   OleObjectBlob   =   "AdminWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AdminWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonDataAccess_Click()
    DataAccessWindow.Show
End Sub
Private Sub ButtonLogoTest_Click()
    LogoWindow.Show
End Sub
Private Sub ButtonManageBrokers_Click()
    ManageBrokersWindow.Show
End Sub

Private Sub ButtonMiscSettings_Click()
    AdminMiscWindow.Show
End Sub

Private Sub ButtonPDFCoord_Click()
    PDFCoordinatesWindow.Show
End Sub
