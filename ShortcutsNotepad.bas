Attribute VB_Name = "ShortcutsNotepad"
Option Explicit
''TODO: Proof this.
Public Function LogAll() As Boolean
    Dim fileNum As Integer, myFile As String, myName As String, ARDoc As ARDocument, myStr As String, objFSO As FileSystemObject
    Dim objFolder As Folder
    AutoMailWindow.WhatAmIDoing = "Logging records..."
    fileNum = FreeFile
    myFile = "Automailer Log"
    myName = Application.ThisWorkbook.Path & "\"
    Open "X:\Parity\" & myFile & ".TXT" For Append As #fileNum
    For Each ARDoc In g_ARCollection
        myStr = myStr & "||" & ARDoc.SO
        If ARDoc.IsPrinter Then: myStr = myStr & " Printed"
        If ARDoc.IsEmailer Then: myStr = myStr & " Emailed to " & ARDoc.EmailAddress
        If ARDoc.HasBroker Then: myStr = myStr & " Emailed to " & ARDoc.Broker
    Next ARDoc
    myStr = Date & myStr
    Print #fileNum, myStr
    Close #fileNum
    LogAll = True
End Function
Public Function Append(ARDoc As ARDocument) As Boolean
    ''TODO
End Function
