Attribute VB_Name = "ShortcutsOutlook"
Option Explicit
Private m_BOLPresent As Boolean, m_BOLDir As String
''TODO: Proof this.
Public Function Emailer(ARDoc As ARDocument) As MailItem
    Dim OutMail As MailItem, OutApp As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    With OutMail
        .To = ARDoc.EmailAddress
        .CC = ""
        .BCC = ""
        .Subject = "PO# " & ARDoc.PO
        .HTMLBody = "Attached is the " & LCase(ARDoc.DocumentType) & " for the above PO#." & vbCrLf & g_Signature
        .Attachments.Add ARDoc.FileDir
    End With
    Set Emailer = OutMail
    Exit Function
EndAll:
    MsgBox "Error occurred while emailing out " & ARDoc.SO & " to " & ARDoc.EmailAddress & ". AutoMail process ended."
    Exit Function
End Function
Public Function Signature() As String
    Dim sigString As String
    sigString = Environ("appdata") & "\Microsoft\Signatures\AutoMail.htm"
    If Dir(sigString) <> "" Then
        Signature = GetBoiler(sigString)
    Else
        Signature = ""
    End If
End Function
Function GetBoiler(ByVal sFile As String) As String
    Dim FSO As Object, ts As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set ts = FSO.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.readall
    ts.Close
End Function
Private Function GetFileDir(fileName As String) As String
    GetFileDir = ThisWorkbook.Path & "\Input Directory\" & fileName & ".pdf"
End Function
Public Function EmailerBroker(ARDoc As ARDocument) As MailItem
    Dim OutMail As MailItem, OutApp As Object, BOLDir As String, BOLPresent As Boolean
    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    m_BOLPresent = False
    m_BOLDir = ""
    With OutMail
        .To = ARDoc.Broker
        .CC = ""
        .BCC = ""
        .Subject = "PO# " & ARDoc.PO
        .HTMLBody = "Attached is the " & LCase(ARDoc.DocumentType) & " for the above PO#." & vbCrLf & g_Signature
        .Attachments.Add ARDoc.FileDir
        Call DefineBOLStr(ARDoc)
        If m_BOLPresent Then
            .Attachments.Add m_BOLDir
        End If
    End With
    Set EmailerBroker = OutMail
    Exit Function
'EndAll:
    MsgBox "Error occurred while emailing out " & ARDoc.DocumentType & " " & ARDoc.SO & " to " & ARDoc.Broker & ". AutoMail process ended."
    Exit Function
End Function
Private Sub DefineBOLStr(ARDoc As ARDocument)
    Dim FSO As New FileSystemObject, myFolder As Folder
    Set myFolder = FSO.GetFolder(ThisWorkbook.Path & "\BOLs")
    If FSO.FileExists(ThisWorkbook.Path & "\BOLs\" & ARDoc.SO & " BOL.pdf") Then
        m_BOLDir = ThisWorkbook.Path & "\BOLs\" & ARDoc.SO & " BOL.pdf"
        m_BOLPresent = True
        Exit Sub
    ElseIf FSO.FileExists(ThisWorkbook.Path & "\BOLs\" & ARDoc.PO & " BOL.pdf") Then
        m_BOLDir = ThisWorkbook.Path & "\BOLs\" & ARDoc.PO & " BOL.pdf"
        m_BOLPresent = True
        Exit Sub
    End If
    m_BOLPresent = False
End Sub
