Attribute VB_Name = "Globals"
Option Explicit

''Reference Groups
Public g_ARCollection As New Collection

''PDF Data Coordinates
Public g_SOTop As Integer, g_SOBottom As Integer, g_SOLeft As Integer, g_SORight As Integer
Public g_POTop As Integer, g_POBottom As Integer, g_POLeft As Integer, g_PORight As Integer
Public g_DocTypeTop As Integer, g_DocTypeBottom As Integer, g_DocTypeLeft As Integer, g_DocTypeRight As Integer
Public g_BrokerTop As Integer, g_BrokerBottom As Integer, g_BrokerLeft As Integer, g_BrokerRight As Integer
Public g_EmailerTop As Integer, g_EmailerBottom As Integer, g_EmailerLeft As Integer, g_EmailerRight As Integer
Public g_EmailAddressTop As Integer, g_EmailAddressBottom As Integer, g_EmailAddressLeft As Integer, g_EmailAddressRight As Integer
Public g_StreetAddressTop As Integer, g_StreetAddressBottom As Integer, g_StreetAddressLeft As Integer, g_StreetAddressRight As Integer
Public g_MessageTop As Integer, g_MessageBottom As Integer, g_MessageLeft As Integer, g_MessageRight As Integer
Public g_CustomerTop As Integer, g_CustomerBottom As Integer, g_CustomerLeft As Integer, g_CustomerRight As Integer

''Value Groups
Public g_EmailArray() As Variant, g_BrokersAry() As String

''Value Items
Public g_Signature As String, g_DataAccessXML As Boolean, g_IsPrintingDisabled As Boolean, g_IsEmailDisabled As Boolean ''DataAccess INCOMPLETE
Public g_HostName As String, g_ExpectedError As Boolean
Public Function Initialize() As Boolean
    Dim myBool As Boolean
    On Error GoTo EndAll
    Select Case myBool
        Case ARCollection: myBool = True
        Case SO: myBool = True
        Case PO: myBool = True
        Case DocType: myBool = True
        Case Broker: myBool = True
        Case EmailerB: myBool = True
        Case EmailAddress: myBool = True
        Case StreetAddress: myBool = True
        Case Message: myBool = True
        Case Customer: myBool = True
        Case PullBroker: myBool = True
        Case Signature: myBool = True
        Case MiscSettings: myBool = True
    End Select
    Initialize = Not myBool
    Exit Function
EndAll:
    Initialize = False
    Exit Function
End Function
Public Function InitializeEmails() As Boolean
    AutoMailWindow.WhatAmIDoing = "Initializing Email array..."
    InitializeEmails = Emails
End Function
Private Function ARCollection() As Boolean
    Set g_ARCollection = Nothing
    Set g_ARCollection = New Collection
    ARCollection = True
End Function
Private Function SO() As Boolean
    Dim myLabel As String
    myLabel = "SO"
    On Error GoTo EndAll:
    g_SOTop = Pull(myLabel, 2)
    g_SOBottom = Pull(myLabel, 3)
    g_SOLeft = Pull(myLabel, 4)
    g_SORight = Pull(myLabel, 5)
    SO = True
    Exit Function
EndAll:
    SO = False
    Exit Function
End Function
Private Function PO() As Boolean
    Dim myLabel As String
    myLabel = "PO"
    On Error GoTo EndAll:
    g_POTop = Pull(myLabel, 2)
    g_POBottom = Pull(myLabel, 3)
    g_POLeft = Pull(myLabel, 4)
    g_PORight = Pull(myLabel, 5)
    PO = True
    Exit Function
EndAll:
    PO = False
    Exit Function
End Function
Private Function DocType() As Boolean
    Dim myLabel As String
    myLabel = "DocType"
    On Error GoTo EndAll:
    g_DocTypeTop = Pull(myLabel, 2)
    g_DocTypeBottom = Pull(myLabel, 3)
    g_DocTypeLeft = Pull(myLabel, 4)
    g_DocTypeRight = Pull(myLabel, 5)
    DocType = True
    Exit Function
EndAll:
    DocType = False
    Exit Function
End Function
Private Function Broker() As Boolean
    Dim myLabel As String
    myLabel = "Broker"
    On Error GoTo EndAll:
    g_BrokerTop = Pull(myLabel, 2)
    g_BrokerBottom = Pull(myLabel, 3)
    g_BrokerLeft = Pull(myLabel, 4)
    g_BrokerRight = Pull(myLabel, 5)
    Broker = True
    Exit Function
EndAll:
    Broker = False
    Exit Function
End Function
Private Function EmailerB() As Boolean
    Dim myLabel As String
    myLabel = "Emailer"
    On Error GoTo EndAll:
    g_EmailerTop = Pull(myLabel, 2)
    g_EmailerBottom = Pull(myLabel, 3)
    g_EmailerLeft = Pull(myLabel, 4)
    g_EmailerRight = Pull(myLabel, 5)
    EmailerB = True
    Exit Function
EndAll:
    EmailerB = False
    Exit Function
End Function
Private Function EmailAddress() As Boolean
    Dim myLabel As String
    myLabel = "EmailAddress"
    On Error GoTo EndAll:
    g_EmailAddressTop = Pull(myLabel, 2)
    g_EmailAddressBottom = Pull(myLabel, 3)
    g_EmailAddressLeft = Pull(myLabel, 4)
    g_EmailAddressRight = Pull(myLabel, 5)
    EmailAddress = True
    Exit Function
EndAll:
    EmailAddress = False
    Exit Function
End Function
Private Function StreetAddress() As Boolean
    Dim myLabel As String
    myLabel = "StreetAddress"
    On Error GoTo EndAll:
    g_StreetAddressTop = Pull(myLabel, 2)
    g_StreetAddressBottom = Pull(myLabel, 3)
    g_StreetAddressLeft = Pull(myLabel, 4)
    g_StreetAddressRight = Pull(myLabel, 5)
    StreetAddress = True
    Exit Function
EndAll:
    StreetAddress = False
    Exit Function
End Function
Private Function Message() As Boolean
    Dim myLabel As String
    myLabel = "Message"
    On Error GoTo EndAll:
    g_MessageTop = Pull(myLabel, 2)
    g_MessageBottom = Pull(myLabel, 3)
    g_MessageLeft = Pull(myLabel, 4)
    g_MessageRight = Pull(myLabel, 5)
    Message = True
    Exit Function
EndAll:
    Message = False
    Exit Function
End Function
Private Function Customer() As Boolean
    Dim myLabel As String
    myLabel = "Customer"
    On Error GoTo EndAll:
    g_CustomerTop = Pull(myLabel, 2)
    g_CustomerBottom = Pull(myLabel, 3)
    g_CustomerLeft = Pull(myLabel, 4)
    g_CustomerRight = Pull(myLabel, 5)
    Customer = True
    Exit Function
EndAll:
    Customer = False
    Exit Function
End Function
Private Function Pull(thisLabel As String, column As Integer) As String
    Dim myRng As Variant, r As Integer
    myRng = Range("Coordinates")
    For r = 1 To UBound(myRng)
        If myRng(r, 1) = thisLabel Then
            Pull = myRng(r, column)
            Exit Function
        End If
    Next r
    Pull = "##ERROR##"
End Function
Private Function PullBroker() As Boolean
    Dim myRng As Range, r As Integer, c As Integer, maxRows As Integer, maxColumns As Integer
    Set myRng = Range("Brokers")
    maxRows = UBound(myRng.value, 1) - 1
    maxColumns = UBound(myRng.value, 2) - 1
    ReDim g_BrokersAry(maxRows, maxColumns) As String
    For r = 0 To maxRows
        For c = 0 To maxColumns
            g_BrokersAry(r, c) = myRng(r + 1, c + 1)
        Next c
    Next r
    PullBroker = True
End Function
Private Function Emails() As Boolean
    Dim ARDoc As ARDocument, emailCount As Integer
    For Each ARDoc In g_ARCollection
        If ARDoc.IsEmailer Then: emailCount = emailCount + 1
        If ARDoc.HasBroker Then: emailCount = emailCount + 1
    Next ARDoc
    If emailCount > 0 Then: ReDim g_EmailArray(emailCount - 1) As Variant
    Emails = True
End Function
Private Function Signature() As Boolean
    g_Signature = ShortcutsOutlook.Signature
    Signature = True
End Function
Private Function MiscSettings() As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Rules")
    g_IsPrintingDisabled = ws.Cells(1, 19)
    g_IsEmailDisabled = ws.Cells(2, 19)
    MiscSettings = True
End Function
