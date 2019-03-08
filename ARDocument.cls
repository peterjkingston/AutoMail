VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ARDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private myDoc As New ARDocument
''TODO: Proof this.
Private m_SO As String, m_PO As String, m_CustomerID As String, m_PageCount As Integer, m_IsWebInvoice As Boolean
Private m_PDDoc As AcroPDDoc, m_IsInvoice As Boolean, m_IsCreditMemo As Boolean, m_DocType As String, m_Broker As String
Private m_EmailAddress As String, m_IsEmailer As Boolean, m_IsPrinter As Boolean, m_HasBroker As Boolean, m_AVDoc As AcroAVDoc
Private m_StreetAddress As String, m_IsOpen As Boolean, m_Directory As String
Private m_CC As String
Event BeforeUpdate()
Event AfterUpdate()
Public Sub OpenPDF()
    Dim myBool As Boolean
    myBool = m_AVDoc.Open(m_Directory, m_SO)
    If m_AVDoc.IsValid = False Then: Error (123)
End Sub
Public Property Set PDDoc(PDDoc As AcroPDDoc)
    RaiseEvent BeforeUpdate
    Set m_PDDoc = PDDoc
    Call AssignData
    Call InferData
    RaiseEvent AfterUpdate
End Property
Public Property Set AVDoc(AVDoc As AcroAVDoc)
    RaiseEvent BeforeUpdate
    Set m_AVDoc = AVDoc
    RaiseEvent AfterUpdate
End Property
Private Sub AssignData()
    Set m_AVDoc = m_PDDoc.OpenAVDoc(m_SO)
    m_DocType = ShortcutsAcrobat.GetRectStr(m_PDDoc, m_AVDoc, g_DocTypeTop, g_DocTypeBottom, g_DocTypeLeft, g_DocTypeRight)
    m_PO = ShortcutsAcrobat.GetRectStr(m_PDDoc, m_AVDoc, g_POTop, g_POBottom, g_POLeft, g_PORight)
    m_CustomerID = ShortcutsAcrobat.GetRectStr(m_PDDoc, m_AVDoc, g_CustomerTop, g_CustomerBottom, g_CustomerLeft, g_CustomerRight)
    m_Broker = ShortcutsAcrobat.GetRectStr(m_PDDoc, m_AVDoc, g_BrokerTop, g_BrokerBottom, g_BrokerLeft, g_BrokerRight)
    m_StreetAddress = ShortcutsAcrobat.GetRectStr(m_PDDoc, m_AVDoc, g_StreetAddressTop, g_StreetAddressBottom, g_StreetAddressLeft, g_StreetAddressRight)
    m_EmailAddress = ShortcutsAcrobat.GetRectStr(m_PDDoc, m_AVDoc, g_EmailAddressTop, g_EmailAddressBottom, g_EmailAddressLeft, g_EmailAddressRight)
    If ShortcutsAcrobat.GetRectStr(m_PDDoc, m_AVDoc, g_EmailerTop, g_EmailerBottom, g_EmailerLeft, g_EmailerRight) = "Email Processed" Then
        m_IsEmailer = True
    Else
        m_IsEmailer = False
    End If
    m_IsOpen = True
End Sub
Public Sub PrintSilent(App As AcroApp)
    Dim myOperator As Boolean, i As Integer
    For i = 0 To App.GetNumAVDocs - 1
        Set m_AVDoc = App.GetAVDoc(i)
        If InStr(1, m_AVDoc.GetTitle, m_SO) > 0 Then
            Exit For
        End If
    Next i
    Set m_PDDoc = m_AVDoc.GetPDDoc
    If myOperator = m_AVDoc.PrintPagesSilent(0, m_PDDoc.GetNumPages - 1, 2, True, True) Then
        MsgBox "Failure to print SO " & m_SO
    End If
End Sub
Private Sub InferData()
    If m_DocType = "I NVO I CE" Then
        m_DocType = "INVOICE"
        m_IsInvoice = True
    End If
    If Trim(m_DocType) = "CRED I T MEMO" Then
        m_DocType = "CREDIT MEMO"
        m_IsCreditMemo = True
    End If
    If m_Broker <> "" Then: m_HasBroker = True
    If m_HasBroker Then: CorrectBrokerToEmail
    If GetWebInvoice Then: m_IsWebInvoice = True
    If Not m_IsWebInvoice And Not m_IsEmailer Then
        m_IsPrinter = True
    Else
        m_IsPrinter = False
    End If
End Sub
Private Sub CorrectBrokerToEmail()
    Dim r As Integer, c As Integer
    For r = 0 To UBound(g_BrokersAry)
        If InStr(1, UCase(m_Broker), UCase(g_BrokersAry(r, 0))) > 0 Then
            m_Broker = g_BrokersAry(r, 1)
            Exit Sub
        End If
    Next r
End Sub
Public Sub SavePDF()
    m_PDDoc.Save PDSaveFull, m_Directory
End Sub
Private Function GetWebInvoice() As Boolean
    Dim myBool As Boolean
    Select Case Mid(m_EmailAddress, 1, 4)
        Case "EDI,": myBool = True
        Case "EDI-": myBool = True
        Case "EDI ": myBool = True
        Case "OB10": myBool = True
    End Select
    GetWebInvoice = myBool
End Function
Public Sub ClosePDF()
    m_AVDoc.Close True
    m_PDDoc.Close
End Sub
Public Property Get IsOpen() As Boolean
    IsOpen = m_IsOpen
End Property
Public Property Get StreetAddress() As String
    StreetAddress = m_StreetAddress
End Property
Public Property Get AVDoc() As AcroAVDoc
    AVDoc = m_AVDoc
End Property
Public Property Get SO() As String
    SO = m_SO
End Property
Public Property Get PO() As String
    PO = m_PO
End Property
Public Property Get CustomerID() As String
    CustomerID = m_CustomerID
End Property
Public Property Get pageCount() As Integer
    pageCount = m_PDDoc.GetNumPages
End Property
Public Property Get IsWebInvoice() As Boolean
    IsWebInvoice = m_IsWebInvoice
End Property
Public Property Get PDDoc() As AcroPDDoc
    Set PDDoc = m_PDDoc
End Property
Public Property Get Broker() As String
    Broker = m_Broker
End Property
Public Property Get HasBroker() As Boolean
    HasBroker = m_HasBroker
End Property
Public Property Get DocumentType() As String
    DocumentType = m_DocType
End Property
Public Property Get IsCreditMemo() As Boolean
    IsCreditMemo = m_IsCreditMemo
End Property
Public Property Get IsInvoice() As Boolean
    IsInvoice = m_IsInvoice
End Property
Public Property Get EmailAddress() As String
    EmailAddress = m_EmailAddress
End Property
Public Property Get CC() As String
    CC = m_CC
End Property
Public Property Get IsPrinter() As Boolean
    IsPrinter = m_IsPrinter
End Property
Public Property Get IsEmailer() As Boolean
    IsEmailer = m_IsEmailer
End Property
Public Property Get FileDir() As String
    FileDir = m_Directory
End Property
Public Property Let CC(myCC As String)
    RaiseEvent BeforeUpdate
    m_CC = myCC
    RaiseEvent AfterUpdate
End Property
Public Property Let IsPrinter(boolDoPrint As Boolean)
    RaiseEvent BeforeUpdate
    m_IsPrinter = boolDoPrint
    RaiseEvent AfterUpdate
End Property
Public Property Let IsEmailer(boolDoEmail As Boolean)
    RaiseEvent BeforeUpdate
    m_IsEmailer = boolDoEmail
    RaiseEvent AfterUpdate
End Property
Public Property Let SO(SO As String)
    m_SO = SO
End Property
Public Property Let IsOpen(IsOpen As Boolean)
    m_IsOpen = IsOpen
End Property
Public Property Let EmailAddress(Address As String)
    m_EmailAddress = Address
End Property
Public Property Let HasBroker(val As Boolean)
    m_HasBroker = val
End Property
Public Property Let FileDir(FileDir As String)
    m_Directory = FileDir
End Property

