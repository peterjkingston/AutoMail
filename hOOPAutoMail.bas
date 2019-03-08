Attribute VB_Name = "hOOPAutoMail"
Option Explicit
Private m_RulesArray() As Variant, m_App As AcroApp, m_AVDoc As AcroAVDoc, m_PDDoc As AcroPDDoc, m_WhatStr As String
Private m_EmailCount As Integer, m_ProgressCount As Integer, m_ProgressMax As Integer, m_CompletionD As Integer, m_CompletionN As Integer
Enum ProcessTime
    pTimeSeparate = 13
    pTimePrinter = 157
    pTimeEmailer = 9
    pTimeBroker = 232
End Enum
Public Sub Main()
    Dim noErrors As Boolean
    AutoMailWindow.WhatAmIDoing.Visible = True
    AutoMailWindow.RunButton.Enabled = False
    'AutoMailWindow.ButtonTest.Visible = False
    'DoEvents
    AutoMailWindow.CompletionLabel.Visible = True
    AutoMailWindow.CompletionProgressBar.Visible = True
    Select Case noErrors ''TODO: Proof this.
        Case CollectRules ''TODO: Proof this.
        Case InitializeGlobals ''Proof this.
        Case InitializeAcrobat ''TODO: Proof this.
        Case SeparateDocuments ''TODO: Proof this.
        Case Printers ''TODO: Proof this.
        Case Globals.InitializeEmails ''Proof this.
        Case Emailers ''TODO: Proof this.
        Case EmailersBroker ''TODO: Proof this.
        Case CloseRemainders ''TODO: Proof this.
        Case PrintersBroker ''TODO: Proof this.
        Case CloseRemainders ''TODO: Proof this.
        Case ShortcutsNotepad.LogAll ''TODO: Proof this. Different logging ideas? (i.e:Log as you go)
        Case DeleteAll ''TODO: Proof this.
        Case DisplayEmails ''TODO: Proof this.
    End Select
    On Error Resume Next
    If Not m_App Is Nothing Then
        m_AVDoc.Close False
        m_PDDoc.Close
        m_App.CloseAllDocs
        m_App.Exit
    End If
    With AutoMailWindow
        .CompletionProgressBar.value = .CompletionProgressBar.Max
        .CompletionLabel = "100% Complete"
        .WhatAmIDoing = "Complete"
        MsgBox "AutoMail process complete.", , "AutoMail"
        .CompletionProgressBar.Visible = False
        .CompletionLabel.Visible = False
    End With
End Sub
Private Function GetProcessingTime() As Integer
    Dim ARDoc As ARDocument, totalTime As Integer
    For Each ARDoc In g_ARCollection
        totalTime = totalTime + pTimeSeparate
        If ARDoc.IsPrinter Then: totalTime = totalTime + pTimePrinter
        If ARDoc.IsEmailer Then: totalTime = totalTime + pTimeEmailer
        If ARDoc.HasBroker Then: totalTime = totalTime + pTimeBroker
    Next ARDoc
    GetProcessingTime = totalTime
End Function
Private Function CollectRules() As Boolean
    Dim i As Integer, r As Integer, c As Integer
    On Error GoTo EndAll
    m_WhatStr = "Gathering rules..."
    With AutoMailWindow.RuleListBox
        ReDim m_RulesArray(UBound(.List), .columncount - 1) As Variant
        r = UBound(.List)
        For i = 0 To UBound(m_RulesArray)
            For c = 0 To .columncount - 1
                m_RulesArray(i, c) = .List(r, c)
            Next c
            r = r - 1
        Next i
    End With
    CollectRules = True
    Exit Function
EndAll:
    CollectRules = False
    MsgBox "Error collecting rules. AutoMail process ended."
    Exit Function
End Function
Private Sub ProgressBarIterate(byThisCount As Integer)
    If AutoMailWindow.CompletionProgressBar.Max < AutoMailWindow.CompletionProgressBar.value + byThisCount Then
        AutoMailWindow.CompletionProgressBar.Max = AutoMailWindow.CompletionProgressBar + byThisCount
    End If
    AutoMailWindow.CompletionProgressBar.value = AutoMailWindow.CompletionProgressBar.value + byThisCount
End Sub
Private Sub ProgressBarIncreaseMax(byThisCount As Integer)
    AutoMailWindow.CompletionProgressBar.Max = AutoMailWindow.CompletionProgressBar.Max + byThisCount
End Sub
Private Sub CompletionIncreaseMax(byThisCount As Integer)
    Call ProgressBarIncreaseMax(byThisCount)
    m_CompletionD = m_CompletionD + byThisCount
    Call CompletionCalculate
End Sub
Private Sub CompletionIterate(byThisCount As Integer)
    Call ProgressBarIterate(byThisCount)
    m_CompletionN = m_CompletionN + byThisCount
    Call CompletionCalculate
End Sub
Private Sub CompletionCalculate()
    Dim myPercent As Integer
    myPercent = (m_CompletionN / m_CompletionD) * 100
    AutoMailWindow.CompletionLabel = CStr(myPercent) & "% Complete"
    AutoMailWindow.WhatAmIDoing = m_WhatStr
End Sub
Private Function InitializeGlobals() As Boolean
    m_WhatStr = "Initilizing global variables..."
    InitializeGlobals = Globals.Initialize
End Function
Private Function InitializeAcrobat() As Boolean
    On Error GoTo EndAll
    m_WhatStr = "Initializing Adobe Acrobat 10..."
    Set m_App = CreateObject("AcroExch.App")
    Set m_AVDoc = CreateObject("AcroExch.AVDoc")
    m_App.Hide
    If ValidatePDF(m_AVDoc, AutoMailWindow.FileTextBox, "Primay PDF") Then
        Set m_PDDoc = m_AVDoc.GetPDDoc
        InitializeAcrobat = True
    Else
        GoTo EndAll
    End If
    InitializeAcrobat = True
    Exit Function
EndAll:
    InitializeAcrobat = False
    MsgBox "Error initializing Adobe Acrobat Reader 10. AutoMail process ended."
    m_App.CloseAllDocs
    m_App.Exit
    Exit Function
End Function
Private Function ValidatePDF(thisAVDoc As AcroAVDoc, thisPath As String, title As String) As Boolean
    ValidatePDF = thisAVDoc.Open(thisPath, title)
    If thisAVDoc.IsValid = False Then: Error (123)
End Function
Private Function SeparateDocuments() As Boolean
    Dim SO As String, tempAVDoc As New AcroAVDoc, tempPDDoc As New AcroPDDoc, newSO As String, tempPage As Integer, tempDir As String
    Dim i As Integer, ARDoc As New ARDocument, pageCount As Integer, myDbl As Double
    On Error GoTo EndAll
    m_WhatStr = "Separating documents..."
    tempDir = ThisWorkbook.Path & "\AdminParts\Blank.pdf"
    'm_App.Show
    m_App.Hide
    pageCount = m_PDDoc.GetNumPages
    Call CompletionIncreaseMax(pageCount * 129)
    For i = 0 To pageCount
        'myDbl = Timer
        m_WhatStr = "Separating documents... (" & i & "/" & pageCount & " pages)"
        Call CompletionIterate(pTimeSeparate)
        newSO = ShortcutsAcrobat.GetRectStr(m_PDDoc, m_AVDoc, g_SOTop, g_SOBottom, g_SOLeft, g_SORight)
        If i = pageCount Then GoTo MakeNew
        If newSO <> SO Then
            If SO <> "" Then
MakeNew:
                Set ARDoc = New ARDocument
                ARDoc.SO = SO
                Set ARDoc.PDDoc = tempPDDoc
                Set ARDoc.AVDoc = tempAVDoc
                Call g_ARCollection.Add(ARDoc)
                If Not ApplyRules(ARDoc) Then GoTo EndAll
                ARDoc.FileDir = ThisWorkbook.Path & "\Input Directory\" & ARDoc.SO & _
                    " " & ARDoc.DocumentType & ".pdf"
                ARDoc.SavePDF
                ARDoc.ClosePDF
                'Debug.Print Timer - myDbl & " Separate 1 document"
                Set ARDoc = Nothing
            End If
            SO = newSO
            tempPage = -1
            Set tempAVDoc = CreateObject("AcroExch.AVDoc")
            If ValidatePDF(tempAVDoc, tempDir, SO) Then
                Set tempPDDoc = tempAVDoc.GetPDDoc
                tempPDDoc.InsertPages tempPage, m_PDDoc, 0, 1, 0
                tempPDDoc.DeletePages 1, 1
                If m_PDDoc.GetNumPages > 1 Then: m_PDDoc.DeletePages 0, 0
                tempPage = tempPage + 1
            End If
        Else
            tempPDDoc.InsertPages tempPage, m_PDDoc, 0, 1, 0
            If m_PDDoc.GetNumPages > 1 Then: m_PDDoc.DeletePages 0, 0
            tempPage = tempPage + 1
        End If
    Next i
    m_CompletionD = GetProcessingTime
    m_AVDoc.Close True
    SeparateDocuments = True
    Exit Function
EndAll:
    m_App.Show
    SeparateDocuments = False
    m_App.Hide
    MsgBox "Error separating documents. AutoMail process ended."
    Exit Function
End Function
Private Function ApplyRules(ARDoc As ARDocument) As Boolean
    Dim i As Integer
    On Error GoTo EndAll
    'm_WhatStr = "Applying rules..."
    For i = 0 To UBound(m_RulesArray)
        Set ARDoc = ApplyRule(ARDoc, i)
    Next i
    ApplyRules = True
    Exit Function
EndAll:
    ApplyRules = False
    MsgBox "Error applying rules. AutoMail process ended."
    Exit Function
End Function
Private Function ApplyRule(ARDoc As ARDocument, index As Integer) As ARDocument
    Dim trigger As String, condition As String, action As String, accessor As String
    trigger = m_RulesArray(index, 0)
    condition = m_RulesArray(index, 1)
    action = m_RulesArray(index, 2)
    accessor = m_RulesArray(index, 3)
    Select Case trigger
        Case "DocType":
            If ARDoc.DocumentType = condition Then: Set ApplyRule = DoAction(ARDoc, action, accessor, trigger, condition)
        Case "SO#"
            If ARDoc.SO = condition Then: Set ApplyRule = DoAction(ARDoc, action, accessor, trigger, condition)
        Case "PO#"
            If ARDoc.PO = condition Then: Set ApplyRule = DoAction(ARDoc, action, accessor, trigger, condition)
        Case "Customer ID"
            If ARDoc.CustomerID = condition Then: Set ApplyRule = DoAction(ARDoc, action, accessor, trigger, condition)
        Case "Broker"
            If ARDoc.Broker = condition Then: Set ApplyRule = DoAction(ARDoc, action, accessor, trigger, condition)
        Case "EmailAddress"
            If ARDoc.EmailAddress = condition Then: Set ApplyRule = DoAction(ARDoc, action, accessor, trigger, condition)
        Case "StreetAddress"
            If ARDoc.StreetAddress = condition Then: Set ApplyRule = DoAction(ARDoc, action, accessor, trigger, condition)
        Case "FindText"
            If ARDoc.AVDoc.findText(condition, 0, 0, 0) Then: Set ApplyRule = DoAction(ARDoc, action, accessor, trigger, condition)
        Case Else
            MsgBox "Error- Invalid trigger. Please contact the program administrator.", vbOKOnly, "Warning!"
    End Select
    Set ApplyRule = ARDoc
End Function
Private Function DoAction(ARDoc As ARDocument, action As String, accessor As String, trigger As String, condition As String) As ARDocument
    Dim ms As Integer
    Select Case action
        Case "Do Not Email"
            ARDoc.IsEmailer = False
        Case "Do Not Print"
            ARDoc.IsPrinter = False
        Case "Email"
            ARDoc.IsEmailer = True
            ARDoc.EmailAddress = accessor
            Call ARDoc.SavePDF
        Case "CC"
            ARDoc.CC = accessor
        Case "Print"
            ARDoc.IsPrinter = True
        Case "Notify"
            MsgBox condition & " " & trigger & " detected!", vbOKOnly, "Warning!"
        Case "Inspect"
            m_App.Show
            ARDoc.AVDoc.BringToFront
            ms = MsgBox("Click OK to resume processing." & vbCrLf & "Click CANCEL to discard this document.", vbOKCancel, "Information")
            If ms = 2 Then
                ARDoc.IsPrinter = False
                ARDoc.IsEmailer = False
                ARDoc.HasBroker = False
            End If
        Case "Do Nothing"
                ARDoc.IsPrinter = False
                ARDoc.IsEmailer = False
                ARDoc.HasBroker = False
        Case Else
            MsgBox "Error- Invalid action: " & action & ". Please contact the program administrator.", vbOKOnly, "Warning!"
    End Select
    Set DoAction = ARDoc
End Function
Private Function DisplayEmails() As Boolean
    Dim i As Integer
    m_WhatStr = "Displaying Emails..."
    Call CompletionIterate(0)
    For i = 0 To UBound(g_EmailArray)
        On Error Resume Next
        g_EmailArray(i).Display
        'DoEvents
    Next i
    DisplayEmails = True
    Exit Function
EndAll:
    MsgBox "An error occurred while displaying emails. AutoMail process ending."
    DisplayEmails = False
    Exit Function
End Function
Private Function Printers() As Boolean
    Dim ARDoc As ARDocument, printMax As Integer, printVal As Integer, myDbl As Double
    On Error GoTo EndAll
    If g_IsPrintingDisabled Then
        Printers = True
        Exit Function
    End If
    For Each ARDoc In g_ARCollection
        If ARDoc.IsPrinter Then: printMax = printMax + 1
    Next ARDoc
    'myDbl = Timer
    For Each ARDoc In g_ARCollection
        If ARDoc.IsPrinter Then
            m_WhatStr = "Printing customer copies... (" & printVal & "/" & printMax & " pages)"
            printVal = printVal + 1
            Call CompletionIterate(pTimePrinter)
            ARDoc.OpenPDF
            Call ARDoc.PrintSilent(m_App)
            'Debug.Print (ARDoc.SO)
            'DoEvents
            'Debug.Print Timer - myDbl & " Printers"
            'Stop
        End If
    Next ARDoc
    Printers = True
    Exit Function
EndAll:
    MsgBox "An error occured while printing. AutoMail process ended."
    Printers = False
    Exit Function
End Function
Private Function PrintersBroker() As Boolean
    Dim ARDoc As ARDocument, myDir As String, newAVDoc As AcroAVDoc, newPDDoc As AcroPDDoc, printVal As Integer, printMax As Integer
    Dim myDbl As Double
    On Error GoTo EndAll
        If g_IsPrintingDisabled Then
        PrintersBroker = True
        Exit Function
    End If
    For Each ARDoc In g_ARCollection
        If ARDoc.HasBroker Then
            printMax = printMax + 1
        End If
    Next ARDoc
    m_AVDoc.Close True
    m_App.CloseAllDocs
    m_App.Exit
    m_AVDoc.Close False
    For Each ARDoc In g_ARCollection
        If ARDoc.HasBroker Then
            'myDbl = Timer
            m_WhatStr = "Printing brokerage copies... (" & printVal & "/" & printMax & " pages)"
            Call CompletionIterate(pTimeBroker) 'Count time %
            printVal = printVal + 1
            ApplyMessage ARDoc, "BROKERAGE COPY"
            'Debug.Print ARDoc.SO & "Brokerage"
            'Debug.Print Timer - myDbl & " Printers Broker"
            'Stop
        End If
    Next ARDoc
    PrintersBroker = True
    Exit Function
EndAll:
    If (MsgBox(prompt:="An error occured while printing brokerage copies. AutoMail process ended." & vbCrLf & vbCrLf & "Debug?", Buttons:=vbYesNo)) = 6 Then
        Stop
        m_App.Show
        Resume
    End If
    PrintersBroker = False
    Exit Function
End Function
Private Sub ApplyMessage(ARDoc As ARDocument, messageStr As String)
    Dim PDFormApp As AFormApp, PDFields As AFORMAUTLib.Fields, PDField As AFORMAUTLib.Field, newApp As AcroApp
    Set newApp = CreateObject("AcroExch.App")
    Set PDFormApp = CreateObject("AFormAut.App")
    ARDoc.OpenPDF
        Set PDFields = PDFormApp.Fields ''Error here
        Set PDField = PDFormApp.Fields.Add("Text", "text", 0, g_MessageLeft, g_MessageTop, g_MessageRight, g_MessageBottom)
        PDField.value = messageStr
        Call ARDoc.PrintSilent(m_App)
        'Debug.Print ARDoc.SO
    ARDoc.ClosePDF
    newApp.CloseAllDocs
    newApp.Exit
End Sub
Private Function SaveOff() As Boolean
    Dim ARDoc As ARDocument, thisPDDoc As AcroPDDoc, saveDir As String
    On Error GoTo EndAll
    m_WhatStr = "Saving off required documents..."
    For Each ARDoc In g_ARCollection
        If ARDoc.IsEmailer Or ARDoc.HasBroker Then
            Set thisPDDoc = ARDoc.AVDoc.GetPDDoc
            saveDir = ThisWorkbook.Path & "\Input Directory\" & ARDoc.SO & _
            " " & ARDoc.DocumentType & ".pdf"
            If Not thisPDDoc.Save(PDSaveFull, saveDir) Then GoTo EndAll
        End If
    Next ARDoc
    SaveOff = True
    Exit Function
EndAll:
    MsgBox "An error occured while attempting to save " & ARDoc.SO & " " & ARDoc.DocumentType & ".pdf. Automail process ending."
    SaveOff = False
    Exit Function
End Function
Private Function CloseEmailers() As Boolean
    Dim ARDoc As ARDocument
    On Error GoTo EndAll
    m_WhatStr = "Closing files for Email..."
    For Each ARDoc In g_ARCollection
        If ARDoc.HasBroker Or ARDoc.IsEmailer Then
            ARDoc.ClosePDF
            ARDoc.IsOpen = False
        End If
    Next ARDoc
    CloseEmailers = True
    Exit Function
EndAll:
    Stop
    Resume
    MsgBox "An error occured while closing PDFs. AutoMail process ending."
    CloseEmailers = False
    Exit Function
End Function
Private Function Emailers() As Boolean
    Dim ARDoc As ARDocument, emailVal As Integer, emailMax As Integer, myDbl As Double
    On Error GoTo EndAll
    If g_IsEmailDisabled Then
        Emailers = True
        Exit Function
    End If
    For Each ARDoc In g_ARCollection
        If ARDoc.IsEmailer Then: emailMax = emailMax + 1
    Next ARDoc
    For Each ARDoc In g_ARCollection
        If ARDoc.IsEmailer Then
            'myDbl = Timer
            m_WhatStr = "Generating Emails for customers... (" & emailVal & "/" & emailMax & " Emails)"
            Call CompletionIterate(pTimeEmailer)
            emailVal = emailVal + 1
            Set g_EmailArray(m_EmailCount) = ShortcutsOutlook.Emailer(ARDoc)
            'Debug.Print ARDoc.SO & " emailed"
            m_EmailCount = m_EmailCount + 1
            'Debug.Print Timer - myDbl & " Emailers"
            'Stop
        End If
    Next ARDoc
    Emailers = True
    Exit Function
EndAll:
    MsgBox "An error occured while generating emails. AutoMail process ended."
    Emailers = False
    Exit Function
End Function
Private Function EmailersBroker() As Boolean
    Dim ARDoc As ARDocument, emailMax As Integer, emailVal As Integer
    On Error Resume Next
    If g_IsEmailDisabled Then
        EmailersBroker = True
        Exit Function
    End If
    For Each ARDoc In g_ARCollection
        If ARDoc.IsEmailer Then: emailMax = emailMax + 1
    Next ARDoc
    For Each ARDoc In g_ARCollection
        If ARDoc.HasBroker Then
            m_WhatStr = "Generating Emails for brokers... (" & emailVal & "/" & emailMax & " Emails)"
            Call CompletionIterate(pTimeEmailer)
            emailVal = emailVal + 1
            Set g_EmailArray(m_EmailCount) = ShortcutsOutlook.EmailerBroker(ARDoc)
            m_EmailCount = m_EmailCount + 1
        End If
    Next ARDoc
    EmailersBroker = True
    Exit Function
EndAll:
    MsgBox "An error occured while generating emails. AutoMail process ended."
    EmailersBroker = False
    Exit Function
End Function
Private Function CloseRemainders() As Boolean
    Dim ARDoc As ARDocument
    On Error GoTo EndAll
    m_WhatStr = "Closing remaining files..."
    For Each ARDoc In g_ARCollection
        If ARDoc.IsOpen Then
            ARDoc.ClosePDF
        End If
    Next ARDoc
    m_AVDoc.Close True
    m_PDDoc.Close
    m_App.CloseAllDocs
    m_App.Exit
    CloseRemainders = True
    Exit Function
EndAll:
    MsgBox "An error occurred while closing PDF Files. AutoMail process ended."
    CloseRemainders = False
    Exit Function
End Function
Private Function DeleteAll() As Boolean
    Dim ARDoc As ARDocument
    On Error Resume Next
    m_WhatStr = "Deleting invoice copies..."
    For Each ARDoc In g_ARCollection
        If Not ARDoc.IsOpen Then
            Kill ARDoc.FileDir
        Else
            ARDoc.ClosePDF
            Kill ARDoc.FileDir
            'If getinputstate <> 0 Then: 'DoEvents
        End If
    Next ARDoc
    DeleteAll = True
    Exit Function
EndAll:
    MsgBox "An error occured while deleting " & ARDoc.SO & " " & ARDoc.DocumentType & ".pdf. Automail process ending."
    DeleteAll = False
    Exit Function
End Function
Private Function DisplayEmailers() As Boolean
    Dim ARDoc As ARDocument, i As Integer
    On Error GoTo EndAll
    If g_IsEmailDisabled Then
        DisplayEmailers = True
        Exit Function
    End If
    For i = 0 To UBound(g_EmailArray)
        g_EmailArray(i).Display
        'DoEvents
    Next i
    DisplayEmailers = True
    Exit Function
EndAll:
    DisplayEmailers = False
    Exit Function
End Function
