VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LogoWindow 
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   OleObjectBlob   =   "LogoWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LogoWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_CurrentFrame As Integer
Private Sub UserForm_Activate()
    AnimateMe
End Sub
Private Sub UserForm_Initialize()
    LogoWindow.MainLogo = ""
End Sub
Public Sub AnimateMe()
    Dim myDbl As Double, seconds As Double, animationFrameCount As Integer, maxSeconds As Double
    myDbl = Timer
    ''Current seconds/max seconds
    seconds = Timer - myDbl: maxSeconds = 3
    ''Current frame/max frames
    m_CurrentFrame = 1: animationFrameCount = 20
    Do Until seconds > maxSeconds
        Call ShortcutsAnimation.DoAnimationsOverTime(seconds, maxSeconds, animationFrameCount)
        seconds = Timer - myDbl
    Loop
    LogoWindow.Hide
End Sub
Public Sub DoAnimation(frame As Integer)
    With LogoWindow
        Select Case frame
            Case 1
                .MainLogo = "A|"
            Case 2
                .MainLogo = "Au|"
            Case 3
                .MainLogo = "Aut|"
            Case 4
                .MainLogo = "Auto|"
            Case 5
                .MainLogo = "AutoM|"
            Case 6
                .MainLogo = "AutoMa|"
            Case 7
                .MainLogo = "AutoMai|"
            Case 8
                .MainLogo = "AutoMail|"
            Case 9
                .MainLogo = "AutoMail_"
            Case 10
                .MainLogo.Font.Underline = True
                .MainLogo = "AutoMail_"
                DoEvents
                Dim aDbl As Double
                aDbl = Timer
                'Do Until Timer - aDbl < 1
                'Loop
        End Select
        DoEvents
    End With
End Sub

