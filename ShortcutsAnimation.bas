Attribute VB_Name = "ShortcutsAnimation"
Option Explicit
Public Sub DoAnimationsOverTime(time As Double, maxTime As Double, frameCount As Integer)
    Dim i As Integer
    If (LogoWindow.m_CurrentFrame / frameCount) < (time / maxTime) Then
        Call LogoWindow.DoAnimation(LogoWindow.m_CurrentFrame)
        LogoWindow.m_CurrentFrame = LogoWindow.m_CurrentFrame + 1
    End If
End Sub
