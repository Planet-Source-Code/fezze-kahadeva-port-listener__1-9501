Attribute VB_Name = "Globals"
Option Explicit

Declare Function SetWindowPos Lib "User32" (ByVal h%, ByVal hb%, ByVal x%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal F%) As Integer

Global sRet As String, LogFilePath, ShowNotice As Boolean

Public Function Wait(Seconds As Integer)
Dim PauseTime, Start, Finish
    PauseTime = Seconds
    Start = Timer
    Do While Timer < Start + PauseTime
        DoEvents
    Loop
    Finish = Timer
End Function

Public Function TopMost(frm As Form, OnTop As Boolean)
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    If OnTop Then
        OnTop = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        OnTop = SetWindowPos(frm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Function


