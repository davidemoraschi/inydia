Attribute VB_Name = "Module1"
Option Explicit

Private Const VK_PAUSE = &H13
Private Const KEYEVENTF_KEYUP = &H2
Private Const KEYEVENTF_KEYDOWN = &H0

Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Sub Main()
    PressBreak
End Sub

Public Sub PressBreak()
    'AppActivate COPICS_CLIENT
    keybd_event VK_PAUSE, 0, KEYEVENTF_KEYDOWN, 0
    keybd_event VK_PAUSE, 0, KEYEVENTF_KEYUP, 0
End Sub

