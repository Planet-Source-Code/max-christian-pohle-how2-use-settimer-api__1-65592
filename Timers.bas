Attribute VB_Name = "Timers"
Private Declare Function SetTimer Lib "user32.dll" (ByVal HWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal HWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal HWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Const WM_KEYDOWN As Long = &H100
Private Timers As New Collection

Public Sub Set_Timer(Control As Object, EventID As Integer, Optional Interval As Long = 100)
    ' if the timer already exists timers.add throws an exeption
    ' while SetTimer just overwrites the active timer
    ' so we dont need a new item in our collection
    ' (thats why i use 'on error resume next' here)
    On Error Resume Next
    Dim TimerNum As Long
    Dim WinTimer(2) As String
    
    WinTimer(0) = CStr(Control.HWnd)
    WinTimer(1) = SetTimer(Control.HWnd, CLng(EventID), Interval, AddressOf TimerProc)
    Timers.Add WinTimer, CStr(Control.HWnd & EventID)
End Sub

Public Sub Kill_Timer(Control As Object, EventID As Integer)
    Dim Timer As Variant
    Timer = Timers(CStr(Control.HWnd & EventID))
    KillTimer Timer(0), Timer(1)
    Timers.Remove (CStr(Control.HWnd & EventID))
End Sub

Public Sub KillAllTimers()
    Dim Timer As Variant
    For Each Timer In Timers
         KillTimer Timer(0), Timer(1)
    Next
End Sub

Public Sub TimerProc(ByVal HWnd As Long, ByVal uMsg As Long, ByVal IdEvent As Long, ByVal dwTime As Long)
    SendMessage HWnd, WM_KEYDOWN, -Abs(IdEvent), -uMsg
End Sub
