<div align="center">

## How2 use SetTimer\-API


</div>

### Description

One timer is not enough? Have you ever tried to add multiple timer-controls to your form? Well- This wont work because they will not work together. This example shows you how to realize it using the SetTimer-API: very simple and effective!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2006-05-29 20:27:44
**By**             |[Max Christian Pohle](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/max-christian-pohle.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[How2\_use\_S199939672006\.zip](https://github.com/Planet-Source-Code/max-christian-pohle-how2-use-settimer-api__1-65592/archive/master.zip)

### API Declarations

```
Private Declare Function SetTimer Lib "user32.dll" ( _
  ByVal HWnd As Long, _
  ByVal nIDEvent As Long, _
  ByVal uElapse As Long, _
  ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" ( _
  ByVal HWnd As Long, _
  ByVal nIDEvent As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
  ByVal HWnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  ByRef lParam As Any) As Long
Private Const WM_KEYDOWN As Long = &amp;H100
Private Timers As New Collection
Public Sub Set_Timer(Control As Object, _
  EventID As Integer, Optional Interval As Long = 100)
  ' if the timer already exists timers.add
  '   throws an exeption
  ' while SetTimer just overwrites the act
  '   ive timer
  ' so we dont need a new item in our coll
  '   ection
  ' (thats why i use 'on error resume next
  ' here)
  On Error Resume Next
  Dim TimerNum As Long
  Dim WinTimer(2) As String
  WinTimer(0) = CStr(Control.HWnd)
  WinTimer(1) = SetTimer(Control.HWnd, CLng(EventID), Interval, AddressOf TimerProc)
  Timers.Add WinTimer, CStr(Control.HWnd &amp; EventID)
End Sub
Public Sub Kill_Timer(Control As Object, EventID As Integer)
  Dim Timer As Variant
  Timer = Timers(CStr(Control.HWnd &amp; EventID))
  KillTimer Timer(0), Timer(1)
  Timers.Remove (CStr(Control.HWnd &amp; EventID))
End Sub
Public Sub KillAllTimers()
  Dim Timer As Variant
  For Each Timer In Timers
    KillTimer Timer(0), Timer(1)
  Next
End Sub
Public Sub TimerProc( _
  ByVal HWnd As Long, ByVal uMsg As Long, ByVal _
  IdEvent As Long, ByVal dwTime As Long)
  SendMessage HWnd, WM_KEYDOWN, -Abs(IdEvent), -uMsg
End Sub
```





