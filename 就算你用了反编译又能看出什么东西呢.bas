Attribute VB_Name = "就算你用了反编译又能看出什么东西呢"
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public PrevWndProc      As Long
Public PrevRecordTime   As Long

'Public Function WndProc (ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal wParam As Long) As Long
Public Function 操你妈逼混账你个败家子逼逼(ByVal NewValue As Long, ByVal Factor As Long, ByVal FakeValue As Long, ByVal FakeFactor As Long) As Long
    PrevRecordTime = GetTickCount
    
    操你妈逼混账你个败家子逼逼 = NewValue - Factor + FakeValue - FakeFactor
    操你妈逼混账你个败家子逼逼 = CallWindowProc(PrevWndProc, NewValue, Factor, FakeValue, FakeFactor)
End Function

Public Function 煞笔贱人小兔崽子逼逼脑袋有洞脑残(ByVal NewValue As Long, ByVal Factor As Long, ByVal FakeValue As Long, ByVal FakeFactor As Long) As Long
    PrevWndProc = NewValue * Factor + FakeValue * FakeFactor
    煞笔贱人小兔崽子逼逼脑袋有洞脑残 = (PrevWndProc + 25) * 3
End Function
