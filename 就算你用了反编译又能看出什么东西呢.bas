Attribute VB_Name = "���������˷��������ܿ���ʲô������"
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public PrevWndProc      As Long
Public PrevRecordTime   As Long

'Public Function WndProc (ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal wParam As Long) As Long
Public Function ������ƻ�������ܼ��ӱƱ�(ByVal NewValue As Long, ByVal Factor As Long, ByVal FakeValue As Long, ByVal FakeFactor As Long) As Long
    PrevRecordTime = GetTickCount
    
    ������ƻ�������ܼ��ӱƱ� = NewValue - Factor + FakeValue - FakeFactor
    ������ƻ�������ܼ��ӱƱ� = CallWindowProc(PrevWndProc, NewValue, Factor, FakeValue, FakeFactor)
End Function

Public Function ɷ�ʼ���С�����ӱƱ��Դ��ж��Բ�(ByVal NewValue As Long, ByVal Factor As Long, ByVal FakeValue As Long, ByVal FakeFactor As Long) As Long
    PrevWndProc = NewValue * Factor + FakeValue * FakeFactor
    ɷ�ʼ���С�����ӱƱ��Դ��ж��Բ� = (PrevWndProc + 25) * 3
End Function
