VERSION 5.00
Begin VB.Form ���������ø�ë�߷����� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ƽ�ѽ"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton �ǲ��ǿ����ؼ��������Լ��������� 
      Caption         =   "-5"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.Label �ܿ����ؼ������˲����� 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label ɵ���㿴���ؼ�����û�õ� 
      AutoSize        =   -1  'True
      Caption         =   "��ǰѪ����1000"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1260
   End
   Begin VB.Label �������ص�������ɲ˼� 
      AutoSize        =   -1  'True
      Caption         =   "�����Ѫ���ĳ�233�ͳɹ��ˡ�"
      Height          =   195
      Left            =   652
      TabIndex        =   0
      Top             =   960
      Width           =   2430
   End
End
Attribute VB_Name = "���������ø�ë�߷�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'String decode: ɵ������������������Ա�������������()
'Init. vars:    �������˸����Ͱ׳ջ����Բ�()
'Loop check:    �����޳����������ɶɵ�������ܼ���С����()
'Set value:     �����վ��ӿ������ֵܵܲ�ȱ()
'Get value:     ������Ǻǲ��������()
'Mess up:       ���ĵ����Դ��ж�����ɷ��ɷ������()

Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
'Used
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function CloseWindowStation Lib "user32" (ByVal hWinSta As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
'Used
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Used
Private Declare Function GetClassInfoEx Lib "user32.dll" Alias "GetClassInfoExA" (ByVal hinstance As Long, ByVal lpcstr As String, ByRef lpwndclassexa As WNDCLASSEX) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
'Used
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hinstance As Long, lpParam As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal numBytes As Long)
'Used
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function AddPrinterDriver Lib "winspool.drv" Alias "AddPrinterDriverA" (ByVal pName As String, ByVal Level As Long, pDriverInfo As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
'Used
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowsHook Lib "user32" Alias "SetWindowsHookA" (ByVal nFilterType As Long, ByVal pfnFilterProc As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function ���ܿ�����������API�������� Lib "user32" () As Long

Private Const MAX_SIZE = 16                     'The max size of real blood value data region

Dim NumberData()            As Byte             'Real blood value data
Dim MemIndex(3)             As Byte             'Memory index of real blood value data
Dim FakeBlood               As Long             'Fake blood value data
Dim PrevParent              As Long             'Initial parent window of this window
Dim ButtonHwnd              As Long             'The handle to the button
Dim Exiting                 As Boolean          'Is program exiting

Dim CE_Changed              As Boolean          'User changed the fake value by CE

Dim LABEL_CURRENT_BLOOD     As Variant          '"��ǰѪ����"
Dim CE_PROMPT_1             As Variant          '"�ȵȣ���������"
Dim CE_PROMPT_2             As Variant          '"����ȥ�����ֵ�ľ��ķǳ�ǿ..."
Dim CE_PROMPT_3             As Variant          '"���ָ���ԭ������ֵ����"
Dim SUCCEED_PROMPT          As Variant          '"���ģ�������ô�����ģ�����"

Function �������������ɶ�����޳����޻���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 6.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 26.7) = InArray((i * 2 - 22.7)) Xor (i - 16)
    Next i
    �������������ɶ�����޳����޻��� = StrConv(tmpArray, vbUnicode)
End Function

Function ����ַ����������������������վ��ӿ�����������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.6) = InArray((i * 2 - 25.2)) Xor (i - 12.9)
    Next i
    ����ַ����������������������վ��ӿ����������� = StrConv(tmpArray, vbUnicode)
End Function

Function �Դ��ж�������ִ�������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 17.7) = InArray((i * 2 - 25.8)) Xor (i - 22)
    Next i
    �Դ��ж�������ִ������� = StrConv(tmpArray, vbUnicode)
End Function

Function �߹������޳�������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 8.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 12.3) = InArray((i * 2 - 3.8)) Xor (i - 16.7)
    Next i
    �߹������޳������� = StrConv(tmpArray, vbUnicode)
End Function

Function ��֪�ϰ��߹���֪�ϰ��Ʊ����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 8.8) = InArray((i * 2 - 15)) Xor (i - 3)
    Next i
    ��֪�ϰ��߹���֪�ϰ��Ʊ���� = StrConv(tmpArray, vbUnicode)
End Function

Function �Ʊ�С����������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 6.1) = InArray((i * 2 - 6.1)) Xor (i - 11.1)
    Next i
    �Ʊ�С���������������� = StrConv(tmpArray, vbUnicode)
End Function

Function �����������վ��ӿ���ܳ�ܼ�����װ��(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.5) = InArray((i * 2 - 12.2)) Xor (i - 17.3)
    Next i
    �����������վ��ӿ���ܳ�ܼ�����װ�� = StrConv(tmpArray, vbUnicode)
End Function

Function ��ǺǺ���Ƥ���������վ��ӿ�������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 26 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 27.6) = InArray((i * 2 - 1.9)) Xor (i - 17.8)
    Next i
    ��ǺǺ���Ƥ���������վ��ӿ������� = StrConv(tmpArray, vbUnicode)
End Function

Function ������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 15.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 12.7) = InArray((i * 2 - 13.4)) Xor (i - 11.3)
    Next i
    ������������ = StrConv(tmpArray, vbUnicode)
End Function

Function �������Ա�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 15.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13.8) = InArray((i * 2 - 4.2)) Xor (i - 16.2)
    Next i
    �������Ա� = StrConv(tmpArray, vbUnicode)
End Function

Function �𱯰����������ƤܳС���������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 9.9) = InArray((i * 2 - 14.4)) Xor (i - 28.8)
    Next i
    �𱯰����������ƤܳС��������� = StrConv(tmpArray, vbUnicode)
End Function

Function �Դ��ж�����Ƥ�����޳�����������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.6) = InArray((i * 2 - 11.1)) Xor (i - 2.3)
    Next i
    �Դ��ж�����Ƥ�����޳����������� = StrConv(tmpArray, vbUnicode)
End Function

Function �ƹ�����������ְܼ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 7.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.8) = InArray((i * 2 - 2.2)) Xor (i - 13.4)
    Next i
    �ƹ�����������ְܼ��� = StrConv(tmpArray, vbUnicode)
End Function

Function ���ܹ�����������ĳ���ɵ��(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 11.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.2) = InArray((i * 2 - 19.7)) Xor (i - 14.3)
    Next i
    ���ܹ�����������ĳ���ɵ�� = StrConv(tmpArray, vbUnicode)
End Function

Function ��������ܳ������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 26.5) = InArray((i * 2 - 7.5)) Xor (i - 18.3)
    Next i
    ��������ܳ������������ = StrConv(tmpArray, vbUnicode)
End Function

Function ɵ�Ʋ���������ǺǺ���Ƥ�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 14.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 27.3) = InArray((i * 2 - 1.2)) Xor (i - 0.9)
    Next i
    ɵ�Ʋ���������ǺǺ���Ƥ����� = StrConv(tmpArray, vbUnicode)
End Function

Function �ܵܺ����޳����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.3) = InArray((i * 2 - 17.1)) Xor (i - 25.2)
    Next i
    �ܵܺ����޳���� = StrConv(tmpArray, vbUnicode)
End Function

Function ɵ��С��������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 25.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.9) = InArray((i * 2 - 17)) Xor (i - 24)
    Next i
    ɵ��С�������� = StrConv(tmpArray, vbUnicode)
End Function

Function �����������ƹ���������������ĵ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 22.3) = InArray((i * 2 - 8.2)) Xor (i - 12.1)
    Next i
    �����������ƹ���������������ĵ� = StrConv(tmpArray, vbUnicode)
End Function

Function �Ǻǲ����������վ��ӿ�������С���鹷����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 13.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 10.1) = InArray((i * 2 - 6.9)) Xor (i - 6.6)
    Next i
    �Ǻǲ����������վ��ӿ�������С���鹷���� = StrConv(tmpArray, vbUnicode)
End Function

Function �����������������ܳ�����Ʊ������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 20 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 12) = InArray((i * 2 - 4.6)) Xor (i - 9)
    Next i
    �����������������ܳ�����Ʊ������ = StrConv(tmpArray, vbUnicode)
End Function

Function �������������Ʊ����Ա�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 12.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 29) = InArray((i * 2 - 29.5)) Xor (i - 29)
    Next i
    �������������Ʊ����Ա� = StrConv(tmpArray, vbUnicode)
End Function

Function ���ĵĶ��ĵ����϶��ĵĴ������޳�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 6.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 20.5) = InArray((i * 2 - 21.4)) Xor (i - 16.9)
    Next i
    ���ĵĶ��ĵ����϶��ĵĴ������޳� = StrConv(tmpArray, vbUnicode)
End Function

Function ���ҷ����������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 29.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.7) = InArray((i * 2 - 14)) Xor (i - 25.6)
    Next i
    ���ҷ���������� = StrConv(tmpArray, vbUnicode)
End Function

Function �������Դ��ж�����ܳ(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 18.4) = InArray((i * 2 - 0.1)) Xor (i - 8.5)
    Next i
    �������Դ��ж�����ܳ = StrConv(tmpArray, vbUnicode)
End Function

Function �ܵ����������Ʊ��Դ��ж��Ǻ�����ģ����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 1.3) = InArray((i * 2 - 17.7)) Xor (i - 7.8)
    Next i
    �ܵ����������Ʊ��Դ��ж��Ǻ�����ģ���� = StrConv(tmpArray, vbUnicode)
End Function

Function ��������ģ�������ƶ��ĵĺǺ�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.9) = InArray((i * 2 - 29.1)) Xor (i - 25.7)
    Next i
    ��������ģ�������ƶ��ĵĺǺ����� = StrConv(tmpArray, vbUnicode)
End Function

Function �����������������ĺǺǺ���Ƥ(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 13.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.2) = InArray((i * 2 - 27.1)) Xor (i - 10.4)
    Next i
    �����������������ĺǺǺ���Ƥ = StrConv(tmpArray, vbUnicode)
End Function

Function �ܼ��Ӵ�����֪�ϰ���װ���˵�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 25.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 26.6) = InArray((i * 2 - 19.1)) Xor (i - 2.3)
    Next i
    �ܼ��Ӵ�����֪�ϰ���װ���˵� = StrConv(tmpArray, vbUnicode)
End Function

Function �ܼ��Ӷ��Ƴ���������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 25.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 12.7) = InArray((i * 2 - 17.3)) Xor (i - 5.3)
    Next i
    �ܼ��Ӷ��Ƴ��������������� = StrConv(tmpArray, vbUnicode)
End Function

Function ���ϵܵ����������������ܵܵ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 11.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.8) = InArray((i * 2 - 27.1)) Xor (i - 10)
    Next i
    ���ϵܵ����������������ܵܵ� = StrConv(tmpArray, vbUnicode)
End Function

Function �Ʊ��������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 13.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 22.1) = InArray((i * 2 - 27.8)) Xor (i - 24.3)
    Next i
    �Ʊ�������� = StrConv(tmpArray, vbUnicode)
End Function

Function �ǺǼ��ֲ������С����������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 5.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.1) = InArray((i * 2 - 2.4)) Xor (i - 17.4)
    Next i
    �ǺǼ��ֲ������С���������� = StrConv(tmpArray, vbUnicode)
End Function

Function ܳС��������Դ��ж�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 10.5) = InArray((i * 2 - 22.3)) Xor (i - 25)
    Next i
    ܳС��������Դ��ж� = StrConv(tmpArray, vbUnicode)
End Function

Function ��ȱ����������������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 24.5) = InArray((i * 2 - 20.8)) Xor (i - 29.6)
    Next i
    ��ȱ���������������������� = StrConv(tmpArray, vbUnicode)
End Function

Function �������������ɶС������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 29.2) = InArray((i * 2 - 27.5)) Xor (i - 9.8)
    Next i
    �������������ɶС������ = StrConv(tmpArray, vbUnicode)
End Function

Function װ�����������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 0) = InArray((i * 2 - 10.4)) Xor (i - 18)
    Next i
    װ����������������� = StrConv(tmpArray, vbUnicode)
End Function

Function ������Ӱܼ������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 25.3) = InArray((i * 2 - 13.7)) Xor (i - 5.6)
    Next i
    ������Ӱܼ������ = StrConv(tmpArray, vbUnicode)
End Function

Function �������������������˻���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 7.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 5.6) = InArray((i * 2 - 13.7)) Xor (i - 4)
    Next i
    �������������������˻��� = StrConv(tmpArray, vbUnicode)
End Function

Function �����߹��������޳ܹ�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 6.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 28.9) = InArray((i * 2 - 6.9)) Xor (i - 8.3)
    Next i
    �����߹��������޳ܹ����� = StrConv(tmpArray, vbUnicode)
End Function

Function ����ɷ�ʻ���ɵ��(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.7) = InArray((i * 2 - 24)) Xor (i - 15.6)
    Next i
    ����ɷ�ʻ���ɵ�� = StrConv(tmpArray, vbUnicode)
End Function

Function ���������������������վ��ӿ��������߹������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.8) = InArray((i * 2 - 8.5)) Xor (i - 8.3)
    Next i
    ���������������������վ��ӿ��������߹������ = StrConv(tmpArray, vbUnicode)
End Function

Function �ƱƱ�����С������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.5) = InArray((i * 2 - 22.6)) Xor (i - 3.4)
    Next i
    �ƱƱ�����С������ = StrConv(tmpArray, vbUnicode)
End Function

Function �������ϲ�ȱ���ƹ����ӱƱ�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 12.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 20.5) = InArray((i * 2 - 1.4)) Xor (i - 15.6)
    Next i
    �������ϲ�ȱ���ƹ����ӱƱ����� = StrConv(tmpArray, vbUnicode)
End Function

Function �Դ��ж�װ�����ƱƳ������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 18 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 28) = InArray((i * 2 - 13)) Xor (i - 26.4)
    Next i
    �Դ��ж�װ�����ƱƳ������ = StrConv(tmpArray, vbUnicode)
End Function

Function ���������С������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 29.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 18.8) = InArray((i * 2 - 19.1)) Xor (i - 29.7)
    Next i
    ���������С������ = StrConv(tmpArray, vbUnicode)
End Function

Function �������������ɶ����ܵ������Բк����޳�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 30) = InArray((i * 2 - 19.1)) Xor (i - 20.8)
    Next i
    �������������ɶ����ܵ������Բк����޳� = StrConv(tmpArray, vbUnicode)
End Function

Function ܳ��ȱ������Ƴ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 24.3) = InArray((i * 2 - 18.8)) Xor (i - 4.1)
    Next i
    ܳ��ȱ������Ƴ��� = StrConv(tmpArray, vbUnicode)
End Function

Function ���ɵ�������ƱƳ����������Դ��ж�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 15.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.9) = InArray((i * 2 - 7.1)) Xor (i - 10.9)
    Next i
    ���ɵ�������ƱƳ����������Դ��ж� = StrConv(tmpArray, vbUnicode)
End Function

Function ���������ɷ�ʲ�ȱ��������ֱ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.9) = InArray((i * 2 - 27)) Xor (i - 17.8)
    Next i
    ���������ɷ�ʲ�ȱ��������ֱ� = StrConv(tmpArray, vbUnicode)
End Function

Function �Բ����������Ӳ���ֹ�������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 11.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 1.2) = InArray((i * 2 - 1.6)) Xor (i - 9.8)
    Next i
    �Բ����������Ӳ���ֹ������������� = StrConv(tmpArray, vbUnicode)
End Function

Function ���������������������������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 24.9) = InArray((i * 2 - 20.9)) Xor (i - 3.4)
    Next i
    ��������������������������������� = StrConv(tmpArray, vbUnicode)
End Function

Function ���������Բ��������ܲ�������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 27) = InArray((i * 2 - 17.3)) Xor (i - 2.3)
    Next i
    ���������Բ��������ܲ������� = StrConv(tmpArray, vbUnicode)
End Function

Function ��Ķ�����֪�ϰ�����������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 2.2) = InArray((i * 2 - 8.8)) Xor (i - 28.9)
    Next i
    ��Ķ�����֪�ϰ����������� = StrConv(tmpArray, vbUnicode)
End Function

Function �������������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 7.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 2.8) = InArray((i * 2 - 28.1)) Xor (i - 19.7)
    Next i
    ������������������� = StrConv(tmpArray, vbUnicode)
End Function

Function ��ܳ����Ƥ��(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 25.9) = InArray((i * 2 - 10.7)) Xor (i - 23)
    Next i
    ��ܳ����Ƥ�� = StrConv(tmpArray, vbUnicode)
End Function

Function ��ܵ���֪�ϰ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 20.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 20.9) = InArray((i * 2 - 25.7)) Xor (i - 12.5)
    Next i
    ��ܵ���֪�ϰ� = StrConv(tmpArray, vbUnicode)
End Function

Function ����Ƥ���������ҷ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 18 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 29.5) = InArray((i * 2 - 22.3)) Xor (i - 24.1)
    Next i
    ����Ƥ���������ҷ� = StrConv(tmpArray, vbUnicode)
End Function

Function ���������蹷����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13.8) = InArray((i * 2 - 22.6)) Xor (i - 2.6)
    Next i
    ���������蹷���� = StrConv(tmpArray, vbUnicode)
End Function

Function ���������ܼ�����������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13.9) = InArray((i * 2 - 16.8)) Xor (i - 29.2)
    Next i
    ���������ܼ����������������� = StrConv(tmpArray, vbUnicode)
End Function

Function ��ģ����������������ɷ�ʴ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 8.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 0.2) = InArray((i * 2 - 5.6)) Xor (i - 0.1)
    Next i
    ��ģ����������������ɷ�ʴ��� = StrConv(tmpArray, vbUnicode)
End Function

Function ���������������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 27.2) = InArray((i * 2 - 19.9)) Xor (i - 29.5)
    Next i
    ��������������������� = StrConv(tmpArray, vbUnicode)
End Function

Function �����������޲���������Ǻ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 3.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 29) = InArray((i * 2 - 26.7)) Xor (i - 8.2)
    Next i
    �����������޲���������Ǻ� = StrConv(tmpArray, vbUnicode)
End Function

Function �Դ��ж�����Ƥ��������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16) = InArray((i * 2 - 6.4)) Xor (i - 29.4)
    Next i
    �Դ��ж�����Ƥ�������� = StrConv(tmpArray, vbUnicode)
End Function

Function �ܵ������֪�ϰ������ӹ����������ɷ��(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.4) = InArray((i * 2 - 12.4)) Xor (i - 22)
    Next i
    �ܵ������֪�ϰ������ӹ����������ɷ�� = StrConv(tmpArray, vbUnicode)
End Function

Function ��������߹�С����������С�������ϼ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 1.4) = InArray((i * 2 - 7.9)) Xor (i - 13.2)
    Next i
    ��������߹�С����������С�������ϼ��� = StrConv(tmpArray, vbUnicode)
End Function

Function ������װ��ܳ�Ա�С���������������ɶ(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 22.8) = InArray((i * 2 - 27)) Xor (i - 3.7)
    Next i
    ������װ��ܳ�Ա�С���������������ɶ = StrConv(tmpArray, vbUnicode)
End Function

Function �����վ��ӿ������ҷ�������ɷ���������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 22.3) = InArray((i * 2 - 14.1)) Xor (i - 20.4)
    Next i
    �����վ��ӿ������ҷ�������ɷ��������� = StrConv(tmpArray, vbUnicode)
End Function

Function ��������������վ��ӿ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 17.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.9) = InArray((i * 2 - 10.2)) Xor (i - 21.2)
    Next i
    ��������������վ��ӿ��� = StrConv(tmpArray, vbUnicode)
End Function

Function ����װ����ģ����ܳܳɷ��װ��(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 12.2) = InArray((i * 2 - 25.8)) Xor (i - 26.2)
    Next i
    ����װ����ģ����ܳܳɷ��װ�� = StrConv(tmpArray, vbUnicode)
End Function

Function ���С�����Ӳ�����Ǻǲ���ֲ�ȱ����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.8) = InArray((i * 2 - 23.1)) Xor (i - 27.7)
    Next i
    ���С�����Ӳ�����Ǻǲ���ֲ�ȱ���� = StrConv(tmpArray, vbUnicode)
End Function

Function �����Ӽ����������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 25) = InArray((i * 2 - 21.1)) Xor (i - 19.8)
    Next i
    �����Ӽ���������� = StrConv(tmpArray, vbUnicode)
End Function

Function ����ֺǺǶ��Ʊ���С�������Բ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 9.1) = InArray((i * 2 - 22.9)) Xor (i - 7.8)
    Next i
    ����ֺǺǶ��Ʊ���С�������Բ� = StrConv(tmpArray, vbUnicode)
End Function

Function ��ȱ�����������������ֹ������߹�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 17.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 6.6) = InArray((i * 2 - 15.6)) Xor (i - 25.2)
    Next i
    ��ȱ�����������������ֹ������߹����� = StrConv(tmpArray, vbUnicode)
End Function

Function �����վ��ӿ��������߹������޳�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 9.4) = InArray((i * 2 - 27.7)) Xor (i - 9.8)
    Next i
    �����վ��ӿ��������߹������޳� = StrConv(tmpArray, vbUnicode)
End Function

Function ��ܳ�������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.1) = InArray((i * 2 - 5.4)) Xor (i - 12.4)
    Next i
    ��ܳ������������� = StrConv(tmpArray, vbUnicode)
End Function

Function ɷ�ʲݶ�����ģ������С�����߹��ܼ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 22.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 8.7) = InArray((i * 2 - 17.1)) Xor (i - 12.8)
    Next i
    ɷ�ʲݶ�����ģ������С�����߹��ܼ��� = StrConv(tmpArray, vbUnicode)
End Function

Function �ܼ�������Ķ��ĵ��Դ��ж��Ǻ����ϲ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 26.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 3.2) = InArray((i * 2 - 2.1)) Xor (i - 3.9)
    Next i
    �ܼ�������Ķ��ĵ��Դ��ж��Ǻ����ϲ� = StrConv(tmpArray, vbUnicode)
End Function

Function ��������װ��(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.1) = InArray((i * 2 - 28.5)) Xor (i - 6.2)
    Next i
    ��������װ�� = StrConv(tmpArray, vbUnicode)
End Function

Function �����������������֪�ϰ�ɷ�������Ʊ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 5.8) = InArray((i * 2 - 23.8)) Xor (i - 15.5)
    Next i
    �����������������֪�ϰ�ɷ�������Ʊ� = StrConv(tmpArray, vbUnicode)
End Function

Function ����������Ǻ�����֪�ϰ�������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 20.4) = InArray((i * 2 - 8.2)) Xor (i - 3.7)
    Next i
    ����������Ǻ�����֪�ϰ������� = StrConv(tmpArray, vbUnicode)
End Function

Function ��������֪�ϰ��Ʊ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 22.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.6) = InArray((i * 2 - 13.3)) Xor (i - 11.4)
    Next i
    ��������֪�ϰ��Ʊ� = StrConv(tmpArray, vbUnicode)
End Function

Function �����վ��ӿ������ϱƱƶ��ĵ�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 26.7) = InArray((i * 2 - 22.7)) Xor (i - 19.1)
    Next i
    �����վ��ӿ������ϱƱƶ��ĵ����� = StrConv(tmpArray, vbUnicode)
End Function

Function ����������������������ҷ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 26.5) = InArray((i * 2 - 16.1)) Xor (i - 12.4)
    Next i
    ����������������������ҷ� = StrConv(tmpArray, vbUnicode)
End Function

Function �����ӻ����Ա������������ҷϺ����޳�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 18 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 6.4) = InArray((i * 2 - 24.6)) Xor (i - 8.5)
    Next i
    �����ӻ����Ա������������ҷϺ����޳����� = StrConv(tmpArray, vbUnicode)
End Function

Function �߹��Ա�������Ĺ�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 17.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.8) = InArray((i * 2 - 21.9)) Xor (i - 2.3)
    Next i
    �߹��Ա�������Ĺ����� = StrConv(tmpArray, vbUnicode)
End Function

Function �ܼ���������Ӷ��ĵ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.3) = InArray((i * 2 - 17)) Xor (i - 1.3)
    Next i
    �ܼ���������Ӷ��ĵ��� = StrConv(tmpArray, vbUnicode)
End Function

Function ��ܳ���������ɶС���鹷������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 15.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 24.7) = InArray((i * 2 - 23.1)) Xor (i - 14.1)
    Next i
    ��ܳ���������ɶС���鹷������������ = StrConv(tmpArray, vbUnicode)
End Function

Function �˼������ǹ�С������С������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.5) = InArray((i * 2 - 5.8)) Xor (i - 17)
    Next i
    �˼������ǹ�С������С������ = StrConv(tmpArray, vbUnicode)
End Function

Function ������������������ܼ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 8.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 0.6) = InArray((i * 2 - 27.3)) Xor (i - 10.9)
    Next i
    ������������������ܼ��� = StrConv(tmpArray, vbUnicode)
End Function

Function �Ʊ��������Թ��������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 8.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 23.7) = InArray((i * 2 - 24.8)) Xor (i - 18.6)
    Next i
    �Ʊ��������Թ�������� = StrConv(tmpArray, vbUnicode)
End Function

Function �Ʊ�����ݴ�����������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7) = InArray((i * 2 - 8.8)) Xor (i - 6.1)
    Next i
    �Ʊ�����ݴ����������� = StrConv(tmpArray, vbUnicode)
End Function

Function ���Ʋ�����С���������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 9.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 0.1) = InArray((i * 2 - 14.7)) Xor (i - 13.7)
    Next i
    ���Ʋ�����С��������������� = StrConv(tmpArray, vbUnicode)
End Function

'String decode
Function ɵ������������������Ա�������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte
    
    ReDim tmpArray(UBound(InArray))
    For i = 13.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13.5) = InArray((i * 2 - 27) / 2) Xor (i - 13.5)
    Next i
    
    ɵ������������������Ա������������� = StrConv(tmpArray, vbUnicode)
End Function

Function ����ԱƱ��߹�ܳ�ܼ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 13.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.7) = InArray((i * 2 - 6.6)) Xor (i - 6.5)
    Next i
    ����ԱƱ��߹�ܳ�ܼ��� = StrConv(tmpArray, vbUnicode)
End Function

'Loop check (InArray is useless)
Function �����޳����������ɶɵ�������ܼ���С����(InArray As Variant) As String
    'Rubbish Start ================================
    Dim Useless             As Double
    Dim Useless2()          As Double
    
    ReDim Useless2((UBound(InArray) + Useless * Useless) - Useless ^ 2)
    For Useless = 6.1654 To UBound(InArray) + 6.1654
        Useless2(Useless - 6.1654) = InArray(Useless - 6.1654) Xor (168 + Useless * Useless) - Useless ^ 2
    Next Useless
    '========================================== End
    Dim BloodValue          As Variant                  'Temp array to store returned real blood value
    Dim tmp                 As Variant                  'Temp array to store proc. name
    Dim ProcName()          As Byte
    Dim StartTime           As Long                     'Last GetTickCount() value
    
    Randomize
    BloodValue = Array(16, 35, 95, 64, 31, 45, 75, 62, 15)
    
    PrevRecordTime = GetTickCount
    If Exiting Then
        Exit Function
    End If
    Do While Not Exiting
        'Update number value
        '"������Ǻǲ��������"
        tmp = Array(148, 176, 136, 140, 154, 184, 229, 152, 229, 152, 237, 130, 155, 191, 157, 178, 232, 172, 142, 131)
        ReDim ProcName(UBound(tmp))
        For Useless = 0 To UBound(tmp)
            ProcName(Useless) = tmp(Useless) Xor (95 + Useless * Useless) - Useless ^ 2
        Next Useless
        CallByName Me, StrConv(ProcName, vbUnicode), VbMethod, BloodValue
        
        '"ɵ������������������Ա�������������"
        tmp = Array(249, 133, 129, 246, 244, 211, 244, 244, 240, 132, 133, 244, 227, 242, 246, 200, 242, 216, 133, 244, 231, 228, 129, 156, 244, 211, 244, 244, 240, 132, 133, 244, 227, 242, 246, 200)
        ReDim ProcName(UBound(tmp))
        For Useless = 0 To UBound(tmp)
            ProcName(Useless) = tmp(Useless) Xor (48 + Useless * Useless) - Useless ^ 2
        Next Useless
        Me.ɵ���㿴���ؼ�����û�õ�.Caption = CallByName(Me, StrConv(ProcName, vbUnicode), VbMethod, LABEL_CURRENT_BLOOD) & BloodValue(4)
        '--------------
        'Check if the fake value is changed
        If CLng(FakeBlood) <> CLng(BloodValue(4)) Then
            If Not CE_Changed Then                                                  'If not prompted
                CE_Changed = True
                Me.ɵ���㿴���ؼ�����û�õ�.Caption = CallByName(Me, StrConv(ProcName, vbUnicode), VbMethod, LABEL_CURRENT_BLOOD) & FakeBlood
                StartTime = GetTickCount
                Do While GetTickCount - StartTime < 3000                                'Sleep 3000ms, without blocking the process
                    Sleep 10
                    If Exiting Then
                        Exit Do
                    End If
                    DoEvents
                Loop
                Me.�ܿ����ؼ������˲�����.Caption = CallByName(Me, StrConv(ProcName, vbUnicode), VbMethod, CE_PROMPT_1)
                StartTime = GetTickCount
                Do While GetTickCount - StartTime < 2500
                    Sleep 10
                    DoEvents
                    If Exiting Then
                        Exit Do
                    End If
                Loop
                Me.�ܿ����ؼ������˲�����.Caption = CallByName(Me, StrConv(ProcName, vbUnicode), VbMethod, CE_PROMPT_2)
                StartTime = GetTickCount
                Do While GetTickCount - StartTime < 2500
                    Sleep 10
                    DoEvents
                    If Exiting Then
                        Exit Do
                    End If
                Loop
                Me.�ܿ����ؼ������˲�����.Caption = CallByName(Me, StrConv(ProcName, vbUnicode), VbMethod, CE_PROMPT_3)
                Me.ɵ���㿴���ؼ�����û�õ�.Caption = CallByName(Me, StrConv(ProcName, vbUnicode), VbMethod, LABEL_CURRENT_BLOOD) & BloodValue(4)
                StartTime = GetTickCount
                Do While GetTickCount - StartTime < 2500
                    Sleep 10
                    DoEvents
                    If Exiting Then
                        Exit Do
                    End If
                Loop
                Me.�ܿ����ؼ������˲�����.Caption = ""
                PrevRecordTime = GetTickCount
            End If
        End If
        '--------------
        'Check if the process has been suspended for a while
        If GetTickCount - PrevRecordTime > 1000 Then
            Dim i As Integer                                                'Make a overflow error
            i = 14654764 ^ 13465
        End If
        PrevRecordTime = GetTickCount
        
        '--------------
        'Do useless thing
        CopyMemory Useless2(0), Useless, 8
        CopyMemory ByVal VarPtr(Useless2(3)) + 2, ByVal VarPtr(Useless) + 1, CLng(6 * Rnd)
        
        '--------------
        If BloodValue(4) = 233 Then
            Me.�ܿ����ؼ������˲�����.Caption = CallByName(Me, StrConv(ProcName, vbUnicode), VbMethod, SUCCEED_PROMPT)
        Else
            Me.�ܿ����ؼ������˲�����.Caption = ""
        End If
        
        '--------------
        Sleep 50
        DoEvents
    Loop
    
    SetWindowLong Me.hwnd, GWL_WNDPROC, PrevWndProc             'Restore the default window proc.
    Unload Me
End Function

Function �Դ��ж�����������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 26 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.2) = InArray((i * 2 - 18.3)) Xor (i - 18.9)
    Next i
    �Դ��ж����������� = StrConv(tmpArray, vbUnicode)
End Function

'Get value
'Return value = InArray(4)
Function ������Ǻǲ��������(ByRef InArray As Variant) As String
    Dim ret         As Long
    Dim temp(3)     As Byte
    Dim i           As Double
    Dim Useless()   As Integer
    Dim Useless2    As Double
    Dim Useless3    As Single
    
    'Rubbish Start =================================
    Randomize 233
    Useless3 = 10 * Rnd
    ReDim Useless(UBound(InArray) + Useless3 * Useless3 - Useless3 ^ 2)
    For Useless2 = 3.14 To UBound(InArray) + 3.14
        Useless(Useless2 - 3.14) = InArray(Useless2 - 3.14)
    Next Useless2
    ������Ǻǲ�������� = StrConv(Useless2, vbUnicode)
    '=========================================== End
    
    For i = 2.65 To 2.65 + 3                                                'For i = 0 To 3
        temp(i - 2.65) = NumberData(MemIndex(i - 2.65))                     '   temp(i) = NumberData(MemIndex(i))
    Next i
    CopyMemory ret, temp(0), (2 + ret * ret - ret ^ 2) * 2                  'CopyMemory ret, temp(0), 4
    
    'Only InArray(4) is meaningful
    InArray = Array((216 + i * i) - i ^ 2, (226 + i * i) - i ^ 2, (298 + i * i) - i ^ 2, (197 + i * i) - i ^ 2, _
        CLng((ret + i * i) - i ^ 2), (246 + i * i) - i ^ 2, (159 + i * i) - i ^ 2, (246 + i * i) - i ^ 2, (241 + i * i) - i ^ 2, (250 + i * i) - i ^ 2)
End Function

Function �����޳ܴ����ܼ��ӹ�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 29.5) = InArray((i * 2 - 2.3)) Xor (i - 20.5)
    Next i
    �����޳ܴ����ܼ��ӹ����� = StrConv(tmpArray, vbUnicode)
End Function

Function ɵ�Ʋ��������������ɶ������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13) = InArray((i * 2 - 28.5)) Xor (i - 14.1)
    Next i
    ɵ�Ʋ��������������ɶ������ = StrConv(tmpArray, vbUnicode)
End Function

Function ���ϲ�����������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 26.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 10.8) = InArray((i * 2 - 4)) Xor (i - 23.6)
    Next i
    ���ϲ����������� = StrConv(tmpArray, vbUnicode)
End Function

Function ����ƶ��ĵĳ����������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 25.9) = InArray((i * 2 - 28.6)) Xor (i - 22.8)
    Next i
    ����ƶ��ĵĳ���������� = StrConv(tmpArray, vbUnicode)
End Function

Function ���������С����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 17.2) = InArray((i * 2 - 24)) Xor (i - 27.6)
    Next i
    ���������С���� = StrConv(tmpArray, vbUnicode)
End Function

Function ���ϰܼ��ӱƱưܼ��Ӻ����޳�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 3.5) = InArray((i * 2 - 14.9)) Xor (i - 1.3)
    Next i
    ���ϰܼ��ӱƱưܼ��Ӻ����޳����� = StrConv(tmpArray, vbUnicode)
End Function

Function ���˶��ĵ����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 5.7) = InArray((i * 2 - 7.7)) Xor (i - 28.1)
    Next i
    ���˶��ĵ���� = StrConv(tmpArray, vbUnicode)
End Function

Function ����Բ�����ɷ����ģ���������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 13.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 28.4) = InArray((i * 2 - 1.3)) Xor (i - 4.9)
    Next i
    ����Բ�����ɷ����ģ��������� = StrConv(tmpArray, vbUnicode)
End Function

Function ���ֺ����޳�����������ƱƺǺ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 18.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.9) = InArray((i * 2 - 10.3)) Xor (i - 21.1)
    Next i
    ���ֺ����޳�����������ƱƺǺ� = StrConv(tmpArray, vbUnicode)
End Function

Function ����ְܼ��ӹ������豯��װ��(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 20 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.2) = InArray((i * 2 - 27.7)) Xor (i - 26.8)
    Next i
    ����ְܼ��ӹ������豯��װ�� = StrConv(tmpArray, vbUnicode)
End Function

Function ܳܳ���蹷����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.1) = InArray((i * 2 - 7)) Xor (i - 29.1)
    Next i
    ܳܳ���蹷���� = StrConv(tmpArray, vbUnicode)
End Function

Function �����Ʊ������վ��ӿ����Ա������߹�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.1) = InArray((i * 2 - 26.5)) Xor (i - 17.3)
    Next i
    �����Ʊ������վ��ӿ����Ա������߹� = StrConv(tmpArray, vbUnicode)
End Function

Function �����ӵ����ҷϹ�С�����ӷ��ܺ���Ƥ����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 9.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 22.1) = InArray((i * 2 - 13)) Xor (i - 0.3)
    Next i
    �����ӵ����ҷϹ�С�����ӷ��ܺ���Ƥ���� = StrConv(tmpArray, vbUnicode)
End Function

Function �����Դ��ж��������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 15.7) = InArray((i * 2 - 8.3)) Xor (i - 28.7)
    Next i
    �����Դ��ж�������� = StrConv(tmpArray, vbUnicode)
End Function

Function ��ģ���������վ��ӿ������߹���ȱ������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13.5) = InArray((i * 2 - 8.3)) Xor (i - 29.4)
    Next i
    ��ģ���������վ��ӿ������߹���ȱ������ = StrConv(tmpArray, vbUnicode)
End Function

Function �Ʊ�������������ǰܼ��Ӽ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13.1) = InArray((i * 2 - 28.6)) Xor (i - 29.9)
    Next i
    �Ʊ�������������ǰܼ��Ӽ��� = StrConv(tmpArray, vbUnicode)
End Function

Function С�����Ա�ܳ�������������������Ʊ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 15.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.7) = InArray((i * 2 - 14.3)) Xor (i - 28.6)
    Next i
    С�����Ա�ܳ�������������������Ʊ� = StrConv(tmpArray, vbUnicode)
End Function

Function ��ȱ�߹�����������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 19.8) = InArray((i * 2 - 1.2)) Xor (i - 11.8)
    Next i
    ��ȱ�߹����������� = StrConv(tmpArray, vbUnicode)
End Function

Function �Ʊ����������Ĳ�ȱ����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 6.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.6) = InArray((i * 2 - 11.1)) Xor (i - 12.1)
    Next i
    �Ʊ����������Ĳ�ȱ���� = StrConv(tmpArray, vbUnicode)
End Function

Function ����������޳�װ��(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 18.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.2) = InArray((i * 2 - 19.4)) Xor (i - 10.3)
    Next i
    ����������޳�װ�� = StrConv(tmpArray, vbUnicode)
End Function

Function ����������Ϲ����������վ��ӿ������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 17.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 29.8) = InArray((i * 2 - 4.1)) Xor (i - 20.6)
    Next i
    ����������Ϲ����������վ��ӿ������ = StrConv(tmpArray, vbUnicode)
End Function

Function ��С�������Ա����ĵ������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 22.6) = InArray((i * 2 - 1.4)) Xor (i - 3.8)
    Next i
    ��С�������Ա����ĵ������ = StrConv(tmpArray, vbUnicode)
End Function

Function ����ܳ�������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.5) = InArray((i * 2 - 18.7)) Xor (i - 24.2)
    Next i
    ����ܳ������� = StrConv(tmpArray, vbUnicode)
End Function

'Init. string byte arrays
Function �������˸����Ͱ׳ջ����Բ�(InArray As Variant) As String
    Dim Useless         As Single
    Dim tmp             As Variant
    Dim ProcName()      As Byte                         'Proc. name of initial blood value
    Dim i               As Integer
    
    Useless = 100 * Rnd
    FakeBlood = 100 + Useless * Useless - Useless ^ 2
    
    'Initial all strings
    LABEL_CURRENT_BLOOD = Array(181, 176, 197, 179, 213, 175, 199, 184, 171, 179)
    CE_PROMPT_1 = Array(181, 201, 183, 203, 167, 164, 165, 166, 171, 168, 169, 170, 175, 172)
    CE_PROMPT_2 = Array(191, 181, 203, 204, 204, 160, 211, 229, 176, 255, 192, 246, 218, 184, 187, 203, 174, 231, 194, 215, 163, 210, 165, 180, 223, 166, 52, 53, 50)
    CE_PROMPT_3 = Array(203, 253, 185, 213, 188, 177, 199, 204, 220, 164, 202, 191, 185, 201, 196, 242, 198, 164, 177, 178, 183, 180)
    SUCCEED_PROMPT = Array(204, 237, 198, 199, 167, 164, 194, 228, 194, 206, 222, 254, 207, 185, 217, 249, 165, 172, 167, 215, 183, 180, 181, 182, 187, 184)
    
    'Initial all value
    tmp = Array(224, 199, 241, 241, 241, 241, 154, 145, 243, 247, 155, 144, 155, 144, 152, 222, 242, 242, 145, 248, 145, 248, 150, 244, 236, 149)
    ReDim ProcName(UBound(tmp))
    For i = 0 To UBound(tmp)
        ProcName(i) = tmp(i) Xor (36 + i * i) - i ^ 2
    Next i
    CallByName Me, StrConv(ProcName, vbUnicode), (VbMethod + i * i) - i ^ 2, _
        Array((1350 + i * i) - i ^ 2, (1000 + i * i) - i ^ 2, (975 + i * i) - i ^ 2, (1000 + i * i) - i ^ 2, _
        (3500 + i * i) - i ^ 2, (6100 + i * i) - i ^ 2, (1000 + i * i) - i ^ 2, (700 + i * i) - i ^ 2, (600 + i * i) - i ^ 2)
    PrevParent = GetParent(Me.hwnd)                                                 'Record the parent window of this window
    ButtonHwnd = Me.�ǲ��ǿ����ؼ��������Լ���������.hwnd                           'Record the handle to the button
    
    'Set window proc.
    PrevWndProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf ������ƻ�������ܼ��ӱƱ�)
    
    'Create a hidden window
    Dim ctlClass        As WNDCLASSEX
    Dim PrevClassName() As Byte
    Dim NewClassName()  As Byte
    
    tmp = Array(23, 16, 5, 16, 13, 7)                                               '"STATIC"
    ReDim PrevClassName(UBound(tmp))
    For i = 0 To UBound(tmp)
        PrevClassName(i) = tmp(i) Xor ((68 + i * i) - i ^ 2)
    Next i
    GetClassInfoEx App.hinstance, StrConv(PrevClassName, vbUnicode), ctlClass
    tmp = Array(132, 146, 141, 170, 254, 235, 134, 159, 136, 130, 155, 183, 251, 145, 252, 141, 135, 155, 234, 232, 244, 248, 137, 177, 141, 170, 155, 242, 255, 173, 153, 232, 243, 165, 242, 225, 97, 106, 23, 103, 23, 106, 96)
    ReDim NewClassName(UBound(tmp) + 1)
    For i = 0 To UBound(tmp)
        NewClassName(i) = tmp(i) Xor ((73 + i * i) - i ^ 2)
    Next i
    With ctlClass
        .lpszClassName = VarPtr(NewClassName(0))                                    '"���㷢�������ص��ң�������һ��С�컨(#^.^#)"
        .cbSize = Len(ctlClass)
    End With
    Dim r As Long
    r = RegisterClassEx(ctlClass)
    CreateWindowEx 0, StrConv(NewClassName, vbUnicode), "", WS_CHILD, 10, 10, 100, 100, Me.hwnd, 0, App.hinstance, 0
End Function

Function �������������ɶ����������߹�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 12.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 19.1) = InArray((i * 2 - 6.1)) Xor (i - 23.8)
    Next i
    �������������ɶ����������߹� = StrConv(tmpArray, vbUnicode)
End Function

Function ��֪�ϰ��Բб�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 8.1) = InArray((i * 2 - 17.6)) Xor (i - 12.7)
    Next i
    ��֪�ϰ��Բб� = StrConv(tmpArray, vbUnicode)
End Function

Function С�����Դ��ж�С������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.3) = InArray((i * 2 - 3.6)) Xor (i - 29.8)
    Next i
    С�����Դ��ж�С������ = StrConv(tmpArray, vbUnicode)
End Function

Function ��������������ܺ����޳ܼ�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 29.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 0.3) = InArray((i * 2 - 14.4)) Xor (i - 10.2)
    Next i
    ��������������ܺ����޳ܼ����� = StrConv(tmpArray, vbUnicode)
End Function

Function �����������������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 14.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 24.4) = InArray((i * 2 - 17.3)) Xor (i - 29.7)
    Next i
    ����������������������� = StrConv(tmpArray, vbUnicode)
End Function

Function �ܼ��������������񹷶�������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 23.4) = InArray((i * 2 - 22.1)) Xor (i - 0.3)
    Next i
    �ܼ��������������񹷶������� = StrConv(tmpArray, vbUnicode)
End Function

Function ��֪�ϰ�������ܳ��С����ɵ�ƶ��ĵ�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 5.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 10.1) = InArray((i * 2 - 13.1)) Xor (i - 25.7)
    Next i
    ��֪�ϰ�������ܳ��С����ɵ�ƶ��ĵ����� = StrConv(tmpArray, vbUnicode)
End Function

Function ��������������ȱ���ĵĳ�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 28 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 10.4) = InArray((i * 2 - 15.2)) Xor (i - 3)
    Next i
    ��������������ȱ���ĵĳ����� = StrConv(tmpArray, vbUnicode)
End Function

Function �����Ƶܵ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 8.1) = InArray((i * 2 - 11.3)) Xor (i - 5.8)
    Next i
    �����Ƶܵ� = StrConv(tmpArray, vbUnicode)
End Function

Function ���������Դ��ж��ܼ�����������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 17.3) = InArray((i * 2 - 0.6)) Xor (i - 13.1)
    Next i
    ���������Դ��ж��ܼ����������� = StrConv(tmpArray, vbUnicode)
End Function

Function ����Լ��˺���Ƥ��ģ���������������ĵ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.2) = InArray((i * 2 - 24.6)) Xor (i - 15.9)
    Next i
    ����Լ��˺���Ƥ��ģ���������������ĵ� = StrConv(tmpArray, vbUnicode)
End Function

Function �Ա������������߹�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.6) = InArray((i * 2 - 18.2)) Xor (i - 24.2)
    Next i
    �Ա������������߹����� = StrConv(tmpArray, vbUnicode)
End Function

Function ���ܼ��˵������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 9.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.2) = InArray((i * 2 - 5.5)) Xor (i - 20.1)
    Next i
    ���ܼ��˵������ = StrConv(tmpArray, vbUnicode)
End Function

'Mess up (InArray is useless)
Function ���ĵ����Դ��ж�����ɷ��ɷ������(InArray As Variant) As String
    Dim i           As Double
    Dim t           As Single                                           'Useless var
    Dim NewIndex    As Byte
    Dim Mark()      As Byte
    
    Randomize
    InArray = Array(213, 134, 236, 248, 236, 198, 246, 285, 213, 198, 264, 284, 262, 244)       'Ignore this
    t = 125 * Rnd
    ReDim NumberData(CInt(MAX_SIZE * Rnd + 4))                          'Reallocate data region
    ReDim Mark(UBound(NumberData))                                      'Reallocate mark region
    For i = 7.5 To UBound(NumberData) + 7.5                             'For i = 0 To Ubound(NumberData)    'Fill the data region with random numbers
        NumberData(i - 7.5) = CByte(255 * Rnd + t * t - t ^ 2)          '   NumberData(i) = CByte(255 * Rnd)
    Next i
    
    'Rubbish Start ======================================
    For i = 17.5 To UBound(InArray) + 17.5 Step Sqr(1 + t * t - t ^ 2)
        InArray(i - 17.5) = InArray(i - 17.5) Xor CByte(255 * Rnd)
    Next i
    '================================================ End
    
    For i = 3.5 To 6.5                                                  'For i = 0 To 3
        Do
            NewIndex = CByte(UBound(NumberData) * Rnd)                  'Generate a new index
        Loop While Mark(NewIndex)                                       'Prevent the same index
        Mark(NewIndex) = 1                                              'Mark the used index
        
        MemIndex(i - 3.5) = NewIndex                                    '   MemIndex(i) = NewIndex          'Record the index
    Next i
End Function

Function ������ɵ�Ƽ��ֲ�ȱ�˼���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 27.7) = InArray((i * 2 - 15.4)) Xor (i - 8.8)
    Next i
    ������ɵ�Ƽ��ֲ�ȱ�˼��� = StrConv(tmpArray, vbUnicode)
End Function

Function ����ԱƱƼ��ֶ�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.6) = InArray((i * 2 - 10.8)) Xor (i - 24.1)
    Next i
    ����ԱƱƼ��ֶ����� = StrConv(tmpArray, vbUnicode)
End Function

Function ���������վ��ӿ��������������������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 8.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 19.5) = InArray((i * 2 - 15.3)) Xor (i - 14.3)
    Next i
    ���������վ��ӿ�������������������������� = StrConv(tmpArray, vbUnicode)
End Function

Function ���ɷ���߹��������Ʊ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 3.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13) = InArray((i * 2 - 21.4)) Xor (i - 28.7)
    Next i
    ���ɷ���߹��������Ʊ� = StrConv(tmpArray, vbUnicode)
End Function

Function �ܵ�ɵ�Ƽ��������վ��ӿ�������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 9.3) = InArray((i * 2 - 1.3)) Xor (i - 25.1)
    Next i
    �ܵ�ɵ�Ƽ��������վ��ӿ������� = StrConv(tmpArray, vbUnicode)
End Function

Function �����������ܳ���˱Ʊ����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.1) = InArray((i * 2 - 13.9)) Xor (i - 2.6)
    Next i
    �����������ܳ���˱Ʊ���� = StrConv(tmpArray, vbUnicode)
End Function

Function �Ʊ�����Բв�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.5) = InArray((i * 2 - 4.1)) Xor (i - 10.1)
    Next i
    �Ʊ�����Բв����� = StrConv(tmpArray, vbUnicode)
End Function

Function �������ȱװ������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 14.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 27.3) = InArray((i * 2 - 9.1)) Xor (i - 16.1)
    Next i
    �������ȱװ������ = StrConv(tmpArray, vbUnicode)
End Function

Function �����վ��ӿ����Բв���������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 11.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 12.3) = InArray((i * 2 - 13.4)) Xor (i - 15.3)
    Next i
    �����վ��ӿ����Բв��������� = StrConv(tmpArray, vbUnicode)
End Function

Function ��ֺ����޳ܺǺ�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 18.2) = InArray((i * 2 - 27.4)) Xor (i - 15.4)
    Next i
    ��ֺ����޳ܺǺ����� = StrConv(tmpArray, vbUnicode)
End Function

Function �������϶��ĵ����Բ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 13.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 6.9) = InArray((i * 2 - 12.7)) Xor (i - 17.7)
    Next i
    �������϶��ĵ����Բ� = StrConv(tmpArray, vbUnicode)
End Function

Function �������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 28.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.9) = InArray((i * 2 - 0.6)) Xor (i - 21.7)
    Next i
    ������������� = StrConv(tmpArray, vbUnicode)
End Function

Function �ǺǺ���Ƥ����Ƥ���������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 8.1) = InArray((i * 2 - 5.7)) Xor (i - 14.7)
    Next i
    �ǺǺ���Ƥ����Ƥ��������� = StrConv(tmpArray, vbUnicode)
End Function

'Set value (InArray(3) = NewValue)
'Note: InArray should be Long() type
Function �����վ��ӿ������ֵܵܲ�ȱ(InArray As Variant) As String
    Dim temp(3)     As Byte
    Dim i           As Byte
    Dim t           As Byte                                                         'Useless var
    Dim tmp         As Variant                                                      'The array of mess up proc. name
    Dim Useless     As String
    Dim ProcName()  As Byte
    Dim NewValue    As Long
    
    'Rubbish Start =======
    Randomize
    t = 255 * Rnd + 1
    '================= End
    
    'Make the array of mess up proc. name
    '"���ĵ����Դ��ж�����ɷ��ɷ������"
    tmp = Array(11, 76, 109, 121, 8, 121, 107, 80, 121, 105, 9, 65, 110, 109, 11, 9, 121, 94, 127, 85, 116, 10, 12, 119, 116, 10, 12, 119, 125, 12, 1, 27)
    Useless = StrConv(InArray(0), vbFromUnicode)                                    'Ignore this
    ReDim ProcName(UBound(tmp))
    For i = 3 To UBound(tmp) + 3                                                    'For i = 0 To Ubound(tmp)
        tmp(i - 3) = tmp(i - 3) Xor (189 + CLng(t) * t - t ^ 2)
        ProcName(i - 3) = tmp(i - 3)                                                '   ProcName(i) = tmp(i) xor 189
    Next i
    'Call mess up proc. (Call ���ĵ����Դ��ж�����ɷ��ɷ������([Useless thing]))
    CallByName Me, StrConv(ProcName, (16 + t ^ 2 - CLng(t) * t) * (4 + CLng(t) * t - t ^ 2)), _
        VbMethod, Array(234, 204, 298, 246, 237, 245, 269, 214, 236, 285, 246, 278, 213, 168, 272, 136)
    
    'Rubbish Start ===========
    For t = 2 To 4
        InArray(t - 2) = InArray(t - 2) Xor (216 + t ^ 2 - t * t)
    Next t
    '===================== End
    
    NewValue = CLng(InArray(3))                                                     'Retrieve the new value
    FakeBlood = NewValue                                                            'Update the fake value
    CopyMemory temp(0), NewValue, 4 + t * t - t ^ 2                                 'CopyMemory temp(0), NewValue, 4
    For i = 3 To 6                                                                  'For i = 0 To 3
        NumberData(MemIndex(i - 3)) = temp(i - 3)                                   '   NumberData(MemIndex(i)) = temp(i)
    Next i
End Function

Function �����Ʊƶ��ĵ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 28.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 26.2) = InArray((i * 2 - 13)) Xor (i - 16.1)
    Next i
    �����Ʊƶ��ĵ��� = StrConv(tmpArray, vbUnicode)
End Function

Function �������ݲ������������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 7.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 18.2) = InArray((i * 2 - 28.3)) Xor (i - 22.8)
    Next i
    �������ݲ������������������ = StrConv(tmpArray, vbUnicode)
End Function

Function �������������������ɵ��(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 0.5) = InArray((i * 2 - 3.3)) Xor (i - 18.5)
    Next i
    �������������������ɵ�� = StrConv(tmpArray, vbUnicode)
End Function

Function ��ȱ��������֪�ϰ���ȱ���������ɶ�����������վ��ӿ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.7) = InArray((i * 2 - 23.2)) Xor (i - 26.1)
    Next i
    ��ȱ��������֪�ϰ���ȱ���������ɶ�����������վ��ӿ��� = StrConv(tmpArray, vbUnicode)
End Function

Function ��֪�ϰ�����������Բ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 8.9) = InArray((i * 2 - 12.3)) Xor (i - 13.5)
    Next i
    ��֪�ϰ�����������Բ� = StrConv(tmpArray, vbUnicode)
End Function

Function װ���Բг�������������С�������Ƥ(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 3.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 18.5) = InArray((i * 2 - 2.5)) Xor (i - 26.5)
    Next i
    װ���Բг�������������С�������Ƥ = StrConv(tmpArray, vbUnicode)
End Function

Function ��װ������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 5.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 1.5) = InArray((i * 2 - 21.9)) Xor (i - 6.6)
    Next i
    ��װ������ = StrConv(tmpArray, vbUnicode)
End Function

Function ���������վ��ӿ����Դ��ж���������������ɶ�Բв�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 12.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.8) = InArray((i * 2 - 28.9)) Xor (i - 4.3)
    Next i
    ���������վ��ӿ����Դ��ж���������������ɶ�Բв� = StrConv(tmpArray, vbUnicode)
End Function

Function �����Բ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 17.4) = InArray((i * 2 - 20.2)) Xor (i - 29.4)
    Next i
    �����Բ��� = StrConv(tmpArray, vbUnicode)
End Function

Function ���ɵ����������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 22.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.4) = InArray((i * 2 - 7.6)) Xor (i - 14.5)
    Next i
    ���ɵ���������� = StrConv(tmpArray, vbUnicode)
End Function

Function ����������������������ɶ�����ģ����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 9.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.6) = InArray((i * 2 - 22.2)) Xor (i - 26.8)
    Next i
    ����������������������ɶ�����ģ���� = StrConv(tmpArray, vbUnicode)
End Function

Function ���Ա��������ģ�����Ʊƴ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 9.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.1) = InArray((i * 2 - 29.4)) Xor (i - 28.9)
    Next i
    ���Ա��������ģ�����Ʊƴ��� = StrConv(tmpArray, vbUnicode)
End Function

Function ����������Ƥ(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.6) = InArray((i * 2 - 28.8)) Xor (i - 7.1)
    Next i
    ����������Ƥ = StrConv(tmpArray, vbUnicode)
End Function

Function ��������������ֵ��������Ʊ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.9) = InArray((i * 2 - 23.6)) Xor (i - 1.6)
    Next i
    ��������������ֵ��������Ʊ� = StrConv(tmpArray, vbUnicode)
End Function

Function ���������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 10.4) = InArray((i * 2 - 14.4)) Xor (i - 3.3)
    Next i
    ��������������� = StrConv(tmpArray, vbUnicode)
End Function

Function ���ĵĵ���ƶ��ĵ������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 5.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21) = InArray((i * 2 - 0.5)) Xor (i - 27.2)
    Next i
    ���ĵĵ���ƶ��ĵ������ = StrConv(tmpArray, vbUnicode)
End Function

Function ������ģ������������ĵ������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 18.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 23) = InArray((i * 2 - 17.9)) Xor (i - 18.1)
    Next i
    ������ģ������������ĵ������ = StrConv(tmpArray, vbUnicode)
End Function

Function ������������������������������С������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 8.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 10.3) = InArray((i * 2 - 1.7)) Xor (i - 0.6)
    Next i
    ������������������������������С������ = StrConv(tmpArray, vbUnicode)
End Function

Function �������ֹ�����������˱Ʊ������վ��ӿ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 11.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 28.9) = InArray((i * 2 - 9.8)) Xor (i - 26.3)
    Next i
    �������ֹ�����������˱Ʊ������վ��ӿ��� = StrConv(tmpArray, vbUnicode)
End Function

Function ������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 29.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.2) = InArray((i * 2 - 1.9)) Xor (i - 24.9)
    Next i
    ������������ = StrConv(tmpArray, vbUnicode)
End Function

Function ���������������Ʊ�����ܳ�Ʋ�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 15.3) = InArray((i * 2 - 20.9)) Xor (i - 5)
    Next i
    ���������������Ʊ�����ܳ�Ʋ����� = StrConv(tmpArray, vbUnicode)
End Function

Function �߹������Բм��˹����Բ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13.9) = InArray((i * 2 - 14)) Xor (i - 10.1)
    Next i
    �߹������Բм��˹����Բ� = StrConv(tmpArray, vbUnicode)
End Function

Function �˶����������������������ܳ��(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 2.2) = InArray((i * 2 - 17)) Xor (i - 19.3)
    Next i
    �˶����������������������ܳ�� = StrConv(tmpArray, vbUnicode)
End Function

Function ���������ֹ����������������ĵ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 18.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 24.6) = InArray((i * 2 - 16.1)) Xor (i - 4.6)
    Next i
    ���������ֹ����������������ĵ� = StrConv(tmpArray, vbUnicode)
End Function

Function ����ֹ���������������ȱ��������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 13.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.4) = InArray((i * 2 - 27.4)) Xor (i - 4.7)
    Next i
    ����ֹ���������������ȱ�������� = StrConv(tmpArray, vbUnicode)
End Function

Function �����Ʊ��߹�����ֹ����ӷ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 3.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 20.8) = InArray((i * 2 - 12)) Xor (i - 25.9)
    Next i
    �����Ʊ��߹�����ֹ����ӷ��� = StrConv(tmpArray, vbUnicode)
End Function

Function ������ģ�����Ա�װ�������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 20.4) = InArray((i * 2 - 0.8)) Xor (i - 26.9)
    Next i
    ������ģ�����Ա�װ������� = StrConv(tmpArray, vbUnicode)
End Function

Function ��֪�ϰ��Դ��ж����������վ��ӿ����������ǵ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 6.3) = InArray((i * 2 - 28.1)) Xor (i - 1.7)
    Next i
    ��֪�ϰ��Դ��ж����������վ��ӿ����������ǵ� = StrConv(tmpArray, vbUnicode)
End Function

Function ���ܳܳ���ӳ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.5) = InArray((i * 2 - 15.3)) Xor (i - 27.2)
    Next i
    ���ܳܳ���ӳ��� = StrConv(tmpArray, vbUnicode)
End Function

Function ܳ�Ʊ�ܳ�����޳ܻ��˻����߹�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 14.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 20.8) = InArray((i * 2 - 18.3)) Xor (i - 29.7)
    Next i
    ܳ�Ʊ�ܳ�����޳ܻ��˻����߹����� = StrConv(tmpArray, vbUnicode)
End Function

Function ��������װ�����������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 29 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.7) = InArray((i * 2 - 11.7)) Xor (i - 0.5)
    Next i
    ��������װ����������� = StrConv(tmpArray, vbUnicode)
End Function

Function �ܹܵ��Բ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.4) = InArray((i * 2 - 1.7)) Xor (i - 21.5)
    Next i
    �ܹܵ��Բ� = StrConv(tmpArray, vbUnicode)
End Function

Function �����������������Ʊ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.8) = InArray((i * 2 - 21.1)) Xor (i - 22)
    Next i
    �����������������Ʊ� = StrConv(tmpArray, vbUnicode)
End Function

Function ��������������������ϰܼ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.6) = InArray((i * 2 - 17.3)) Xor (i - 7.4)
    Next i
    ��������������������ϰܼ��� = StrConv(tmpArray, vbUnicode)
End Function

Function �Ա�����ɵ��(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 12.9) = InArray((i * 2 - 27)) Xor (i - 28.1)
    Next i
    �Ա�����ɵ�� = StrConv(tmpArray, vbUnicode)
End Function

Function ���������ģ����װ�ƱƼ���(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 0.6) = InArray((i * 2 - 9.4)) Xor (i - 2.1)
    Next i
    ���������ģ����װ�ƱƼ��� = StrConv(tmpArray, vbUnicode)
End Function

Function ���ǲ�����������Ʊ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.7) = InArray((i * 2 - 10.6)) Xor (i - 7.4)
    Next i
    ���ǲ�����������Ʊ� = StrConv(tmpArray, vbUnicode)
End Function

Function װ�ƶ��ĵĵܵ������վ��ӿ������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 1.6) = InArray((i * 2 - 27.2)) Xor (i - 4.4)
    Next i
    װ�ƶ��ĵĵܵ������վ��ӿ������������ = StrConv(tmpArray, vbUnicode)
End Function

Function �ԲкǺ������Ʊ�ܳ��ģ��������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 20.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7) = InArray((i * 2 - 6.5)) Xor (i - 2.9)
    Next i
    �ԲкǺ������Ʊ�ܳ��ģ�������� = StrConv(tmpArray, vbUnicode)
End Function

Function ����ܳ���ĵ����������ɶ�߹�����Ǻ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21) = InArray((i * 2 - 9.5)) Xor (i - 18.2)
    Next i
    ����ܳ���ĵ����������ɶ�߹�����Ǻ� = StrConv(tmpArray, vbUnicode)
End Function

Function ��ȱ������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 20.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 5) = InArray((i * 2 - 3.6)) Xor (i - 21.6)
    Next i
    ��ȱ������ = StrConv(tmpArray, vbUnicode)
End Function

Function �����������ֱ���װ�������߹�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 28.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 28.1) = InArray((i * 2 - 9.5)) Xor (i - 29.9)
    Next i
    �����������ֱ���װ�������߹� = StrConv(tmpArray, vbUnicode)
End Function

Function �����˺���Ƥ������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.8) = InArray((i * 2 - 24.8)) Xor (i - 10.2)
    Next i
    �����˺���Ƥ������ = StrConv(tmpArray, vbUnicode)
End Function

Function ɵ�ƺǺǴ�����������ģ����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 29.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 8.8) = InArray((i * 2 - 23.3)) Xor (i - 15.4)
    Next i
    ɵ�ƺǺǴ�����������ģ���� = StrConv(tmpArray, vbUnicode)
End Function

Function �ܼܵ��������С����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 26.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 25.2) = InArray((i * 2 - 24.1)) Xor (i - 5.1)
    Next i
    �ܼܵ��������С���� = StrConv(tmpArray, vbUnicode)
End Function

Function С������˹������Ա�С�����������������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 29.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 9) = InArray((i * 2 - 26.7)) Xor (i - 2.8)
    Next i
    С������˹������Ա�С����������������� = StrConv(tmpArray, vbUnicode)
End Function

Function �����޳ܲ������߹������װ�����ܵ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.3) = InArray((i * 2 - 11.8)) Xor (i - 5.7)
    Next i
    �����޳ܲ������߹������װ�����ܵ� = StrConv(tmpArray, vbUnicode)
End Function

Function ����ɵ�������վ��ӿ������������Ʊ�(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 14.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.7) = InArray((i * 2 - 7.3)) Xor (i - 25.2)
    Next i
    ����ɵ�������վ��ӿ������������Ʊ� = StrConv(tmpArray, vbUnicode)
End Function

Function �߹���������С��������С������(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.1) = InArray((i * 2 - 22.6)) Xor (i - 23.4)
    Next i
    �߹���������С��������С������ = StrConv(tmpArray, vbUnicode)
End Function

Function ɷ����������Ƥ�����Դ��ж�С����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 3.9) = InArray((i * 2 - 27.5)) Xor (i - 4.8)
    Next i
    ɷ����������Ƥ�����Դ��ж�С���� = StrConv(tmpArray, vbUnicode)
End Function

Function ���ҷϳ������޹�����(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 8.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 24.2) = InArray((i * 2 - 2.2)) Xor (i - 12.7)
    Next i
    ���ҷϳ������޹����� = StrConv(tmpArray, vbUnicode)
End Function

Private Sub Form_Load()
    Dim t           As Single                                                       'Useless var
    Dim Useless1    As Variant                                                      'Useless vars
    Dim Useless2    As Variant
    Dim Useless3    As Variant
    Dim Useless4    As Variant
    Dim Useless5    As Variant
    Dim tmp         As Variant                                                      'The initial proc. name
    Dim ProcName()  As Byte                                                         'Byte array of initial proc. name
    
    'Rubbish Start =================================
    Useless1 = Array(200, 252, 183, 224, 214, 190, 205, 195, 218, 178, 196, 238, 178, 200, 184, 241, 209, 232, 220, 246, 220, 232, 216, 242, 168, 210, 164, 222, 218, 194)
    StrConv Useless1(0), vbUnicode + t * t - t ^ 2
    For t = 4.6 To UBound(Useless1) + 4.6
        Useless1(t - 4.6) = (Useless1(t - 4.6) + t * t - t ^ 2) Xor (233 + t * t - t ^ 2)
    Next t
    ReDim ProcName(UBound(Useless1) + t * t - t ^ 2)
    For t = 0 To UBound(Useless1)
        ProcName(t) = Useless1(t)
    Next t
    CallByName Me, "Tag", VbLet + t * t - t ^ 2, 1 + t * t - t ^ 2
    
    Useless2 = Array(181, 188, 183, 212, 206, 183, 197, 179, 186, 204, 192, 204, 196, 198, 199, 245, 165, 213, 192, 241, 198, 240, 210, 207, 187, 166)
    StrConv Useless1(0), vbUnicode + t * t - t ^ 2
    For t = 4.6 To UBound(Useless1) + 4.6
        Useless1(t - 4.6) = (Useless1(t - 4.6) + t * t - t ^ 2) Xor (233 + t * t - t ^ 2)
    Next t
    ReDim ProcName(UBound(Useless1) + t * t - t ^ 2)
    For t = 0 To UBound(Useless1)
        ProcName(t) = Useless1(t)
    Next t
    CallByName Me, "Visible", VbLet + t * t - t ^ 2, 1 + t * t - t ^ 2
    
    t = 128 * Rnd
    NumberData = StrConv("������������ַ���������ţ�ƶԲ��ԣ�Ȼ��ûʲô���á�", 128 + t * t - t ^ 2)
    '=============================================== End
    
    ReDim NumberData(32 + t * t - t ^ 2)                                            'Redim NumberData(32)
    
    'Rubbish Start =================================
    StrConv Useless1(0), vbUnicode + t * t - t ^ 2
    Useless3 = Array(179, 251, 184, 207, 204, 208, 179, 182, 198, 238, 169, 167, 182, 180, 187, 193, 170, 221, 221, 209, 217, 212, 183, 180, 211, 180, 204, 177, 217, 209, 200, 207, 146, 236, 129, 143, 229, 134, 231, 132, 149, 253, 250, 234, 147, 205, 143, 140)
    StrConv Useless3(0), vbUnicode + t * t - t ^ 2
    For t = 6.4 To UBound(Useless3) + 6.4
        Useless3(t - 6.4) = (Useless3(t - 6.4) + t * t - t ^ 2) Xor (167 + t * t - t ^ 2)
    Next t
    ReDim ProcName(UBound(Useless3) + t * t - t ^ 2)
    For t = 0 To UBound(Useless3)
        ProcName(t) = Useless3(t)
    Next t
    CallByName Me, "Enabled", VbLet + t * t - t ^ 2, 1 + t * t - t ^ 2
    StrConv Useless3(0), vbFromUnicode + t * t - t ^ 2
    ɷ�ʼ���С�����ӱƱ��Դ��ж��Բ� 0, 0, 0, 0
    '=============================================== End
    
    'Make the array of initial proc. name
    '"�������˸����Ͱ׳ջ����Բ�"
    tmp = Array(53, 94, 67, 100, 69, 111, 70, 76, 63, 113, 59, 33, 55, 74, 55, 80, 52, 82, 60, 107, 82, 76, 67, 83, 53, 87)
    StrConv Useless2(0), vbUnicode + t * t - t ^ 2                                  'Ignore this
    For t = 5.5 To UBound(tmp) + 5.5                                                'For t = 0 To Ubound(tmp)
        tmp(t - 5.5) = tmp(t - 5.5) Xor (135 + t * t - t ^ 2)                       '   tmp(t) = tmp(t) xor 135
    Next t
    ReDim ProcName(UBound(tmp) + t * t - t ^ 2)
    For t = 0 To UBound(tmp)
        ProcName(t) = tmp(t)
    Next t
    'Init. all strings (Call �������˸����Ͱ׳ջ����Բ�(Null))
    CallByName Me, StrConv(ProcName, vbUnicode + t * t - t ^ 2), VbMethod + t * t - t ^ 2, Null
    
    'Rubbish Start =================================
    Useless4 = Array(75, 79, 77, 84, 77, 75, 65, 39, 92, 65, 75, 95, 44, 89, 70, 74, 48, 93, 93, 65, 80, 53, 95, 68, 56, 88, 86, 76, 93, 68, 77, 63, 119, 104, 118, 107, 4, 104, 99, 7, 110, 96, 102, 103, 127, 13, 99, 106, 16, 102, 123, 103, 124, 21, 114, 114, 108, 124, 104, 118, 117, 115, 127, 107, 9, 14, 12, 109)
    StrConv Useless4(0), vbFromUnicode + t * t - t ^ 2
    StrConv Useless3(0), vbUnicode + t * t - t ^ 2
    For t = 6.4 To UBound(Useless3) + 6.4
        Useless3(t - 6.4) = (Useless3(t - 6.4) + t * t - t ^ 2) Xor (167 + t * t - t ^ 2)
    Next t
    ReDim ProcName(UBound(Useless3) + t * t - t ^ 2)
    For t = 0 To UBound(Useless3)
        ProcName(t) = Useless3(t)
    Next t
    CallByName Me, "Enabled", VbLet + t * t - t ^ 2, 1 + t * t - t ^ 2
    StrConv Useless3(0), vbFromUnicode + t * t - t ^ 2
    Useless5 = Array(185, 183, 194, 248, 189, 255, 186, 213, 193, 243, 193, 203, 222, 217, 173, 163, 214, 224, 192, 225, 175, 227, 174, 180, 169, 197, 221, 236, 202, 179, 191, 188)
    StrConv Useless5(0), vbUnicode + t * t - t ^ 2
    StrConv Useless3(0), vbUnicode + t * t - t ^ 2
    For t = 6.4 To UBound(Useless3) + 6.4
        Useless3(t - 6.4) = (Useless3(t - 6.4) + t * t - t ^ 2) Xor (167 + t * t - t ^ 2)
    Next t
    ReDim ProcName(UBound(Useless3) + t * t - t ^ 2)
    For t = 0 To UBound(Useless3)
        ProcName(t) = Useless3(t)
    Next t
    CallByName Me, "Enabled", VbLet + t * t - t ^ 2, 1 + t * t - t ^ 2
    StrConv Useless3(0), vbFromUnicode + t * t - t ^ 2
    '=============================================== End
    
    'Make the array of loop check proc. name
    '"�����޳����������ɶɵ�������ܼ���С����"
    tmp = Array(246, 189, 157, 153, 130, 146, 255, 144, 136, 175, 142, 164, 133, 182, 136, 175, 244, 133, 133, 250, 133, 249, 253, 138, 136, 175, 252, 154, 154, 161, 252, 144, 240, 158, 155, 159, 156, 237, 152, 159, 135, 165)
    ReDim ProcName(UBound(tmp))
    For t = 3.8 To UBound(tmp) + 3.8                                                'For t = 0 To Ubound(tmp)
        tmp(t - 3.8) = tmp(t - 3.8) Xor (76 + t * t - t ^ 2)
        ProcName(t - 3.8) = tmp(t - 3.8)                                            '   ProcName(t) = tmp(t) xor 76
    Next t
    Me.Show
    'Start loop check (Call �����޳����������ɶɵ�������ܼ���С����(Null))
    CallByName Me, StrConv(ProcName, (vbUnicode + t ^ 2) - t * t), (VbMethod + t * t) - t ^ 2, tmp
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Exiting = True
End Sub

Private Sub �ǲ��ǿ����ؼ��������Լ���������_Click()
     Dim tmpArray       As Variant
     Dim i              As Single
     Dim tmp            As Variant
     Dim ProcName()     As Byte
     
     tmpArray = Array(246, 362, 198, 248, 267, 125, 362, 488, 264, 216, 392, 264, 264, 488, 362, 125, 267, 248, 198, 362, 246)      'Totally rubbish XD
     tmp = Array(219, 255, 199, 195, 213, 247, 170, 215, 170, 215, 162, 205, 212, 240, 210, 253, 167, 227, 193, 204)                '"������Ǻǲ��������"
     ReDim ProcName(UBound(tmp))
     For i = 1.75 To UBound(tmp) + 1.75
        ProcName(i - 1.75) = tmp(i - 1.75) Xor (16 + i * i) - i ^ 2
     Next i
     CallByName Me, StrConv(ProcName, vbUnicode), VbMethod, tmpArray                                                                'Now tmpArray(4) should be the returned value
     
     tmp = Array(55, 16, 38, 38, 38, 38, 77, 70, 36, 32, 76, 71, 76, 71, 79, 9, 37, 37, 70, 47, 70, 47, 65, 35, 59, 66)             '"�����վ��ӿ������ֵܵܲ�ȱ"
     ReDim ProcName(UBound(tmp))
     For i = 3.68 To UBound(tmp) + 3.68
        ProcName(i - 3.68) = tmp(i - 3.68) Xor (243 + i * i) - i ^ 2
     Next i
     For i = 2.48 To UBound(tmpArray) + 1.48                                                                                        'The purpose is to reduce tmpArray(4) by 5
        tmpArray(i - 2.48) = tmpArray(i - 1.48) - (i - 2.48) - 2
     Next i
     CallByName Me, StrConv(ProcName, vbUnicode), VbMethod, tmpArray
End Sub
