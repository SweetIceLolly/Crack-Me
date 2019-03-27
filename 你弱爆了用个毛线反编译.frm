VERSION 5.00
Begin VB.Form 你弱爆了用个毛线反编译 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "来破解呀"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton 是不是看到控件名觉得自己很厉害呢 
      Caption         =   "-5"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.Label 能看到控件名很了不起吗 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label 傻逼你看到控件名是没用的 
      AutoSize        =   -1  'True
      Caption         =   "当前血量：1000"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1260
   End
   Begin VB.Label 哈哈哈关掉反编译吧菜鸡 
      AutoSize        =   -1  'True
      Caption         =   "把你的血量改成233就成功了。"
      Height          =   195
      Left            =   652
      TabIndex        =   0
      Top             =   960
      Width           =   2430
   End
End
Attribute VB_Name = "你弱爆了用个毛线反编译"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'String decode: 傻逼你哪来的勇气妈的自爆你哪来的勇气()
'Init. vars:    操你妈了个鸡巴白痴混账脑残()
'Loop check:    厚颜无耻你妈生你干啥傻逼你爸猪败家子小杂碎()
'Set value:     你照照镜子看看贱种弟弟残缺()
'Get value:     孙子喷呵呵草泥马敷衍()
'Mess up:       恶心的猪脑袋有洞你妈煞笔煞笔辣鸡()

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
Private Declare Function 你能看到我声明的API又怎样呢 Lib "user32" () As Long

Private Const MAX_SIZE = 16                     'The max size of real blood value data region

Dim NumberData()            As Byte             'Real blood value data
Dim MemIndex(3)             As Byte             'Memory index of real blood value data
Dim FakeBlood               As Long             'Fake blood value data
Dim PrevParent              As Long             'Initial parent window of this window
Dim ButtonHwnd              As Long             'The handle to the button
Dim Exiting                 As Boolean          'Is program exiting

Dim CE_Changed              As Boolean          'User changed the fake value by CE

Dim LABEL_CURRENT_BLOOD     As Variant          '"当前血量："
Dim CE_PROMPT_1             As Variant          '"等等！！！！！"
Dim CE_PROMPT_2             As Variant          '"看上去这个数值的决心非常强..."
Dim CE_PROMPT_3             As Variant          '"它恢复了原来的数值！！"
Dim SUCCEED_PROMPT          As Variant          '"天哪！你是怎么做到的！！！"

Function 二逼你妈生你干啥厚颜无耻虚无混账(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 6.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 26.7) = InArray((i * 2 - 22.7)) Xor (i - 16)
    Next i
    二逼你妈生你干啥厚颜无耻虚无混账 = StrConv(tmpArray, vbUnicode)
End Function

Function 操你爸敷衍你哪来的勇气你你照照镜子看看二逼弱智(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.6) = InArray((i * 2 - 25.2)) Xor (i - 12.9)
    Next i
    操你爸敷衍你哪来的勇气你你照照镜子看看二逼弱智 = StrConv(tmpArray, vbUnicode)
End Function

Function 脑袋有洞辣鸡你爸蠢货贱人(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 17.7) = InArray((i * 2 - 25.8)) Xor (i - 22)
    Next i
    脑袋有洞辣鸡你爸蠢货贱人 = StrConv(tmpArray, vbUnicode)
End Function

Function 走狗厚颜无耻猪狗儿子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 8.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 12.3) = InArray((i * 2 - 3.8)) Xor (i - 16.7)
    Next i
    走狗厚颜无耻猪狗儿子 = StrConv(tmpArray, vbUnicode)
End Function

Function 认知障碍走狗认知障碍逼逼厌恶(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 8.8) = InArray((i * 2 - 15)) Xor (i - 3)
    Next i
    认知障碍走狗认知障碍逼逼厌恶 = StrConv(tmpArray, vbUnicode)
End Function

Function 逼逼小兔崽子辣鸡狗蠢狗(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 6.1) = InArray((i * 2 - 6.1)) Xor (i - 11.1)
    Next i
    逼逼小兔崽子辣鸡狗蠢狗 = StrConv(tmpArray, vbUnicode)
End Function

Function 草你妈你照照镜子看看艹败家子猪装逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.5) = InArray((i * 2 - 12.2)) Xor (i - 17.3)
    Next i
    草你妈你照照镜子看看艹败家子猪装逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 你呵呵厚脸皮杂种你照照镜子看看蠢货(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 26 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 27.6) = InArray((i * 2 - 1.9)) Xor (i - 17.8)
    Next i
    你呵呵厚脸皮杂种你照照镜子看看蠢货 = StrConv(tmpArray, vbUnicode)
End Function

Function 算个鸟猪算个鸟(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 15.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 12.7) = InArray((i * 2 - 13.4)) Xor (i - 11.3)
    Next i
    算个鸟猪算个鸟 = StrConv(tmpArray, vbUnicode)
End Function

Function 垃圾叼自爆(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 15.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13.8) = InArray((i * 2 - 4.2)) Xor (i - 16.2)
    Next i
    垃圾叼自爆 = StrConv(tmpArray, vbUnicode)
End Function

Function 叼悲哀蠢他妈厚脸皮艹小杂碎滚犊子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 9.9) = InArray((i * 2 - 14.4)) Xor (i - 28.8)
    Next i
    叼悲哀蠢他妈厚脸皮艹小杂碎滚犊子 = StrConv(tmpArray, vbUnicode)
End Function

Function 脑袋有洞厚脸皮厚颜无耻辣鸡狗儿子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.6) = InArray((i * 2 - 11.1)) Xor (i - 2.3)
    Next i
    脑袋有洞厚脸皮厚颜无耻辣鸡狗儿子 = StrConv(tmpArray, vbUnicode)
End Function

Function 逼狗儿子厌恶操你爸败家子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 7.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.8) = InArray((i * 2 - 2.2)) Xor (i - 13.4)
    Next i
    逼狗儿子厌恶操你爸败家子 = StrConv(tmpArray, vbUnicode)
End Function

Function 敷衍狗儿子杂种妈的耻辱傻逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 11.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.2) = InArray((i * 2 - 19.7)) Xor (i - 14.3)
    Next i
    敷衍狗儿子杂种妈的耻辱傻逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 喷叼叼杂种艹智障你个算个鸟(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 26.5) = InArray((i * 2 - 7.5)) Xor (i - 18.3)
    Next i
    喷叼叼杂种艹智障你个算个鸟 = StrConv(tmpArray, vbUnicode)
End Function

Function 傻逼操你爸算个鸟呵呵厚脸皮妈的喷(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 14.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 27.3) = InArray((i * 2 - 1.2)) Xor (i - 0.9)
    Next i
    傻逼操你爸算个鸟呵呵厚脸皮妈的喷 = StrConv(tmpArray, vbUnicode)
End Function

Function 弟弟厚颜无耻妈的(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.3) = InArray((i * 2 - 17.1)) Xor (i - 25.2)
    Next i
    弟弟厚颜无耻妈的 = StrConv(tmpArray, vbUnicode)
End Function

Function 傻逼小杂碎孙子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 25.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.9) = InArray((i * 2 - 17)) Xor (i - 24)
    Next i
    傻逼小杂碎孙子 = StrConv(tmpArray, vbUnicode)
End Function

Function 贱种辣鸡二逼狗儿子你妈喷喷恶心的(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 22.3) = InArray((i * 2 - 8.2)) Xor (i - 12.1)
    Next i
    贱种辣鸡二逼狗儿子你妈喷喷恶心的 = StrConv(tmpArray, vbUnicode)
End Function

Function 呵呵草泥马你照照镜子看看你妈小杂碎狗弱智(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 13.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 10.1) = InArray((i * 2 - 6.9)) Xor (i - 6.6)
    Next i
    呵呵草泥马你照照镜子看看你妈小杂碎狗弱智 = StrConv(tmpArray, vbUnicode)
End Function

Function 你哪来的勇气算个鸟艹垃圾逼逼你个蠢(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 20 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 12) = InArray((i * 2 - 4.6)) Xor (i - 9)
    Next i
    你哪来的勇气算个鸟艹垃圾逼逼你个蠢 = StrConv(tmpArray, vbUnicode)
End Function

Function 算个鸟贱人垃圾逼逼猪自爆(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 12.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 29) = InArray((i * 2 - 29.5)) Xor (i - 29)
    Next i
    算个鸟贱人垃圾逼逼猪自爆 = StrConv(tmpArray, vbUnicode)
End Function

Function 恶心的恶心的智障恶心的蠢厚颜无耻(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 6.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 20.5) = InArray((i * 2 - 21.4)) Xor (i - 16.9)
    Next i
    恶心的恶心的智障恶心的蠢厚颜无耻 = StrConv(tmpArray, vbUnicode)
End Function

Function 窝囊废你妈你你爸(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 29.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.7) = InArray((i * 2 - 14)) Xor (i - 25.6)
    Next i
    窝囊废你妈你你爸 = StrConv(tmpArray, vbUnicode)
End Function

Function 孙子猪脑袋有洞贱人艹(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 18.4) = InArray((i * 2 - 0.1)) Xor (i - 8.5)
    Next i
    孙子猪脑袋有洞贱人艹 = StrConv(tmpArray, vbUnicode)
End Function

Function 弟弟智障垃圾逼逼脑袋有洞呵呵你人模狗样(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 1.3) = InArray((i * 2 - 17.7)) Xor (i - 7.8)
    Next i
    弟弟智障垃圾逼逼脑袋有洞呵呵你人模狗样 = StrConv(tmpArray, vbUnicode)
End Function

Function 草泥马人模狗样二逼恶心的呵呵辣鸡(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.9) = InArray((i * 2 - 29.1)) Xor (i - 25.7)
    Next i
    草泥马人模狗样二逼恶心的呵呵辣鸡 = StrConv(tmpArray, vbUnicode)
End Function

Function 悲哀你妈辣鸡厌恶妈的呵呵厚脸皮(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 13.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.2) = InArray((i * 2 - 27.1)) Xor (i - 10.4)
    Next i
    悲哀你妈辣鸡厌恶妈的呵呵厚脸皮 = StrConv(tmpArray, vbUnicode)
End Function

Function 败家子蠢货认知障碍怂装逼怂叼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 25.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 26.6) = InArray((i * 2 - 19.1)) Xor (i - 2.3)
    Next i
    败家子蠢货认知障碍怂装逼怂叼 = StrConv(tmpArray, vbUnicode)
End Function

Function 败家子二逼耻辱你哪来的勇气(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 25.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 12.7) = InArray((i * 2 - 17.3)) Xor (i - 5.3)
    Next i
    败家子二逼耻辱你哪来的勇气 = StrConv(tmpArray, vbUnicode)
End Function

Function 智障弟弟你哪来的勇气敷衍弟弟(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 11.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.8) = InArray((i * 2 - 27.1)) Xor (i - 10)
    Next i
    智障弟弟你哪来的勇气敷衍弟弟 = StrConv(tmpArray, vbUnicode)
End Function

Function 逼逼喷草虚无(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 13.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 22.1) = InArray((i * 2 - 27.8)) Xor (i - 24.3)
    Next i
    逼逼喷草虚无 = StrConv(tmpArray, vbUnicode)
End Function

Function 呵呵贱种草泥马叼小杂碎他妈怂(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 5.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.1) = InArray((i * 2 - 2.4)) Xor (i - 17.4)
    Next i
    呵呵贱种草泥马叼小杂碎他妈怂 = StrConv(tmpArray, vbUnicode)
End Function

Function 艹小杂碎厌恶脑袋有洞(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 10.5) = InArray((i * 2 - 22.3)) Xor (i - 25)
    Next i
    艹小杂碎厌恶脑袋有洞 = StrConv(tmpArray, vbUnicode)
End Function

Function 残缺虚无你算个鸟叼垃圾狗弱智(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 24.5) = InArray((i * 2 - 20.8)) Xor (i - 29.6)
    Next i
    残缺虚无你算个鸟叼垃圾狗弱智 = StrConv(tmpArray, vbUnicode)
End Function

Function 猪猪你妈生你干啥小兔崽子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 29.2) = InArray((i * 2 - 27.5)) Xor (i - 9.8)
    Next i
    猪猪你妈生你干啥小兔崽子 = StrConv(tmpArray, vbUnicode)
End Function

Function 装逼妈的你哪来的勇气(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 0) = InArray((i * 2 - 10.4)) Xor (i - 18)
    Next i
    装逼妈的你哪来的勇气 = StrConv(tmpArray, vbUnicode)
End Function

Function 喷滚犊子败家子你个(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 25.3) = InArray((i * 2 - 13.7)) Xor (i - 5.6)
    Next i
    喷滚犊子败家子你个 = StrConv(tmpArray, vbUnicode)
End Function

Function 你哪来的勇气悲哀贱人混账(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 7.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 5.6) = InArray((i * 2 - 13.7)) Xor (i - 4)
    Next i
    你哪来的勇气悲哀贱人混账 = StrConv(tmpArray, vbUnicode)
End Function

Function 敷衍走狗蠢厚颜无耻狗儿子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 6.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 28.9) = InArray((i * 2 - 6.9)) Xor (i - 8.3)
    Next i
    敷衍走狗蠢厚颜无耻狗儿子 = StrConv(tmpArray, vbUnicode)
End Function

Function 耻辱煞笔混账傻逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.7) = InArray((i * 2 - 24)) Xor (i - 15.6)
    Next i
    耻辱煞笔混账傻逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 你哪来的勇气孙子你照照镜子看看贱人走狗你混账(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.8) = InArray((i * 2 - 8.5)) Xor (i - 8.3)
    Next i
    你哪来的勇气孙子你照照镜子看看贱人走狗你混账 = StrConv(tmpArray, vbUnicode)
End Function

Function 逼逼逼垃圾小兔崽子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.5) = InArray((i * 2 - 22.6)) Xor (i - 3.4)
    Next i
    逼逼逼垃圾小兔崽子 = StrConv(tmpArray, vbUnicode)
End Function

Function 贱人智障残缺二逼滚犊子逼逼垃圾(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 12.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 20.5) = InArray((i * 2 - 1.4)) Xor (i - 15.6)
    Next i
    贱人智障残缺二逼滚犊子逼逼垃圾 = StrConv(tmpArray, vbUnicode)
End Function

Function 脑袋有洞装逼厌恶逼逼耻辱妈的(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 18 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 28) = InArray((i * 2 - 13)) Xor (i - 26.4)
    Next i
    脑袋有洞装逼厌恶逼逼耻辱妈的 = StrConv(tmpArray, vbUnicode)
End Function

Function 草你妈蠢货小兔崽子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 29.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 18.8) = InArray((i * 2 - 19.1)) Xor (i - 29.7)
    Next i
    草你妈蠢货小兔崽子 = StrConv(tmpArray, vbUnicode)
End Function

Function 垃圾你妈生你干啥怂喷弟弟虚无脑残厚颜无耻(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 30) = InArray((i * 2 - 19.1)) Xor (i - 20.8)
    Next i
    垃圾你妈生你干啥怂喷弟弟虚无脑残厚颜无耻 = StrConv(tmpArray, vbUnicode)
End Function

Function 艹残缺孙子你逼耻辱(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 24.3) = InArray((i * 2 - 18.8)) Xor (i - 4.1)
    Next i
    艹残缺孙子你逼耻辱 = StrConv(tmpArray, vbUnicode)
End Function

Function 猪蠢傻逼垃圾逼逼耻辱猪弱智脑袋有洞(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 15.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.9) = InArray((i * 2 - 7.1)) Xor (i - 10.9)
    Next i
    猪蠢傻逼垃圾逼逼耻辱猪弱智脑袋有洞 = StrConv(tmpArray, vbUnicode)
End Function

Function 你麻痹他妈煞笔残缺叼厌恶操你爸逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.9) = InArray((i * 2 - 27)) Xor (i - 17.8)
    Next i
    你麻痹他妈煞笔残缺叼厌恶操你爸逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 脑残算个鸟滚犊子操你爸狗你哪来的勇气(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 11.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 1.2) = InArray((i * 2 - 1.6)) Xor (i - 9.8)
    Next i
    脑残算个鸟滚犊子操你爸狗你哪来的勇气 = StrConv(tmpArray, vbUnicode)
End Function

Function 怂你哪来的勇气你滚犊子你哪来的勇气(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 24.9) = InArray((i * 2 - 20.9)) Xor (i - 3.4)
    Next i
    怂你哪来的勇气你滚犊子你哪来的勇气 = StrConv(tmpArray, vbUnicode)
End Function

Function 杂种虚无脑残辣鸡敷衍操你爸厌恶(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 27) = InArray((i * 2 - 17.3)) Xor (i - 2.3)
    Next i
    杂种虚无脑残辣鸡敷衍操你爸厌恶 = StrConv(tmpArray, vbUnicode)
End Function

Function 妈的二逼认知障碍草泥马弱智(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 2.2) = InArray((i * 2 - 8.8)) Xor (i - 28.9)
    Next i
    妈的二逼认知障碍草泥马弱智 = StrConv(tmpArray, vbUnicode)
End Function

Function 他妈滚犊子猪垃圾辣鸡(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 7.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 2.8) = InArray((i * 2 - 28.1)) Xor (i - 19.7)
    Next i
    他妈滚犊子猪垃圾辣鸡 = StrConv(tmpArray, vbUnicode)
End Function

Function 怂艹厚脸皮猪(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 25.9) = InArray((i * 2 - 10.7)) Xor (i - 23)
    Next i
    怂艹厚脸皮猪 = StrConv(tmpArray, vbUnicode)
End Function

Function 叼弟弟认知障碍(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 20.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 20.9) = InArray((i * 2 - 25.7)) Xor (i - 12.5)
    Next i
    叼弟弟认知障碍 = StrConv(tmpArray, vbUnicode)
End Function

Function 厚脸皮垃圾狗窝囊废(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 18 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 29.5) = InArray((i * 2 - 22.3)) Xor (i - 24.1)
    Next i
    厚脸皮垃圾狗窝囊废 = StrConv(tmpArray, vbUnicode)
End Function

Function 蠢货草你妈狗儿子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13.8) = InArray((i * 2 - 22.6)) Xor (i - 2.6)
    Next i
    蠢货草你妈狗儿子 = StrConv(tmpArray, vbUnicode)
End Function

Function 你麻痹他妈败家子喷你哪来的勇气(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13.9) = InArray((i * 2 - 16.8)) Xor (i - 29.2)
    Next i
    你麻痹他妈败家子喷你哪来的勇气 = StrConv(tmpArray, vbUnicode)
End Function

Function 人模狗样算个鸟草你妈猪煞笔蠢货(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 8.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 0.2) = InArray((i * 2 - 5.6)) Xor (i - 0.1)
    Next i
    人模狗样算个鸟草你妈猪煞笔蠢货 = StrConv(tmpArray, vbUnicode)
End Function

Function 蠢货虚无猪算个鸟你弱智(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 27.2) = InArray((i * 2 - 19.9)) Xor (i - 29.5)
    Next i
    蠢货虚无猪算个鸟你弱智 = StrConv(tmpArray, vbUnicode)
End Function

Function 悲哀贱人虚无草你妈你个呵呵(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 3.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 29) = InArray((i * 2 - 26.7)) Xor (i - 8.2)
    Next i
    悲哀贱人虚无草你妈你个呵呵 = StrConv(tmpArray, vbUnicode)
End Function

Function 脑袋有洞厚脸皮猪他妈怂(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16) = InArray((i * 2 - 6.4)) Xor (i - 29.4)
    Next i
    脑袋有洞厚脸皮猪他妈怂 = StrConv(tmpArray, vbUnicode)
End Function

Function 弟弟你个认知障碍滚犊子狗操你爸杂种煞笔(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.4) = InArray((i * 2 - 12.4)) Xor (i - 22)
    Next i
    弟弟你个认知障碍滚犊子狗操你爸杂种煞笔 = StrConv(tmpArray, vbUnicode)
End Function

Function 辣鸡妈的走狗小兔崽子弱智小杂碎智障贱人(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 1.4) = InArray((i * 2 - 7.9)) Xor (i - 13.2)
    Next i
    辣鸡妈的走狗小兔崽子弱智小杂碎智障贱人 = StrConv(tmpArray, vbUnicode)
End Function

Function 草泥马装逼艹自爆小兔崽子你妈生你干啥(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 22.8) = InArray((i * 2 - 27)) Xor (i - 3.7)
    Next i
    草泥马装逼艹自爆小兔崽子你妈生你干啥 = StrConv(tmpArray, vbUnicode)
End Function

Function 你照照镜子看看窝囊废猪弱智煞笔猪你敷衍(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 22.3) = InArray((i * 2 - 14.1)) Xor (i - 20.4)
    Next i
    你照照镜子看看窝囊废猪弱智煞笔猪你敷衍 = StrConv(tmpArray, vbUnicode)
End Function

Function 草泥马耻辱你照照镜子看看(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 17.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.9) = InArray((i * 2 - 10.2)) Xor (i - 21.2)
    Next i
    草泥马耻辱你照照镜子看看 = StrConv(tmpArray, vbUnicode)
End Function

Function 贱种装逼人模狗样艹艹煞笔装逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 12.2) = InArray((i * 2 - 25.8)) Xor (i - 26.2)
    Next i
    贱种装逼人模狗样艹艹煞笔装逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 你爸小兔崽子草泥马呵呵操你爸残缺辣鸡(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.8) = InArray((i * 2 - 23.1)) Xor (i - 27.7)
    Next i
    你爸小兔崽子草泥马呵呵操你爸残缺辣鸡 = StrConv(tmpArray, vbUnicode)
End Function

Function 狗儿子贱种垃圾妈的(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 25) = InArray((i * 2 - 21.1)) Xor (i - 19.8)
    Next i
    狗儿子贱种垃圾妈的 = StrConv(tmpArray, vbUnicode)
End Function

Function 操你爸呵呵二逼悲哀小兔崽子脑残(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 9.1) = InArray((i * 2 - 22.9)) Xor (i - 7.8)
    Next i
    操你爸呵呵二逼悲哀小兔崽子脑残 = StrConv(tmpArray, vbUnicode)
End Function

Function 残缺你哪来的勇气狗贱种狗儿子走狗混账(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 17.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 6.6) = InArray((i * 2 - 15.6)) Xor (i - 25.2)
    Next i
    残缺你哪来的勇气狗贱种狗儿子走狗混账 = StrConv(tmpArray, vbUnicode)
End Function

Function 你照照镜子看看你怂走狗厚颜无耻(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 9.4) = InArray((i * 2 - 27.7)) Xor (i - 9.8)
    Next i
    你照照镜子看看你怂走狗厚颜无耻 = StrConv(tmpArray, vbUnicode)
End Function

Function 蠢艹喷你爸辣鸡智障(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.1) = InArray((i * 2 - 5.4)) Xor (i - 12.4)
    Next i
    蠢艹喷你爸辣鸡智障 = StrConv(tmpArray, vbUnicode)
End Function

Function 煞笔草二逼人模狗样叼小杂碎走狗败家子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 22.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 8.7) = InArray((i * 2 - 17.1)) Xor (i - 12.8)
    Next i
    煞笔草二逼人模狗样叼小杂碎走狗败家子 = StrConv(tmpArray, vbUnicode)
End Function

Function 败家子猪妈的恶心的脑袋有洞呵呵智障草(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 26.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 3.2) = InArray((i * 2 - 2.1)) Xor (i - 3.9)
    Next i
    败家子猪妈的恶心的脑袋有洞呵呵智障草 = StrConv(tmpArray, vbUnicode)
End Function

Function 智障他妈装逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.1) = InArray((i * 2 - 28.5)) Xor (i - 6.2)
    Next i
    智障他妈装逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 狗儿子你蠢货垃圾认知障碍煞笔垃圾逼逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 5.8) = InArray((i * 2 - 23.8)) Xor (i - 15.5)
    Next i
    狗儿子你蠢货垃圾认知障碍煞笔垃圾逼逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 怂算个鸟耻辱呵呵猪认知障碍草泥马(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 20.4) = InArray((i * 2 - 8.2)) Xor (i - 3.7)
    Next i
    怂算个鸟耻辱呵呵猪认知障碍草泥马 = StrConv(tmpArray, vbUnicode)
End Function

Function 怂辣鸡认知障碍逼逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 22.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.6) = InArray((i * 2 - 13.3)) Xor (i - 11.4)
    Next i
    怂辣鸡认知障碍逼逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 你照照镜子看看智障逼逼恶心的弱智(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 26.7) = InArray((i * 2 - 22.7)) Xor (i - 19.1)
    Next i
    你照照镜子看看智障逼逼恶心的弱智 = StrConv(tmpArray, vbUnicode)
End Function

Function 妈的垃圾喷敷衍厌恶杂种窝囊废(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 26.5) = InArray((i * 2 - 16.1)) Xor (i - 12.4)
    Next i
    妈的垃圾喷敷衍厌恶杂种窝囊废 = StrConv(tmpArray, vbUnicode)
End Function

Function 滚犊子混账自爆弱智虚无窝囊废厚颜无耻虚无(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 18 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 6.4) = InArray((i * 2 - 24.6)) Xor (i - 8.5)
    Next i
    滚犊子混账自爆弱智虚无窝囊废厚颜无耻虚无 = StrConv(tmpArray, vbUnicode)
End Function

Function 走狗自爆悲哀妈的狗儿子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 17.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.8) = InArray((i * 2 - 21.9)) Xor (i - 2.3)
    Next i
    走狗自爆悲哀妈的狗儿子 = StrConv(tmpArray, vbUnicode)
End Function

Function 败家子你滚犊子恶心的你(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.3) = InArray((i * 2 - 17)) Xor (i - 1.3)
    Next i
    败家子你滚犊子恶心的你 = StrConv(tmpArray, vbUnicode)
End Function

Function 叼艹你妈生你干啥小杂碎狗儿子你个操你爸(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 15.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 24.7) = InArray((i * 2 - 23.1)) Xor (i - 14.1)
    Next i
    叼艹你妈生你干啥小杂碎狗儿子你个操你爸 = StrConv(tmpArray, vbUnicode)
End Function

Function 怂贱种弱智狗小兔崽子小兔崽子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.5) = InArray((i * 2 - 5.8)) Xor (i - 17)
    Next i
    怂贱种弱智狗小兔崽子小兔崽子 = StrConv(tmpArray, vbUnicode)
End Function

Function 悲哀你哪来的勇气猪败家子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 8.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 0.6) = InArray((i * 2 - 27.3)) Xor (i - 10.9)
    Next i
    悲哀你哪来的勇气猪败家子 = StrConv(tmpArray, vbUnicode)
End Function

Function 逼逼你个你麻痹狗他妈蠢货(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 8.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 23.7) = InArray((i * 2 - 24.8)) Xor (i - 18.6)
    Next i
    逼逼你个你麻痹狗他妈蠢货 = StrConv(tmpArray, vbUnicode)
End Function

Function 逼逼你妈草蠢货猪你爸你爸(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7) = InArray((i * 2 - 8.8)) Xor (i - 6.1)
    Next i
    逼逼你妈草蠢货猪你爸你爸 = StrConv(tmpArray, vbUnicode)
End Function

Function 二逼草你妈小杂碎蠢草泥马草你个(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 9.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 0.1) = InArray((i * 2 - 14.7)) Xor (i - 13.7)
    Next i
    二逼草你妈小杂碎蠢草泥马草你个 = StrConv(tmpArray, vbUnicode)
End Function

'String decode
Function 傻逼你哪来的勇气妈的自爆你哪来的勇气(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte
    
    ReDim tmpArray(UBound(InArray))
    For i = 13.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13.5) = InArray((i * 2 - 27) / 2) Xor (i - 13.5)
    Next i
    
    傻逼你哪来的勇气妈的自爆你哪来的勇气 = StrConv(tmpArray, vbUnicode)
End Function

Function 你麻痹逼逼走狗艹败家子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 13.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.7) = InArray((i * 2 - 6.6)) Xor (i - 6.5)
    Next i
    你麻痹逼逼走狗艹败家子 = StrConv(tmpArray, vbUnicode)
End Function

'Loop check (InArray is useless)
Function 厚颜无耻你妈生你干啥傻逼你爸猪败家子小杂碎(InArray As Variant) As String
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
        '"孙子喷呵呵草泥马敷衍"
        tmp = Array(148, 176, 136, 140, 154, 184, 229, 152, 229, 152, 237, 130, 155, 191, 157, 178, 232, 172, 142, 131)
        ReDim ProcName(UBound(tmp))
        For Useless = 0 To UBound(tmp)
            ProcName(Useless) = tmp(Useless) Xor (95 + Useless * Useless) - Useless ^ 2
        Next Useless
        CallByName Me, StrConv(ProcName, vbUnicode), VbMethod, BloodValue
        
        '"傻逼你哪来的勇气妈的自爆你哪来的勇气"
        tmp = Array(249, 133, 129, 246, 244, 211, 244, 244, 240, 132, 133, 244, 227, 242, 246, 200, 242, 216, 133, 244, 231, 228, 129, 156, 244, 211, 244, 244, 240, 132, 133, 244, 227, 242, 246, 200)
        ReDim ProcName(UBound(tmp))
        For Useless = 0 To UBound(tmp)
            ProcName(Useless) = tmp(Useless) Xor (48 + Useless * Useless) - Useless ^ 2
        Next Useless
        Me.傻逼你看到控件名是没用的.Caption = CallByName(Me, StrConv(ProcName, vbUnicode), VbMethod, LABEL_CURRENT_BLOOD) & BloodValue(4)
        '--------------
        'Check if the fake value is changed
        If CLng(FakeBlood) <> CLng(BloodValue(4)) Then
            If Not CE_Changed Then                                                  'If not prompted
                CE_Changed = True
                Me.傻逼你看到控件名是没用的.Caption = CallByName(Me, StrConv(ProcName, vbUnicode), VbMethod, LABEL_CURRENT_BLOOD) & FakeBlood
                StartTime = GetTickCount
                Do While GetTickCount - StartTime < 3000                                'Sleep 3000ms, without blocking the process
                    Sleep 10
                    If Exiting Then
                        Exit Do
                    End If
                    DoEvents
                Loop
                Me.能看到控件名很了不起吗.Caption = CallByName(Me, StrConv(ProcName, vbUnicode), VbMethod, CE_PROMPT_1)
                StartTime = GetTickCount
                Do While GetTickCount - StartTime < 2500
                    Sleep 10
                    DoEvents
                    If Exiting Then
                        Exit Do
                    End If
                Loop
                Me.能看到控件名很了不起吗.Caption = CallByName(Me, StrConv(ProcName, vbUnicode), VbMethod, CE_PROMPT_2)
                StartTime = GetTickCount
                Do While GetTickCount - StartTime < 2500
                    Sleep 10
                    DoEvents
                    If Exiting Then
                        Exit Do
                    End If
                Loop
                Me.能看到控件名很了不起吗.Caption = CallByName(Me, StrConv(ProcName, vbUnicode), VbMethod, CE_PROMPT_3)
                Me.傻逼你看到控件名是没用的.Caption = CallByName(Me, StrConv(ProcName, vbUnicode), VbMethod, LABEL_CURRENT_BLOOD) & BloodValue(4)
                StartTime = GetTickCount
                Do While GetTickCount - StartTime < 2500
                    Sleep 10
                    DoEvents
                    If Exiting Then
                        Exit Do
                    End If
                Loop
                Me.能看到控件名很了不起吗.Caption = ""
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
            Me.能看到控件名很了不起吗.Caption = CallByName(Me, StrConv(ProcName, vbUnicode), VbMethod, SUCCEED_PROMPT)
        Else
            Me.能看到控件名很了不起吗.Caption = ""
        End If
        
        '--------------
        Sleep 50
        DoEvents
    Loop
    
    SetWindowLong Me.hwnd, GWL_WNDPROC, PrevWndProc             'Restore the default window proc.
    Unload Me
End Function

Function 脑袋有洞草你妈操你爸(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 26 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.2) = InArray((i * 2 - 18.3)) Xor (i - 18.9)
    Next i
    脑袋有洞草你妈操你爸 = StrConv(tmpArray, vbUnicode)
End Function

'Get value
'Return value = InArray(4)
Function 孙子喷呵呵草泥马敷衍(ByRef InArray As Variant) As String
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
    孙子喷呵呵草泥马敷衍 = StrConv(Useless2, vbUnicode)
    '=========================================== End
    
    For i = 2.65 To 2.65 + 3                                                'For i = 0 To 3
        temp(i - 2.65) = NumberData(MemIndex(i - 2.65))                     '   temp(i) = NumberData(MemIndex(i))
    Next i
    CopyMemory ret, temp(0), (2 + ret * ret - ret ^ 2) * 2                  'CopyMemory ret, temp(0), 4
    
    'Only InArray(4) is meaningful
    InArray = Array((216 + i * i) - i ^ 2, (226 + i * i) - i ^ 2, (298 + i * i) - i ^ 2, (197 + i * i) - i ^ 2, _
        CLng((ret + i * i) - i ^ 2), (246 + i * i) - i ^ 2, (159 + i * i) - i ^ 2, (246 + i * i) - i ^ 2, (241 + i * i) - i ^ 2, (250 + i * i) - i ^ 2)
End Function

Function 厚颜无耻蠢货败家子滚犊子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 29.5) = InArray((i * 2 - 2.3)) Xor (i - 20.5)
    Next i
    厚颜无耻蠢货败家子滚犊子 = StrConv(tmpArray, vbUnicode)
End Function

Function 傻逼草你妈你妈生你干啥草泥马(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 19.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13) = InArray((i * 2 - 28.5)) Xor (i - 14.1)
    Next i
    傻逼草你妈你妈生你干啥草泥马 = StrConv(tmpArray, vbUnicode)
End Function

Function 智障草泥马辣鸡叼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 26.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 10.8) = InArray((i * 2 - 4)) Xor (i - 23.6)
    Next i
    智障草泥马辣鸡叼 = StrConv(tmpArray, vbUnicode)
End Function

Function 耻辱逼恶心的耻辱算个鸟你(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 25.9) = InArray((i * 2 - 28.6)) Xor (i - 22.8)
    Next i
    耻辱逼恶心的耻辱算个鸟你 = StrConv(tmpArray, vbUnicode)
End Function

Function 悲哀算个鸟小杂碎(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 17.2) = InArray((i * 2 - 24)) Xor (i - 27.6)
    Next i
    悲哀算个鸟小杂碎 = StrConv(tmpArray, vbUnicode)
End Function

Function 智障败家子逼逼败家子厚颜无耻杂种(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 3.5) = InArray((i * 2 - 14.9)) Xor (i - 1.3)
    Next i
    智障败家子逼逼败家子厚颜无耻杂种 = StrConv(tmpArray, vbUnicode)
End Function

Function 混账恶心的你爸(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 5.7) = InArray((i * 2 - 7.7)) Xor (i - 28.1)
    Next i
    混账恶心的你爸 = StrConv(tmpArray, vbUnicode)
End Function

Function 你麻痹草你妈煞笔人模狗样操你爸(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 13.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 28.4) = InArray((i * 2 - 1.3)) Xor (i - 4.9)
    Next i
    你麻痹草你妈煞笔人模狗样操你爸 = StrConv(tmpArray, vbUnicode)
End Function

Function 杂种厚颜无耻喷混账垃圾逼逼呵呵(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 18.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.9) = InArray((i * 2 - 10.3)) Xor (i - 21.1)
    Next i
    杂种厚颜无耻喷混账垃圾逼逼呵呵 = StrConv(tmpArray, vbUnicode)
End Function

Function 操你爸败家子狗草你妈悲哀装逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 20 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.2) = InArray((i * 2 - 27.7)) Xor (i - 26.8)
    Next i
    操你爸败家子狗草你妈悲哀装逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 艹艹耻辱狗儿子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.1) = InArray((i * 2 - 7)) Xor (i - 29.1)
    Next i
    艹艹耻辱狗儿子 = StrConv(tmpArray, vbUnicode)
End Function

Function 垃圾逼逼你照照镜子看看自爆蠢逼走狗(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.1) = InArray((i * 2 - 26.5)) Xor (i - 17.3)
    Next i
    垃圾逼逼你照照镜子看看自爆蠢逼走狗 = StrConv(tmpArray, vbUnicode)
End Function

Function 滚犊子叼窝囊废狗小兔崽子敷衍厚脸皮悲哀(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 9.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 22.1) = InArray((i * 2 - 13)) Xor (i - 0.3)
    Next i
    滚犊子叼窝囊废狗小兔崽子敷衍厚脸皮悲哀 = StrConv(tmpArray, vbUnicode)
End Function

Function 敷衍脑袋有洞操你爸猪(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 15.7) = InArray((i * 2 - 8.3)) Xor (i - 28.7)
    Next i
    敷衍脑袋有洞操你爸猪 = StrConv(tmpArray, vbUnicode)
End Function

Function 人模狗样你照照镜子看看叼走狗残缺弱智喷(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13.5) = InArray((i * 2 - 8.3)) Xor (i - 29.4)
    Next i
    人模狗样你照照镜子看看叼走狗残缺弱智喷 = StrConv(tmpArray, vbUnicode)
End Function

Function 逼逼虚无你麻痹弱智败家子贱人(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13.1) = InArray((i * 2 - 28.6)) Xor (i - 29.9)
    Next i
    逼逼虚无你麻痹弱智败家子贱人 = StrConv(tmpArray, vbUnicode)
End Function

Function 小杂碎自爆艹你哪来的勇气猪垃圾逼逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 15.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.7) = InArray((i * 2 - 14.3)) Xor (i - 28.6)
    Next i
    小杂碎自爆艹你哪来的勇气猪垃圾逼逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 残缺走狗你爸喷草你妈(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 19.8) = InArray((i * 2 - 1.2)) Xor (i - 11.8)
    Next i
    残缺走狗你爸喷草你妈 = StrConv(tmpArray, vbUnicode)
End Function

Function 逼逼猪你麻痹妈的残缺贱种(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 6.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.6) = InArray((i * 2 - 11.1)) Xor (i - 12.1)
    Next i
    逼逼猪你麻痹妈的残缺贱种 = StrConv(tmpArray, vbUnicode)
End Function

Function 悲哀叼厚颜无耻装逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 18.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.2) = InArray((i * 2 - 19.4)) Xor (i - 10.3)
    Next i
    悲哀叼厚颜无耻装逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 贱人猪叼智障狗儿子你照照镜子看看你爸(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 17.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 29.8) = InArray((i * 2 - 4.1)) Xor (i - 20.6)
    Next i
    贱人猪叼智障狗儿子你照照镜子看看你爸 = StrConv(tmpArray, vbUnicode)
End Function

Function 狗小兔崽子自爆恶心的算个鸟(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 22.6) = InArray((i * 2 - 1.4)) Xor (i - 3.8)
    Next i
    狗小兔崽子自爆恶心的算个鸟 = StrConv(tmpArray, vbUnicode)
End Function

Function 垃圾艹垃圾你个(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.5) = InArray((i * 2 - 18.7)) Xor (i - 24.2)
    Next i
    垃圾艹垃圾你个 = StrConv(tmpArray, vbUnicode)
End Function

'Init. string byte arrays
Function 操你妈了个鸡巴白痴混账脑残(InArray As Variant) As String
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
    ButtonHwnd = Me.是不是看到控件名觉得自己很厉害呢.hwnd                           'Record the handle to the button
    
    'Set window proc.
    PrevWndProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf 操你妈逼混账你个败家子逼逼)
    
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
        .lpszClassName = VarPtr(NewClassName(0))                                    '"哇你发现了隐藏的我！奖励你一朵小红花(#^.^#)"
        .cbSize = Len(ctlClass)
    End With
    Dim r As Long
    r = RegisterClassEx(ctlClass)
    CreateWindowEx 0, StrConv(NewClassName, vbUnicode), "", WS_CHILD, 10, 10, 100, 100, Me.hwnd, 0, App.hinstance, 0
End Function

Function 虚无你妈生你干啥草你妈耻辱走狗(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 12.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 19.1) = InArray((i * 2 - 6.1)) Xor (i - 23.8)
    Next i
    虚无你妈生你干啥草你妈耻辱走狗 = StrConv(tmpArray, vbUnicode)
End Function

Function 认知障碍脑残逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 8.1) = InArray((i * 2 - 17.6)) Xor (i - 12.7)
    Next i
    认知障碍脑残逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 小杂碎脑袋有洞小兔崽子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.3) = InArray((i * 2 - 3.6)) Xor (i - 29.8)
    Next i
    小杂碎脑袋有洞小兔崽子 = StrConv(tmpArray, vbUnicode)
End Function

Function 耻辱你个蠢蠢敷衍厚颜无耻贱人喷(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 29.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 0.3) = InArray((i * 2 - 14.4)) Xor (i - 10.2)
    Next i
    耻辱你个蠢蠢敷衍厚颜无耻贱人喷 = StrConv(tmpArray, vbUnicode)
End Function

Function 你妈你妈滚犊子猪他妈辣鸡(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 14.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 24.4) = InArray((i * 2 - 17.3)) Xor (i - 29.7)
    Next i
    你妈你妈滚犊子猪他妈辣鸡 = StrConv(tmpArray, vbUnicode)
End Function

Function 败家子你妈混账算个鸟狗儿子孙子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 23.4) = InArray((i * 2 - 22.1)) Xor (i - 0.3)
    Next i
    败家子你妈混账算个鸟狗儿子孙子 = StrConv(tmpArray, vbUnicode)
End Function

Function 认知障碍滚犊子艹蠢小杂碎傻逼恶心的弱智(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 5.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 10.1) = InArray((i * 2 - 13.1)) Xor (i - 25.7)
    Next i
    认知障碍滚犊子艹蠢小杂碎傻逼恶心的弱智 = StrConv(tmpArray, vbUnicode)
End Function

Function 你哪来的勇气残缺恶心的耻辱喷(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 28 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 10.4) = InArray((i * 2 - 15.2)) Xor (i - 3)
    Next i
    你哪来的勇气残缺恶心的耻辱喷 = StrConv(tmpArray, vbUnicode)
End Function

Function 悲哀逼弟弟(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 8.1) = InArray((i * 2 - 11.3)) Xor (i - 5.8)
    Next i
    悲哀逼弟弟 = StrConv(tmpArray, vbUnicode)
End Function

Function 蠢狗儿子脑袋有洞败家子猪叼草你妈(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 17.3) = InArray((i * 2 - 0.6)) Xor (i - 13.1)
    Next i
    蠢狗儿子脑袋有洞败家子猪叼草你妈 = StrConv(tmpArray, vbUnicode)
End Function

Function 你麻痹贱人厚脸皮人模狗样杂种算个鸟恶心的(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.2) = InArray((i * 2 - 24.6)) Xor (i - 15.9)
    Next i
    你麻痹贱人厚脸皮人模狗样杂种算个鸟恶心的 = StrConv(tmpArray, vbUnicode)
End Function

Function 自爆草泥马喷猪走狗贱人(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.6) = InArray((i * 2 - 18.2)) Xor (i - 24.2)
    Next i
    自爆草泥马喷猪走狗贱人 = StrConv(tmpArray, vbUnicode)
End Function

Function 敷衍贱人叼你麻痹(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 9.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.2) = InArray((i * 2 - 5.5)) Xor (i - 20.1)
    Next i
    敷衍贱人叼你麻痹 = StrConv(tmpArray, vbUnicode)
End Function

'Mess up (InArray is useless)
Function 恶心的猪脑袋有洞你妈煞笔煞笔辣鸡(InArray As Variant) As String
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

Function 狗儿子傻逼贱种残缺怂贱种(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 27.7) = InArray((i * 2 - 15.4)) Xor (i - 8.8)
    Next i
    狗儿子傻逼贱种残缺怂贱种 = StrConv(tmpArray, vbUnicode)
End Function

Function 你麻痹逼逼贱种二逼你(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 0.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.6) = InArray((i * 2 - 10.8)) Xor (i - 24.1)
    Next i
    你麻痹逼逼贱种二逼你 = StrConv(tmpArray, vbUnicode)
End Function

Function 贱种你照照镜子看看垃圾混账你哪来的勇气你个(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 8.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 19.5) = InArray((i * 2 - 15.3)) Xor (i - 14.3)
    Next i
    贱种你照照镜子看看垃圾混账你哪来的勇气你个 = StrConv(tmpArray, vbUnicode)
End Function

Function 你个煞笔走狗蠢垃圾逼逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 3.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13) = InArray((i * 2 - 21.4)) Xor (i - 28.7)
    Next i
    你个煞笔走狗蠢垃圾逼逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 弟弟傻逼贱种你照照镜子看看悲哀(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 9.3) = InArray((i * 2 - 1.3)) Xor (i - 25.1)
    Next i
    弟弟傻逼贱种你照照镜子看看悲哀 = StrConv(tmpArray, vbUnicode)
End Function

Function 敷衍你怂你个艹贱人逼逼妈的(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.1) = InArray((i * 2 - 13.9)) Xor (i - 2.6)
    Next i
    敷衍你怂你个艹贱人逼逼妈的 = StrConv(tmpArray, vbUnicode)
End Function

Function 逼逼厌恶脑残草你妈(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.5) = InArray((i * 2 - 4.1)) Xor (i - 10.1)
    Next i
    逼逼厌恶脑残草你妈 = StrConv(tmpArray, vbUnicode)
End Function

Function 草泥马残缺装逼辣鸡(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 14.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 27.3) = InArray((i * 2 - 9.1)) Xor (i - 16.1)
    Next i
    草泥马残缺装逼辣鸡 = StrConv(tmpArray, vbUnicode)
End Function

Function 你照照镜子看看脑残草泥马垃圾(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 11.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 12.3) = InArray((i * 2 - 13.4)) Xor (i - 15.3)
    Next i
    你照照镜子看看脑残草泥马垃圾 = StrConv(tmpArray, vbUnicode)
End Function

Function 你爸厚颜无耻呵呵孙子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 16.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 18.2) = InArray((i * 2 - 27.4)) Xor (i - 15.4)
    Next i
    你爸厚颜无耻呵呵孙子 = StrConv(tmpArray, vbUnicode)
End Function

Function 辣鸡智障恶心的怂脑残(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 13.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 6.9) = InArray((i * 2 - 12.7)) Xor (i - 17.7)
    Next i
    辣鸡智障恶心的怂脑残 = StrConv(tmpArray, vbUnicode)
End Function

Function 操你爸你爸你麻痹(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 28.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.9) = InArray((i * 2 - 0.6)) Xor (i - 21.7)
    Next i
    操你爸你爸你麻痹 = StrConv(tmpArray, vbUnicode)
End Function

Function 呵呵厚脸皮厚脸皮辣鸡操你爸(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 8.1) = InArray((i * 2 - 5.7)) Xor (i - 14.7)
    Next i
    呵呵厚脸皮厚脸皮辣鸡操你爸 = StrConv(tmpArray, vbUnicode)
End Function

'Set value (InArray(3) = NewValue)
'Note: InArray should be Long() type
Function 你照照镜子看看贱种弟弟残缺(InArray As Variant) As String
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
    '"恶心的猪脑袋有洞你妈煞笔煞笔辣鸡"
    tmp = Array(11, 76, 109, 121, 8, 121, 107, 80, 121, 105, 9, 65, 110, 109, 11, 9, 121, 94, 127, 85, 116, 10, 12, 119, 116, 10, 12, 119, 125, 12, 1, 27)
    Useless = StrConv(InArray(0), vbFromUnicode)                                    'Ignore this
    ReDim ProcName(UBound(tmp))
    For i = 3 To UBound(tmp) + 3                                                    'For i = 0 To Ubound(tmp)
        tmp(i - 3) = tmp(i - 3) Xor (189 + CLng(t) * t - t ^ 2)
        ProcName(i - 3) = tmp(i - 3)                                                '   ProcName(i) = tmp(i) xor 189
    Next i
    'Call mess up proc. (Call 恶心的猪脑袋有洞你妈煞笔煞笔辣鸡([Useless thing]))
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

Function 垃圾逼逼恶心的怂(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 28.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 26.2) = InArray((i * 2 - 13)) Xor (i - 16.1)
    Next i
    垃圾逼逼恶心的怂 = StrConv(tmpArray, vbUnicode)
End Function

Function 辣鸡蠢草草泥马你个草泥马孙子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 7.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 18.2) = InArray((i * 2 - 28.3)) Xor (i - 22.8)
    Next i
    辣鸡蠢草草泥马你个草泥马孙子 = StrConv(tmpArray, vbUnicode)
End Function

Function 你哪来的勇气厌恶孙子傻逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 0.5) = InArray((i * 2 - 3.3)) Xor (i - 18.5)
    Next i
    你哪来的勇气厌恶孙子傻逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 残缺滚犊子认知障碍残缺你妈生你干啥草你妈你照照镜子看看(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.7) = InArray((i * 2 - 23.2)) Xor (i - 26.1)
    Next i
    残缺滚犊子认知障碍残缺你妈生你干啥草你妈你照照镜子看看 = StrConv(tmpArray, vbUnicode)
End Function

Function 认知障碍草你妈妈的脑残(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 8.9) = InArray((i * 2 - 12.3)) Xor (i - 13.5)
    Next i
    认知障碍草你妈妈的脑残 = StrConv(tmpArray, vbUnicode)
End Function

Function 装逼脑残耻辱草你妈滚犊子小杂碎厚脸皮(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 3.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 18.5) = InArray((i * 2 - 2.5)) Xor (i - 26.5)
    Next i
    装逼脑残耻辱草你妈滚犊子小杂碎厚脸皮 = StrConv(tmpArray, vbUnicode)
End Function

Function 狗装逼你妈(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 5.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 1.5) = InArray((i * 2 - 21.9)) Xor (i - 6.6)
    Next i
    狗装逼你妈 = StrConv(tmpArray, vbUnicode)
End Function

Function 二逼你照照镜子看看脑袋有洞算个鸟你妈生你干啥脑残草(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 12.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21.8) = InArray((i * 2 - 28.9)) Xor (i - 4.3)
    Next i
    二逼你照照镜子看看脑袋有洞算个鸟你妈生你干啥脑残草 = StrConv(tmpArray, vbUnicode)
End Function

Function 你猪脑残猪(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 17.4) = InArray((i * 2 - 20.2)) Xor (i - 29.4)
    Next i
    你猪脑残猪 = StrConv(tmpArray, vbUnicode)
End Function

Function 你爸傻逼虚无辣鸡(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 22.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.4) = InArray((i * 2 - 7.6)) Xor (i - 14.5)
    Next i
    你爸傻逼虚无辣鸡 = StrConv(tmpArray, vbUnicode)
End Function

Function 他妈贱种厌恶贱种你妈生你干啥你个人模狗样(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 9.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.6) = InArray((i * 2 - 22.2)) Xor (i - 26.8)
    Next i
    他妈贱种厌恶贱种你妈生你干啥你个人模狗样 = StrConv(tmpArray, vbUnicode)
End Function

Function 叼自爆喷耻辱人模狗样逼逼蠢货(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 9.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.1) = InArray((i * 2 - 29.4)) Xor (i - 28.9)
    Next i
    叼自爆喷耻辱人模狗样逼逼蠢货 = StrConv(tmpArray, vbUnicode)
End Function

Function 猪喷耻辱厚脸皮(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.6) = InArray((i * 2 - 28.8)) Xor (i - 7.1)
    Next i
    猪喷耻辱厚脸皮 = StrConv(tmpArray, vbUnicode)
End Function

Function 厌恶他妈算个鸟你爸叼喷垃圾逼逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.9) = InArray((i * 2 - 23.6)) Xor (i - 1.6)
    Next i
    厌恶他妈算个鸟你爸叼喷垃圾逼逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 你虚无厌恶叼敷衍喷(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 10.4) = InArray((i * 2 - 14.4)) Xor (i - 3.3)
    Next i
    你虚无厌恶叼敷衍喷 = StrConv(tmpArray, vbUnicode)
End Function

Function 恶心的叼二逼恶心的喷混账(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 5.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21) = InArray((i * 2 - 0.5)) Xor (i - 27.2)
    Next i
    恶心的叼二逼恶心的喷混账 = StrConv(tmpArray, vbUnicode)
End Function

Function 贱种人模狗样草泥马恶心的算个鸟(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 18.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 23) = InArray((i * 2 - 17.9)) Xor (i - 18.1)
    Next i
    贱种人模狗样草泥马恶心的算个鸟 = StrConv(tmpArray, vbUnicode)
End Function

Function 二逼智障你哪来的勇气他妈你妈你小兔崽子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 8.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 10.3) = InArray((i * 2 - 1.7)) Xor (i - 0.6)
    Next i
    二逼智障你哪来的勇气他妈你妈你小兔崽子 = StrConv(tmpArray, vbUnicode)
End Function

Function 喷你操你爸狗儿子你个混账逼逼你照照镜子看看(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 11.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 28.9) = InArray((i * 2 - 9.8)) Xor (i - 26.3)
    Next i
    喷你操你爸狗儿子你个混账逼逼你照照镜子看看 = StrConv(tmpArray, vbUnicode)
End Function

Function 贱种猪杂种你(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 29.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.2) = InArray((i * 2 - 1.9)) Xor (i - 24.9)
    Next i
    贱种猪杂种你 = StrConv(tmpArray, vbUnicode)
End Function

Function 垃圾草你妈垃圾逼逼他妈艹逼草你妈(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 15.3) = InArray((i * 2 - 20.9)) Xor (i - 5)
    Next i
    垃圾草你妈垃圾逼逼他妈艹逼草你妈 = StrConv(tmpArray, vbUnicode)
End Function

Function 走狗垃圾脑残贱人狗猪脑残(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 13.9) = InArray((i * 2 - 14)) Xor (i - 10.1)
    Next i
    走狗垃圾脑残贱人狗猪脑残 = StrConv(tmpArray, vbUnicode)
End Function

Function 怂二逼你哪来的勇气你妈贱人艹蠢(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 2.2) = InArray((i * 2 - 17)) Xor (i - 19.3)
    Next i
    怂二逼你哪来的勇气你妈贱人艹蠢 = StrConv(tmpArray, vbUnicode)
End Function

Function 叼他妈操你爸狗耻辱你妈草你妈恶心的(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 18.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 24.6) = InArray((i * 2 - 16.1)) Xor (i - 4.6)
    Next i
    叼他妈操你爸狗耻辱你妈草你妈恶心的 = StrConv(tmpArray, vbUnicode)
End Function

Function 逼你爸狗儿子杂种辣鸡残缺智障你妈(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 13.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.4) = InArray((i * 2 - 27.4)) Xor (i - 4.7)
    Next i
    逼你爸狗儿子杂种辣鸡残缺智障你妈 = StrConv(tmpArray, vbUnicode)
End Function

Function 垃圾逼逼走狗狗你爸滚犊子敷衍(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 3.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 20.8) = InArray((i * 2 - 12)) Xor (i - 25.9)
    Next i
    垃圾逼逼走狗狗你爸滚犊子敷衍 = StrConv(tmpArray, vbUnicode)
End Function

Function 虚无人模狗样自爆装逼猪你爸(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 20.4) = InArray((i * 2 - 0.8)) Xor (i - 26.9)
    Next i
    虚无人模狗样自爆装逼猪你爸 = StrConv(tmpArray, vbUnicode)
End Function

Function 认知障碍脑袋有洞贱人你照照镜子看看弱智弱智叼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 23.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 6.3) = InArray((i * 2 - 28.1)) Xor (i - 1.7)
    Next i
    认知障碍脑袋有洞贱人你照照镜子看看弱智弱智叼 = StrConv(tmpArray, vbUnicode)
End Function

Function 厌恶艹艹孙子耻辱(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 24.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.5) = InArray((i * 2 - 15.3)) Xor (i - 27.2)
    Next i
    厌恶艹艹孙子耻辱 = StrConv(tmpArray, vbUnicode)
End Function

Function 艹逼逼艹厚颜无耻混账混账走狗垃圾(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 14.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 20.8) = InArray((i * 2 - 18.3)) Xor (i - 29.7)
    Next i
    艹逼逼艹厚颜无耻混账混账走狗垃圾 = StrConv(tmpArray, vbUnicode)
End Function

Function 蠢货辣鸡装逼妈的猪逼你个(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 29 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 4.7) = InArray((i * 2 - 11.7)) Xor (i - 0.5)
    Next i
    蠢货辣鸡装逼妈的猪逼你个 = StrConv(tmpArray, vbUnicode)
End Function

Function 弟弟狗脑残(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.4) = InArray((i * 2 - 1.7)) Xor (i - 21.5)
    Next i
    弟弟狗脑残 = StrConv(tmpArray, vbUnicode)
End Function

Function 虚无你哪来的勇气逼逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 16.8) = InArray((i * 2 - 21.1)) Xor (i - 22)
    Next i
    虚无你哪来的勇气逼逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 操你爸你哪来的勇气智障败家子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.9 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.6) = InArray((i * 2 - 17.3)) Xor (i - 7.4)
    Next i
    操你爸你哪来的勇气智障败家子 = StrConv(tmpArray, vbUnicode)
End Function

Function 自爆垃圾傻逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 1 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 12.9) = InArray((i * 2 - 27)) Xor (i - 28.1)
    Next i
    自爆垃圾傻逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 操你爸怂人模狗样装逼逼贱种(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 0.6) = InArray((i * 2 - 9.4)) Xor (i - 2.1)
    Next i
    操你爸怂人模狗样装逼逼贱种 = StrConv(tmpArray, vbUnicode)
End Function

Function 弱智草你妈垃圾叼逼逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 21.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.7) = InArray((i * 2 - 10.6)) Xor (i - 7.4)
    Next i
    弱智草你妈垃圾叼逼逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 装逼恶心的弟弟你照照镜子看看狗儿子你爸(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 1.6) = InArray((i * 2 - 27.2)) Xor (i - 4.4)
    Next i
    装逼恶心的弟弟你照照镜子看看狗儿子你爸 = StrConv(tmpArray, vbUnicode)
End Function

Function 脑残呵呵垃圾逼逼艹人模狗样弱智(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 20.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7) = InArray((i * 2 - 6.5)) Xor (i - 2.9)
    Next i
    脑残呵呵垃圾逼逼艹人模狗样弱智 = StrConv(tmpArray, vbUnicode)
End Function

Function 贱人艹恶心的你妈生你干啥走狗猪你呵呵(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 10.7 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 21) = InArray((i * 2 - 9.5)) Xor (i - 18.2)
    Next i
    贱人艹恶心的你妈生你干啥走狗猪你呵呵 = StrConv(tmpArray, vbUnicode)
End Function

Function 残缺智障猪(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 20.2 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 5) = InArray((i * 2 - 3.6)) Xor (i - 21.6)
    Next i
    残缺智障猪 = StrConv(tmpArray, vbUnicode)
End Function

Function 耻辱操你爸杂种悲哀装逼弱智走狗(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 28.6 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 28.1) = InArray((i * 2 - 9.5)) Xor (i - 29.9)
    Next i
    耻辱操你爸杂种悲哀装逼弱智走狗 = StrConv(tmpArray, vbUnicode)
End Function

Function 厌恶贱人厚脸皮滚犊子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 4.5 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 7.8) = InArray((i * 2 - 24.8)) Xor (i - 10.2)
    Next i
    厌恶贱人厚脸皮滚犊子 = StrConv(tmpArray, vbUnicode)
End Function

Function 傻逼呵呵蠢货蠢蠢货人模狗样(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 29.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 8.8) = InArray((i * 2 - 23.3)) Xor (i - 15.4)
    Next i
    傻逼呵呵蠢货蠢蠢货人模狗样 = StrConv(tmpArray, vbUnicode)
End Function

Function 弟弟贱种你妈逼小杂碎(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 26.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 25.2) = InArray((i * 2 - 24.1)) Xor (i - 5.1)
    Next i
    弟弟贱种你妈逼小杂碎 = StrConv(tmpArray, vbUnicode)
End Function

Function 小杂碎贱人滚犊子自爆小杂碎你妈算个鸟你妈(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 29.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 9) = InArray((i * 2 - 26.7)) Xor (i - 2.8)
    Next i
    小杂碎贱人滚犊子自爆小杂碎你妈算个鸟你妈 = StrConv(tmpArray, vbUnicode)
End Function

Function 厚颜无耻草你妈走狗你麻痹装逼喷叼弟弟(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 2.3 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.3) = InArray((i * 2 - 11.8)) Xor (i - 5.7)
    Next i
    厚颜无耻草你妈走狗你麻痹装逼喷叼弟弟 = StrConv(tmpArray, vbUnicode)
End Function

Function 智障傻逼你照照镜子看看杂种垃圾逼逼(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 14.8 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 14.7) = InArray((i * 2 - 7.3)) Xor (i - 25.2)
    Next i
    智障傻逼你照照镜子看看杂种垃圾逼逼 = StrConv(tmpArray, vbUnicode)
End Function

Function 走狗智障杂种小杂碎孙子小兔崽子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 11.1) = InArray((i * 2 - 22.6)) Xor (i - 23.4)
    Next i
    走狗智障杂种小杂碎孙子小兔崽子 = StrConv(tmpArray, vbUnicode)
End Function

Function 煞笔垃圾厚脸皮敷衍脑袋有洞小杂碎(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 27 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 3.9) = InArray((i * 2 - 27.5)) Xor (i - 4.8)
    Next i
    煞笔垃圾厚脸皮敷衍脑袋有洞小杂碎 = StrConv(tmpArray, vbUnicode)
End Function

Function 窝囊废耻辱虚无狗儿子(InArray As Variant) As String
    Dim i           As Single
    Dim tmpArray()  As Byte: ReDim tmpArray(UBound(InArray))
    
    For i = 8.4 To UBound(InArray) + 13.5 Step 1
        tmpArray(i - 24.2) = InArray((i * 2 - 2.2)) Xor (i - 12.7)
    Next i
    窝囊废耻辱虚无狗儿子 = StrConv(tmpArray, vbUnicode)
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
    NumberData = StrConv("你觉得你能用字符串搜索很牛逼对不对？然而没什么卵用。", 128 + t * t - t ^ 2)
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
    煞笔贱人小兔崽子逼逼脑袋有洞脑残 0, 0, 0, 0
    '=============================================== End
    
    'Make the array of initial proc. name
    '"操你妈了个鸡巴白痴混账脑残"
    tmp = Array(53, 94, 67, 100, 69, 111, 70, 76, 63, 113, 59, 33, 55, 74, 55, 80, 52, 82, 60, 107, 82, 76, 67, 83, 53, 87)
    StrConv Useless2(0), vbUnicode + t * t - t ^ 2                                  'Ignore this
    For t = 5.5 To UBound(tmp) + 5.5                                                'For t = 0 To Ubound(tmp)
        tmp(t - 5.5) = tmp(t - 5.5) Xor (135 + t * t - t ^ 2)                       '   tmp(t) = tmp(t) xor 135
    Next t
    ReDim ProcName(UBound(tmp) + t * t - t ^ 2)
    For t = 0 To UBound(tmp)
        ProcName(t) = tmp(t)
    Next t
    'Init. all strings (Call 操你妈了个鸡巴白痴混账脑残(Null))
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
    '"厚颜无耻你妈生你干啥傻逼你爸猪败家子小杂碎"
    tmp = Array(246, 189, 157, 153, 130, 146, 255, 144, 136, 175, 142, 164, 133, 182, 136, 175, 244, 133, 133, 250, 133, 249, 253, 138, 136, 175, 252, 154, 154, 161, 252, 144, 240, 158, 155, 159, 156, 237, 152, 159, 135, 165)
    ReDim ProcName(UBound(tmp))
    For t = 3.8 To UBound(tmp) + 3.8                                                'For t = 0 To Ubound(tmp)
        tmp(t - 3.8) = tmp(t - 3.8) Xor (76 + t * t - t ^ 2)
        ProcName(t - 3.8) = tmp(t - 3.8)                                            '   ProcName(t) = tmp(t) xor 76
    Next t
    Me.Show
    'Start loop check (Call 厚颜无耻你妈生你干啥傻逼你爸猪败家子小杂碎(Null))
    CallByName Me, StrConv(ProcName, (vbUnicode + t ^ 2) - t * t), (VbMethod + t * t) - t ^ 2, tmp
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Exiting = True
End Sub

Private Sub 是不是看到控件名觉得自己很厉害呢_Click()
     Dim tmpArray       As Variant
     Dim i              As Single
     Dim tmp            As Variant
     Dim ProcName()     As Byte
     
     tmpArray = Array(246, 362, 198, 248, 267, 125, 362, 488, 264, 216, 392, 264, 264, 488, 362, 125, 267, 248, 198, 362, 246)      'Totally rubbish XD
     tmp = Array(219, 255, 199, 195, 213, 247, 170, 215, 170, 215, 162, 205, 212, 240, 210, 253, 167, 227, 193, 204)                '"孙子喷呵呵草泥马敷衍"
     ReDim ProcName(UBound(tmp))
     For i = 1.75 To UBound(tmp) + 1.75
        ProcName(i - 1.75) = tmp(i - 1.75) Xor (16 + i * i) - i ^ 2
     Next i
     CallByName Me, StrConv(ProcName, vbUnicode), VbMethod, tmpArray                                                                'Now tmpArray(4) should be the returned value
     
     tmp = Array(55, 16, 38, 38, 38, 38, 77, 70, 36, 32, 76, 71, 76, 71, 79, 9, 37, 37, 70, 47, 70, 47, 65, 35, 59, 66)             '"你照照镜子看看贱种弟弟残缺"
     ReDim ProcName(UBound(tmp))
     For i = 3.68 To UBound(tmp) + 3.68
        ProcName(i - 3.68) = tmp(i - 3.68) Xor (243 + i * i) - i ^ 2
     Next i
     For i = 2.48 To UBound(tmpArray) + 1.48                                                                                        'The purpose is to reduce tmpArray(4) by 5
        tmpArray(i - 2.48) = tmpArray(i - 1.48) - (i - 2.48) - 2
     Next i
     CallByName Me, StrConv(ProcName, vbUnicode), VbMethod, tmpArray
End Sub
