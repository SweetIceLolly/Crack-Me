VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   14235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   8760
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Data()  As String

Private Function FuncName(InArray() As Byte) As String
    Dim i           As Single
    Dim tmpArray()  As Byte
    Dim VarName     As VarType
    
    For i = Rnd To rnd2 Step rnd3
        tmpArray(i - rnd4) = InArray((i * 2 - rnd5)) Xor (i - rnd6)
    Next i
    FuncName = StrConv(tmpArray, vbUnicode)
End Function

Private Sub Command1_Click()
    Dim n   As Integer
    Dim i   As Integer
    Dim s   As String
    Dim c   As String
    Dim t   As String
    Dim j   As Integer
    
    Open "C:\code.txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, t
            c = c & t & vbCrLf
        Loop
    Close #1
    
    Randomize
    Open "C:\1.txt" For Append As #1
        For j = 1 To 200
            s = ""
            n = 5 * Rnd + 3
            For i = 1 To n
                s = s & Data(UBound(Data) * Rnd)
            Next i
            Print #1, Replace(Replace(Replace(Replace(Replace(Replace(Replace(c, "¡¾FuncName¡¿", s), "¡¾Rnd1¡¿", CInt(300 * Rnd)), "¡¾Rnd2¡¿", CInt(300 * Rnd)), "¡¾Rnd3¡¿", CInt(300 * Rnd)), "¡¾Rnd4¡¿", CInt(300 * Rnd)), "¡¾Rnd5¡¿", CInt(300 * Rnd)), "¡¾Rnd6¡¿", CInt(300 * Rnd))
        Next j
    Close #1
End Sub

Private Sub Form_Load()
    ReDim Data(0)
    Open "C:\list.txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, Data(UBound(Data))
            ReDim Preserve Data(UBound(Data) + 1)
        Loop
    Close #1
    
End Sub
