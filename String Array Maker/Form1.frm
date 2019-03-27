VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim bArray()    As Byte
    Dim i           As Integer
    Dim s           As String
    
    bArray = StrConv(Me.Text1.Text, vbFromUnicode)
    For i = 0 To UBound(bArray)
        bArray(i) = bArray(i) Xor i
    Next i
    
    For i = 0 To UBound(bArray)
        s = s & bArray(i) & ", "
    Next i
    
    Me.Text2.Text = s
End Sub
