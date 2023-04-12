VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnLock 
      Caption         =   "БЛОКИРОВАТЬ"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3600
      Top             =   3120
   End
   Begin VB.CommandButton btnUnlock 
      Caption         =   "РАЗБЛОКИРОВАТЬ"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   3495
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox txtLogin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "..."
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const timeout = 60
Private leftTimeSec As Long

Private Sub btnLock_Click()
    leftTimeSec = 0
End Sub

Private Sub Form_Load()
    loadConfig
    leftTimeSec = 5
End Sub

Private Sub btnUnlock_Click()
   result = logOn(txtLogin.Text, txtPass.Text)
    If (result) Then
        txtPass.Text = ""
        leftTimeSec = timeout
        Shell "explorer"
    Else
        MsgBox "Неправильное имя пользователя или пароль", vbExclamation
        txtPass.SetFocus
    End If
End Sub

Private Sub Timer1_Timer()
    If (leftTimeSec <= 0) Then
        logOff
        Label1.Caption = "Режим КИОСК"
    Else
        leftTimeSec = leftTimeSec - 1
        Label1.Caption = "Режим КИОСК черех: " & leftTimeSec & " сек."
    End If
End Sub
