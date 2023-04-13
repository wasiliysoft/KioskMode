VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "KIOSK"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnChangePassword 
      Caption         =   "»«Ã≈Õ»“‹ œ¿–ŒÀ‹"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   3495
   End
   Begin VB.CommandButton btnLock 
      Caption         =   "¡ÀŒ »–Œ¬¿“‹"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   4200
   End
   Begin VB.CommandButton btnUnlock 
      Caption         =   "–¿«¡ÀŒ »–Œ¬¿“‹"
      Height          =   615
      Left            =   360
      TabIndex        =   1
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
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "..."
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3720
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnChangePassword_Click()
    frmChangePass.Show
End Sub

Private Sub btnLock_Click()
    logOff
End Sub

Private Sub Form_Load()
    loadConfig
    gTimeout = 5
End Sub

Private Sub btnUnlock_Click()
    If (logOn(Trim(CStr(txtPass.Text)))) Then
        Label2.Caption = ""
        txtPass.Text = ""
        gTimeout = timeout
        Timer1.Enabled = True
        Shell "explorer.exe"
    Else
        Label2.Caption = "ÕÂÔ‡‚ËÎ¸Ì˚È Ô‡ÓÎ¸"
        txtPass.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Timer1_Timer()
    If (isLocked) Then
        Label1.Caption = "–ÂÊËÏ  »Œ— "
        logOff
        Timer1.Enabled = False
    Else
        gTimeout = gTimeout - 1
        Label1.Caption = "–ÂÊËÏ  »Œ—  ˜ÂÂı: " & gTimeout & " ÒÂÍ."
    End If
End Sub

