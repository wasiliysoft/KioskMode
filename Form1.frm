VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ËÓÒÍ"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnReboot 
      Caption         =   "œ≈–≈«¿√–”«»“‹"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   4800
      Width           =   3855
   End
   Begin VB.CommandButton btnLogoff 
      Caption         =   "«¿¬≈–ÿ»“‹ —≈¿Õ—"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   4080
      Width           =   3855
   End
   Begin VB.CommandButton btnChangePassword 
      Caption         =   "»«Ã≈Õ»“‹ œ¿–ŒÀ‹"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   3855
   End
   Begin VB.CommandButton btnLock 
      Caption         =   "¡ÀŒ »–Œ¬¿“‹"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   3600
   End
   Begin VB.CommandButton btnUnlock 
      Caption         =   "–¿«¡ÀŒ »–Œ¬¿“‹"
      Default         =   -1  'True
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   3855
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
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label LabelLang 
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "œ‡ÓÎ¸ ‡Á·ÎÓÍËÓ‚ÍË ‡·Ó˜Â„Ó ÒÚÓÎ‡"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label2 
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "..."
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   5640
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


Private Sub btnLogoff_Click()
    If (logOn(Trim(CStr(txtPass.Text)))) Then
        onCorrectPass
        Shell "LOGOFF"
    Else
        onIncorrectPass
    End If
End Sub

Private Sub btnReboot_Click()
    If (logOn(Trim(CStr(txtPass.Text)))) Then
        onCorrectPass
        Shell "shutdown -r -t 0 -f"
    Else
        onIncorrectPass
    End If
End Sub

Private Sub Form_Load()
    load_Config
    gTimeout = gConfig.timeout_FirstLock
    updateLangIndicator
End Sub

Private Sub btnUnlock_Click()
    If (logOn(Trim(CStr(txtPass.Text)))) Then
        onCorrectPass
        Shell "explorer.exe"
    Else
        onIncorrectPass
    End If
End Sub
Private Sub onCorrectPass()
        Label2.Caption = ""
        txtPass.Text = ""
        gTimeout = gConfig.timeout_ReLock
        Timer1.Enabled = True
End Sub
Private Sub onIncorrectPass()
        Label2.Caption = "ÕÂÔ‡‚ËÎ¸Ì˚È Ô‡ÓÎ¸"
        txtPass.SetFocus
        SendKeys "{Home}+{End}"
End Sub
Private Sub updateLangIndicator()
    LabelLang.Caption = IIf(GetKeyboardLayout(0) = 67699721, "EN", "RU")
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    updateLangIndicator
    If ((KeyCode = vbKeyEnd) And (Shift = 6)) Then
        If (isLocked = False) Then
   '         Unload Me
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    updateLangIndicator
End Sub


