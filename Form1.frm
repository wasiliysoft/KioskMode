VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6270
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnReboot 
      Caption         =   "ПЕРЕЗАГРУЗКА"
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   3960
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   360
      TabIndex        =   8
      Top             =   4680
      Width           =   3855
      Begin VB.Label Label1 
         Caption         =   "..."
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.CommandButton btnLogoff 
      Caption         =   "ВЫХОД ИЗ СИТЕМЫ"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   3855
   End
   Begin VB.CommandButton btnChangePassword 
      Caption         =   "НАСТРОЙКИ"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   3855
   End
   Begin VB.CommandButton btnLock 
      Caption         =   "БЛОКИРОВАТЬ"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   5640
   End
   Begin VB.CommandButton btnUnlock 
      Caption         =   "РАЗБЛОКИРОВАТЬ"
      Default         =   -1  'True
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
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
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Пароль разблокировки"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label2 
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnReboot_Click()
    If (vbYes = MsgBox("Выполнить перезагрузку?", vbYesNo + vbQuestion + vbDefaultButton2)) Then
        On Error Resume Next
        If (Dir(Environ("windir") & "\system32\shutdown.exe") <> "") Then
             Shell "shutdown /r /f /t 0"
        Else
             Shell App.Path + "\shutdown.exe /r /f /t 0"
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub Form_Load()
    load_Config
    gTimeout = gConfig.timeout_FirstLock
    Timer1.Enabled = True
    updateLangIndicator
    setTaskMgrLockedMode True
End Sub

Private Sub btnUnlock_Click()
    If (logOn) Then
        ' Работает на WinXP но не работает на Win11 и вохмож
        ' Shell "cmd /c explorer.exe"
        
        Shell Environ("windir") & "\explorer.exe"
    End If
End Sub

Private Sub btnLock_Click()
    gTimeout = 0
    On Error Resume Next
        If (Dir(Environ("windir") & "\system32\taskkill.exe") <> "") Then
             Shell "taskkill /f /im explorer.exe"
        Else
             Shell App.Path + "\taskkill_win2000.exe -f explorer.exe"
        End If
        setTaskMgrLockedMode True
    On Error GoTo 0
End Sub

Private Sub btnChangePassword_Click()
    If (logOn) Then frmChangePass.Show
End Sub

Private Sub btnLogoff_Click()
    If (logOn) Then
        On Error Resume Next
            If (Dir(Environ("windir") & "\system32\shutdown.exe") <> "") Then
                 Shell "shutdown /l /f"
            Else
                 Shell App.Path + "\shutdown.exe /l /f"
            End If
        On Error GoTo 0
    End If
End Sub

Function logOn() As Boolean
    'If (gTimeout > 0) Then
    '     logOn = True
    '     Exit Function
    'End If
    
    Dim pass As String: pass = Trim(CStr(txtPass.Text))
    logOn = False
            
    If (Trim(gConfig.pwd) = pass) Then
        Label2.Caption = ""
        txtPass.Text = ""
        gTimeout = gConfig.timeout_ReLock
        Timer1.Enabled = True
        logOn = True
        setTaskMgrLockedMode False
    Else
        Label2.Caption = "Неправильный пароль"
        txtPass.Text = ""
        txtPass.SetFocus
    End If
End Function


Private Sub Timer1_Timer()
    If (gTimeout <= 0) Then
        btnLock_Click
        Label1.Caption = "Режим АСУ ТП"
        Timer1.Enabled = False
    Else
        gTimeout = gTimeout - 1
        Label1.Caption = "Режим АСУ ТП черех: " & gTimeout & " сек."
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Label2.Caption = ""
    updateLangIndicator
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    updateLangIndicator
End Sub

Private Sub updateLangIndicator()
    LabelLang.Caption = IIf(GetKeyboardLayout(0) = 67699721, "EN", "RU")
End Sub

