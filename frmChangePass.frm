VERSION 5.00
Begin VB.Form frmChangePass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Настройки"
   ClientHeight    =   7500
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   8730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4431.245
   ScaleMode       =   0  'User
   ScaleWidth      =   8196.996
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "О программе"
      Height          =   2175
      Left            =   360
      TabIndex        =   13
      Top             =   5040
      Width           =   4695
      Begin VB.Label LabelVersion 
         Caption         =   "Версия:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "2023 г."
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Васильченко Виталий Юрьевич"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Служба корпоративной защиты"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "ООО ""Газпром трансгаз Екатеринбург"""
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Настройка таймаутов"
      Height          =   1935
      Left            =   360
      TabIndex        =   10
      Top             =   2280
      Width           =   8055
      Begin VB.TextBox txtFirstLock 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3720
         MaxLength       =   16
         TabIndex        =   2
         Top             =   360
         Width           =   3285
      End
      Begin VB.TextBox txtReLock 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3720
         MaxLength       =   16
         TabIndex        =   3
         Top             =   1080
         Width           =   3285
      End
      Begin VB.Label lblLabels 
         Caption         =   "Повторный (сек.)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   3570
      End
      Begin VB.Label lblLabels 
         Caption         =   "Первичный (сек.)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   3810
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Изменение пароля"
      Height          =   1815
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   8055
      Begin VB.TextBox txtNewPass1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3720
         MaxLength       =   16
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   360
         Width           =   3285
      End
      Begin VB.TextBox txtNewPass2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3720
         MaxLength       =   16
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   3285
      End
      Begin VB.Label lblLabels 
         Caption         =   "Новый пароль"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   2850
      End
      Begin VB.Label lblLabels 
         Caption         =   "Подтверждение"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   3330
      End
      Begin VB.Label LabelLang 
         Caption         =   "LNG"
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
         Left            =   7200
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "СОХРАНИТЬ"
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   6720
      Width           =   2100
   End
   Begin VB.Label LabelMsg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   4440
      Width           =   8055
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const minPassLen = 10

Private Sub Form_Load()
    updateLangIndicator
    LabelVersion.Caption = "Версия: " & App.Major & "." & App.Minor & "." & App.Revision
    txtFirstLock.Text = gConfig.timeout_FirstLock
    txtReLock.Text = gConfig.timeout_ReLock
End Sub

Private Sub cmdOK_Click()
    Dim newPass As String: newPass = Trim(txtNewPass1.Text)
    If (Len(newPass) > 0) Then
        If (passwordComplexityTest(newPass)) Then
            If (newPass = Trim(txtNewPass2.Text)) Then
                gConfig.pwd = newPass
            Else
                txtNewPass2.Text = ""
                txtNewPass2.SetFocus
                LabelMsg.Caption = "Новый пароль и подтверждение не совпадают"
                Exit Sub
            End If
        Else
            txtNewPass1.SetFocus
            LabelMsg.Caption = "Новый пароль слишком простой"
            Exit Sub
        End If
    End If
    
    gConfig.timeout_FirstLock = Val(txtFirstLock.Text)
    gConfig.timeout_ReLock = Val(txtReLock.Text)
    
    If (save_Config) Then
        LabelMsg.Caption = "Настройки сохранены"
        load_Config
    End If
End Sub

Function passwordComplexityTest(ByVal psw As String) As Boolean

    Dim i As Integer, j As Integer, k As Integer
    Dim hasNum As Boolean, hasUpper As Boolean, hasLower As Boolean

    Dim complexityLvl As Integer: complexityLvl = 0
    

    'see if there is a number in the password
    'NOTE: the following For loops uses the ASCII values for numbers and letters.
    For k = 33 To 47
        If (InStr(1, psw, Chr(k))) Then
            complexityLvl = complexityLvl + 1
            Exit For
        End If
    Next k
    
    For k = 58 To 64
        If (InStr(1, psw, Chr(k))) Then
            complexityLvl = complexityLvl + 1
            Exit For
        End If
    Next k
    
    For k = 48 To 57
        If (InStr(1, psw, Chr(k))) Then
            complexityLvl = complexityLvl + 1
            Exit For
        End If
    Next k

    'See if there is an upper case
    For i = 65 To 90
        If (InStr(1, psw, Chr(i))) Then
            complexityLvl = complexityLvl + 1
            Exit For
        End If
    Next i

    'See if there is a lower case
    For j = 97 To 122
        If (InStr(1, psw, Chr(j))) Then
            complexityLvl = complexityLvl + 1
            Exit For
        End If
    Next j
    
    If (Len(psw) < minPassLen) Then
        complexityLvl = 0
    End If
    
    If complexityLvl < 3 Then
       passwordComplexityTest = False
    Else
       passwordComplexityTest = True
    End If
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    updateLangIndicator
    LabelMsg.Caption = ""
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    updateLangIndicator
End Sub

Private Sub updateLangIndicator()
    LabelLang.Caption = IIf(GetKeyboardLayout(0) = 67699721, "EN", "RU")
End Sub

