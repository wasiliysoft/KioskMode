VERSION 5.00
Begin VB.Form frmChangePass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Смена пароля"
   ClientHeight    =   3915
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2313.111
   ScaleMode       =   0  'User
   ScaleWidth      =   6633.652
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3600
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1560
      Width           =   3285
   End
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
      Left            =   3600
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   3285
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ИЗМЕНИТЬ"
      Default         =   -1  'True
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   6780
   End
   Begin VB.TextBox txtOldPass 
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
      Left            =   3600
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   3285
   End
   Begin VB.Label LabelError 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   6735
   End
   Begin VB.Label lblLabels 
      Caption         =   "Подтверждение пароля"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   3330
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
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   2850
   End
   Begin VB.Label lblLabels 
      Caption         =   "Текущий пароль"
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
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2850
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    If (logOn(Trim(txtOldPass.Text))) Then
            If (passwordComplexityTest(Trim(txtNewPass1.Text))) Then
                If (Trim(txtNewPass1.Text) = Trim(txtNewPass2.Text)) Then
                    
                    
                    LabelError.Caption = "Пароль успешно изменен"
                    saveConfig
                    loadConfig
                Else
                    LabelError.Caption = "Новый пароль и подтверждение не совпадают"
                End If
            Else
                txtNewPass1.SetFocus
                SendKeys "{Home}+{End}"
                LabelError.Caption = "Новый пароль слишком простой"
            End If
    Else
        LabelError.Caption = "Неправильный текущий пароль"
        txtOldPass.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub


