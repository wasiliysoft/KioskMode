Attribute VB_Name = "Module2"
Option Explicit

Public gConfig As AppConfigType

'структура для конфигурационного файла (configFilePath)
Public Type AppConfigType
    pwd As String * 16   ' Пароль
    timeout_FirstLock As Long
    timeout_ReLock As Long
End Type

Private Function configFilePath() As String
   configFilePath = App.Path + "\config.bin"
End Function

Sub load_Config()
    Dim fh As Long: fh = FreeFile
    Dim fLen As Long
    On Error Resume Next
        fLen = FileLen(configFilePath)
    On Error GoTo 0
    If fLen = 0 Then
        With gConfig
            .pwd = "SKZ"
        End With
        MsgBox "Отсутвует файл конфигурации: " & configFilePath, vbExclamation
    Else
        Open configFilePath For Random As fh Len = Len(gConfig)
            Get #fh, 1, gConfig
        Close #fh
    End If
    
    load_custom_timeout
    
    
End Sub

Private Sub load_custom_timeout()
    Dim s As String: s = App.Path + "\custom_timeout.txt"
    Dim timeout_FirstLock As Long: timeout_FirstLock = 15
    Dim timeout_ReLock As Long: timeout_ReLock = 600
    
    Dim fh As Long: fh = FreeFile
    Dim fLen As Long
    On Error Resume Next
        fLen = FileLen(s)
    On Error GoTo 0
    
    If fLen > 0 Then
        Open s For Input Access Read As fh
            Seek #fh, 1
            Line Input #fh, s
            timeout_FirstLock = CLng(s)
            Line Input #fh, s
            timeout_ReLock = CLng(s)
        Close #fh
    End If
    
    If timeout_FirstLock < 0 Then timeout_FirstLock = 0
    If timeout_ReLock < 0 Then timeout_ReLock = 0
    
    If timeout_FirstLock > 600 Then timeout_FirstLock = 600
    If timeout_ReLock > 1800 Then timeout_ReLock = 1800
    
    With gConfig
        .timeout_FirstLock = timeout_FirstLock
        .timeout_ReLock = timeout_ReLock
    End With
End Sub
Sub save_Config()
    Dim fh As Long: fh = FreeFile
    Open configFilePath For Random As fh Len = Len(gConfig)
        Put #fh, 1, gConfig
    Close #fh
End Sub
