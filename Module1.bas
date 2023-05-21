Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Public gTimeout As Long
Public gConfig As AppConfigType

'структура для конфигурационного файла (configFilePath)
Public Type AppConfigType
    pwd As String
    timeout_FirstLock As Long
    timeout_ReLock As Long
End Type

Private Function configFilePath() As String
   configFilePath = App.Path + "\config.bin"
End Function

Public Sub load_Config()
    ' Конфигурация по умолчанию
    gConfig.pwd = "111"
    gConfig.timeout_FirstLock = 15
    gConfig.timeout_ReLock = 15 * 60
    
    ' Проверка файла конфигурации
    Dim fh As Long: fh = FreeFile
    Dim fLen As Long
    On Error Resume Next
        fLen = FileLen(configFilePath)
    On Error GoTo 0
    If fLen = 0 Then
        MsgBox "Отсутвует файл конфигурации: " & configFilePath, vbExclamation
    Else
        ' Загрузка конфигурации
        Dim s As String
        Open configFilePath For Input Access Read As #fh
            Seek #fh, 1
            Line Input #fh, s
            gConfig.pwd = s
            Line Input #fh, s
            gConfig.timeout_FirstLock = Val(s)
            Line Input #fh, s
            gConfig.timeout_ReLock = Val(s)
        Close #fh
    End If
End Sub

Function save_Config() As Boolean
    save_Config = False
    Dim fh          As Long: fh = FreeFile
    Dim s As String
    On Error GoTo errorHandler
        Open configFilePath For Output As fh
            Print #fh, gConfig.pwd
            Print #fh, gConfig.timeout_FirstLock
            Print #fh, gConfig.timeout_ReLock
        Close #fh
    On Error GoTo 0
    save_Config = True
    Exit Function
    
errorHandler:
    MsgBox Err.Number & ": " & Err.Description, vbCritical
End Function

