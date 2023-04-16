Attribute VB_Name = "Module2"
Option Explicit

Public gConfig As AppConfigType

'структура для конфигурационного файла (configFilePath)
Public Type AppConfigType
    pwd As String * 16   ' Пароль
End Type

Private Function configFilePath() As String
   configFilePath = App.Path + "\kiosk.config"
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
End Sub

Sub save_Config()
    Dim fh As Long: fh = FreeFile
    Open configFilePath For Random As fh Len = Len(gConfig)
        Put #fh, 1, gConfig
    Close #fh
End Sub
