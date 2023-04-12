Attribute VB_Name = "Module1"

Private users(10) As User

Public Type User
    login As String * 24    ' Логин
    pwd As String * 16      ' Пароль
End Type

Sub loadConfig()
    Dim u0 As User
    u0.login = "1"
    u0.pwd = "2"
    
    users(0) = u0
    MsgBox users(0).login
End Sub

Function logOn(ByVal login As String, ByVal pass As String) As Boolean
    logOn = False
    login = Trim(CStr(login))
    pass = Trim(CStr(pass))
    
    For i = 0 To 9
        If (Trim(users(i).login) = login) Then
            If (Trim(users(i).pwd) = pass) Then
                logOn = True
                Exit For
            End If
        End If
    Next
End Function

Sub logOff()
    Shell "taskkill /f /im explorer.exe"
End Sub
