Attribute VB_Name = "Module1"
Option Explicit

Public gPswd As String
Public Const timeout = 60
Public gTimeout As Long
Public Const minPassLen = 10

Sub loadConfig()
    gPswd = Trim("2")
End Sub
Sub saveConfig()
    MsgBox "не рализованно"
End Sub
Function logOn(ByVal pass As String) As Boolean
    logOn = False
    pass = Trim(pass)
    If (gPswd = pass) Then
        logOn = True
    End If
End Function

Sub logOff()
    gTimeout = 0
    Shell "taskkill /f /im explorer.exe"
End Sub


Function isLocked() As Boolean
    isLocked = gTimeout <= 0
End Function

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


    

