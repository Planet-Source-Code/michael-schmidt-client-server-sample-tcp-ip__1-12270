Attribute VB_Name = "modPublic"
Option Explicit

Public iSock As SocketInfo
Type SocketInfo

    Address As String
    Port As Integer
    LogData As Boolean

End Type


'============================================
'   LogIO Sub - Log Text Data To FRMMAIN
'============================================
Public Sub LogIO(TextToLog As String, Optional Force As Boolean = False)

    If iSock.LogData Or Force Then frmMain.txtLog = frmMain.txtLog & vbCrLf & TextToLog
    If Force = True Then Beep

End Sub


'============================================
'   Registry - Load Settings
'============================================
Public Sub LoadAppSettings()
On Error Resume Next

    ' Load User Settings
    iSock.Address = GetSetting(App.Title, "SETTINGS", "LASTIP", "127.0.0.1")
    iSock.Port = GetSetting(App.Title, "SETTINGS", "LASTPORT", 1251)
    iSock.LogData = GetSetting(App.Title, "SETTINGS", "LOGDATA", False)

End Sub


'============================================
'   Registry - Save Settings
'============================================
Public Sub SaveAppSettings()

    ' Load User Settings
    SaveSetting App.Title, "SETTINGS", "LASTIP", iSock.Address
    SaveSetting App.Title, "SETTINGS", "LASTPORT", iSock.Port
    SaveSetting App.Title, "SETTINGS", "LOGDATA", iSock.LogData

End Sub
