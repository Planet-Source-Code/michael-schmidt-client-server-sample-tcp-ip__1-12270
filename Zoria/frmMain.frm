VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zoria"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5505
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTerminate 
      Caption         =   "Terminate"
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      Top             =   450
      Width           =   1155
   End
   Begin VB.Timer timSock 
      Interval        =   500
      Left            =   510
      Top             =   4080
   End
   Begin MSWinsockLib.Winsock sckSys 
      Left            =   30
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   405
      Left            =   30
      TabIndex        =   6
      Top             =   3600
      Width           =   5445
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   3570
         TabIndex        =   11
         Top             =   150
         Width           =   90
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "On Port:"
         Height          =   195
         Left            =   2910
         TabIndex        =   10
         Top             =   150
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Remote Connection:"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   150
         Width           =   1470
      End
      Begin VB.Label lblRemote 
         AutoSize        =   -1  'True
         Caption         =   "0.0.0.0"
         Height          =   195
         Left            =   1590
         TabIndex        =   8
         Top             =   150
         Width           =   540
      End
      Begin VB.Shape shpState 
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   5220
         Shape           =   3  'Circle
         Top             =   180
         Width           =   135
      End
      Begin VB.Shape shpState 
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   5010
         Shape           =   3  'Circle
         Top             =   180
         Width           =   195
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3045
      Left            =   30
      TabIndex        =   5
      Top             =   540
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   5371
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Log"
      TabPicture(0)   =   "frmMain.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkLogData"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtLog"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdResetLog"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Chat"
      TabPicture(1)   =   "frmMain.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtDataSend"
      Tab(1).Control(1)=   "txtChat"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "System"
      TabPicture(2)   =   "frmMain.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Server"
      TabPicture(3)   =   "frmMain.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.CommandButton txtDataSend 
         Caption         =   "Send"
         Height          =   315
         Left            =   -70650
         TabIndex        =   16
         Top             =   840
         Width           =   1035
      End
      Begin VB.TextBox txtChat 
         BackColor       =   &H00730009&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -74880
         TabIndex        =   15
         Text            =   "Sample of sending data over tcp/ip. once connected, everything's easy"
         Top             =   450
         Width           =   5235
      End
      Begin VB.CommandButton cmdResetLog 
         Caption         =   "Reset"
         Height          =   255
         Left            =   90
         TabIndex        =   14
         Top             =   2700
         Width           =   1365
      End
      Begin VB.TextBox txtLog 
         BackColor       =   &H00730009&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2205
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   420
         Width           =   5235
      End
      Begin VB.CheckBox chkLogData 
         Caption         =   "Enable Log"
         Height          =   255
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2700
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Establish"
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Top             =   150
      Width           =   1155
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      BackColor       =   &H00731100&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2910
      TabIndex        =   2
      Text            =   "1251"
      Top             =   150
      Width           =   645
   End
   Begin VB.TextBox txtRemote 
      Alignment       =   2  'Center
      BackColor       =   &H00731100&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   150
      Width           =   2355
   End
   Begin VB.Label Label2 
      Caption         =   "- Port"
      Height          =   225
      Left            =   3630
      TabIndex        =   3
      Top             =   180
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "- IP"
      Height          =   225
      Left            =   2520
      TabIndex        =   1
      Top             =   180
      Width           =   285
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'============================================
'   TCP / IP Connection Sample
'============================================
'============================================
'   Rather bored, so I wrote this little
'   app, it listens for a connection on a
'   random port, at the same time, you can
'   tell it to connect to a listening
'   app. The part I like is the timer and
'   simple gfx telling you your connection
'   state...
'
'   Michael A.Schmidt
'   Written in October, 2000
'
'============================================
'   I'll probably use it as a connection
'   from my home to office, since we run
'   a firewall and can't use the UDP
'   protocol, so I'll base future apps
'   off this...
'============================================
'   To use, simply run two instances of
'   Zoria, then tell one the other's port,
'   and hit ESTABLISH, you should connect!
'============================================
'============================================
Enum Color
    green = &HFF00&
    red = &HFF&
    off = &H80&
End Enum


Private Sub cmdResetLog_Click()
    
    txtLog = "Log " & Now

End Sub

'============================================
'   Form - Load
'============================================
Private Sub Form_Load()

    LoadAppSettings

    txtPort = iSock.Port
    txtRemote = iSock.Address
    chkLogData.Value = Abs(iSock.LogData)
    
    txtLog = "Log " & Now

End Sub


'============================================
'   Form - UnLoad
'============================================
Private Sub Form_UnLoad(Cancel As Integer)

    SaveAppSettings

End Sub

'============================================
'   CheckBox - LogData
'============================================
Private Sub chkLogData_Click()
    iSock.LogData = CInt(chkLogData.Value)
End Sub


'============================================
'   Button - Connect
'============================================
Private Sub cmdConnect_Click()
    
    sckSys.Close
    DoEvents
    sckSys.Connect txtRemote, txtPort
    LogIO "Connecting " & txtRemote & ":" & txtPort

End Sub


'============================================
'   Button - Terminate
'============================================
Private Sub cmdTerminate_Click()

    sckSys.Close
    LogIO "Terminating Connection"

End Sub


'============================================
'   Timer - Socket Status
'============================================
Private Sub timSock_Timer()
On Error GoTo ErrSub
Dim SockState As String

    lblRemote = "0.0.0.0"
    lblPort = 0
    shpState(0).FillColor = off
    shpState(1).FillColor = red

    Select Case sckSys.State
        Case 7:         ' Connected
        shpState(0).FillColor = green
        shpState(1).FillColor = off
        lblRemote = sckSys.RemoteHostIP
        lblPort = sckSys.RemotePort
        Case 0, 8, 9:   ' Closed, Peer Closed, Error - Listen
        sckSys.Close
        DoEvents
        sckSys.LocalPort = iSock.Port
        sckSys.Listen
        Case Else       ' Busy, No Connect
    End Select

    Me.Caption = App.Title & " - (" & sckSys.LocalIP & ":" & sckSys.LocalPort & ") - " & SocketState(sckSys.State)

ErrSub:
    Select Case Err.Number
        Case 10048:
        sckSys.Close
        sckSys.LocalPort = 0
        sckSys.Listen
        txtPort = sckSys.LocalPort
        LogIO "Address In Use, Switching To Port " & sckSys.LocalPort, True
        Resume Next
        
    End Select

End Sub


'============================================
'   Function - Connection State
'============================================
Private Function SocketState(SCKSTATE As Integer) As String
    
    Select Case SCKSTATE
        Case 0: SocketState = "Closed"
        Case 1: SocketState = "Open"
        Case 2: SocketState = "Listening"
        Case 3: SocketState = "Pending"
        Case 4: SocketState = "Resolving Host"
        Case 5: SocketState = "Host Resolved"
        Case 6: SocketState = "Connecting"
        Case 7: SocketState = "Connected"
        Case 8: SocketState = "Peer Closed"
        Case 9: SocketState = "Error"
    End Select

End Function


Private Sub txtDataSend_Click()
    
    
    If sckSys.State = 7 Then sckSys.SendData txtChat _
    Else: MsgBox "Not Connected!", vbInformation
    
    
    
End Sub

'============================================
'   Text - Port Change
'============================================
Private Sub txtPort_Change()
    iSock.Port = txtPort
End Sub


'============================================
'   Text - Remote Change
'============================================
Private Sub txtRemote_Change()
    iSock.Address = txtRemote
End Sub


'============================================
'   Winsock - Connection Request
'============================================
Private Sub sckSys_ConnectionRequest(ByVal requestID As Long)


    sckSys.Close
    Debug.Print sckSys.State
    sckSys.Accept requestID
    
    LogIO "Connected!"
      
End Sub


'============================================
'   Winsock - Error
'============================================
Private Sub sckSys_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    ' Winsock Error - Write to log and beep.
    LogIO "####Socket Error#### " & vbCrLf & Err.Number & " - " & Err.Description & vbCrLf & _
          "#################### ", True

End Sub


'============================================
'   Winsock - Data Arrival
'============================================
Private Sub sckSys_DataArrival(ByVal bytesTotal As Long)
Dim iData As String
    
    sckSys.GetData iData        ' Pull Data From Buffer

    LogIO iData, True
    
End Sub
