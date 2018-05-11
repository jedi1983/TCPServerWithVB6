VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form serverForm 
   Caption         =   ".: Server Proto :."
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox logList 
      Height          =   1425
      Left            =   5400
      TabIndex        =   13
      Top             =   600
      Width           =   6855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Received"
      Height          =   3735
      Left            =   7800
      TabIndex        =   12
      Top             =   2280
      Width           =   4575
      Begin VB.TextBox receivedText 
         Height          =   3135
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   15
         Top             =   360
         Width           =   4095
      End
   End
   Begin MSWinsockLib.Winsock SocketWin 
      Index           =   0
      Left            =   7800
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox targetIP 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton sendCmd 
      Caption         =   "SEND"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox messageSend 
      Height          =   1935
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   3480
      Width           =   3615
   End
   Begin VB.ListBox clientList 
      Height          =   3570
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox portText 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "8000"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server Property"
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton stopBtn 
         Caption         =   "Stop Listening"
         Height          =   375
         Left            =   3480
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton startBtn 
         Caption         =   "Start Listening"
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox maxPort 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   "5"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "MAX CLIENT :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "PORT :"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame sendFrame 
      Caption         =   "Send Data"
      Height          =   3735
      Left            =   3240
      TabIndex        =   10
      Top             =   2280
      Width           =   4215
      Begin VB.Label Label3 
         Caption         =   "Target IP :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Log Box :"
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "serverForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Global Variable
' Roland 2018
Dim socketPort As Integer
Dim maxClient As Integer
Dim sServerMsg As String
Dim iSockets As Integer
Dim sRequestID As String
Dim indexSend As Integer
    
Private Sub clientList_Click()
    targetIP.Text = clientList.List(clientList.ListIndex)
    sendCmd.Enabled = True
End Sub

Private Sub Form_Load()

socketPort = portText.Text
maxClient = maxPort.Text
stopBtn.Enabled = False
sendCmd.Enabled = False

End Sub

Private Sub sendCmd_Click()
    ' check IP
    Call findIPToSend
    
    SocketWin(indexSend).SendData (messageSend.Text)
    
    'empty box
    
    targetIP.Text = ""
    messageSend.Text = ""
    
    sendCmd.Enabled = False
    
End Sub

'force close socket because Max Reached

Private Sub forceCloseSocket()

SocketWin(iSockets).Close
Unload SocketWin(iSockets)

' update list of Connected

End Sub

' Socket Closed
Private Sub SocketWin_Close(Index As Integer)
    
    Dim ipRemoved As String
    Dim dataList As Integer
    
    dataList = 0
    
    ipRemoved = ""
    
    ipRemoved = SocketWin(Index).RemoteHostIP
    SocketWin(Index).Close
    'Unload SocketWin(Index)
    
    dataList = clientList.ListCount
    
    For x = 0 To dataList Step 1
        
        If clientList.List(x) = ipRemoved Then
            'remove
            clientList.RemoveItem (x)
        End If
        
    Next
        
    'refresh sending box, eliminate historical data send
    
    targetIP.Text = ""
    messageSend.Text = ""
    
    
End Sub

Private Sub findIPToSend()
Dim ObjectCount As Integer
ObjectCount = 0
indexSend = 0
' check whether there are already selected IP
If targetIP.Text = "" Then
sta = MsgBox("Please Fill Out Target IP", vbOKOnly, "Target IP")
End If

' Find Out IP

For Each Sock In SocketWin
    If Sock.State = sckConnected Then
        If Sock.RemoteHostIP = targetIP.Text Then
            indexSend = Sock.Index
        End If
    End If
Next Sock

End Sub

'connection Request
Private Sub SocketWin_ConnectionRequest(Index As Integer, ByVal requestID As Long)
  If Index = 0 Then
    Dim IpRequester As String
    Dim indexSocket As Integer
    IpRequester = ""
    indexSocket = 0
    sRequestID = requestID
    
    ' check IP had connected before?
    
    IpRequester = SocketWin(0).RemoteHostIP
    For Each Sock In SocketWin
    If Sock.State = sckClosed Then
        If Sock.RemoteHostIP = IpRequester Then
            indexSocket = Sock.Index
        End If
    End If
    Next Sock
    
    ' end
    If indexSocket = 0 Then
        iSockets = SocketWin.UBound + 1
        Load SocketWin(iSockets)
        SocketWin(iSockets).LocalPort = 8000
        SocketWin(iSockets).Accept requestID
        clientList.AddItem (SocketWin(iSockets).RemoteHostIP)
    Else
        'Load SocketWin(indexSocket)
        SocketWin(indexSocket).LocalPort = 8000
        SocketWin(indexSocket).Accept requestID
        clientList.AddItem (SocketWin(indexSocket).RemoteHostIP)
    End If

    ' check max Reached
    If iSockets > maxClient Then
        Call forceCloseSocket
    End If
  End If
End Sub

' Data Arrival / Receive Socket

Private Sub SocketWin_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    Dim ipString As String
    Dim indexReceive As Integer
    Dim getDataS As String
    Dim stringWrite As String
    
    stringWrite = ""
    indexReceive = 0
    ipString = ""
    getDataS = ""

    ipString = SocketWin(Index).RemoteHostIP
    SocketWin(Index).GetData getDataS
    stringWrite = "Data From : " & ipString & " Message : " & getDataS & vbCrLf
    receivedText.Text = receivedText.Text + stringWrite
            
End Sub


Private Sub startBtn_Click()

socketPort = portText.Text
maxClient = maxPort.Text

If socketPort = 0 Then
    sta = MsgBox("Please Fill In Port Settings", vbOKOnly, "Settings")
ElseIf socketPort < 0 Then
    sta = MsgBox("Port Settings Need To Be Bigger Than 0", vbOKOnly, "Settings")
ElseIf socketPort > 65500 Then
    sta = MsgBox("Port Settings Need To Be Smaller Than 65500", vbOKOnly, "Settings")
End If

' Connect

SocketWin(0).LocalPort = socketPort
SocketWin(0).Listen
sServerMsg = "Listening To Port: " & SocketWin(0).LocalPort
logList.AddItem (sServerMsg)

startBtn.Enabled = False
stopBtn.Enabled = True

End Sub

Private Sub stopBtn_Click()

    For Each Sock In SocketWin
        If Sock.State = sckConnected Then
            Sock.Close
        End If
    Next Sock
    
    sServerMsg = "Closing Connection on Port: " & SocketWin(0).LocalPort
    logList.AddItem (sServerMsg)
    ' close socket 0
    
    SocketWin(0).Close
    
    startBtn.Enabled = True
    
    clientList.Clear

End Sub
