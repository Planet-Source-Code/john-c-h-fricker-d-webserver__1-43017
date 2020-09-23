VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl Server 
   CanGetFocus     =   0   'False
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Server.ctx":0000
   ScaleHeight     =   420
   ScaleWidth      =   420
   ToolboxBitmap   =   "Server.ctx":0974
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   1
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------
'---------Written by Drew Lederman---------
'------------------------------------------

Option Explicit
'Event Declarations:
Event DataArrival(ByVal idx As Integer, ByVal Data As String, ByVal bytesTotal As Long, ByVal RemoteIP As String, ByVal RemoteHost As String)
Event Error(ByVal idx As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String)
Event SocketOpened(ByVal idx As Integer, ByVal LocalPort As Long, ByVal RemoteIP As String, ByVal RemoteHost As String, PortType As String)
Event SocketClosed(ByVal idx As Integer, ByVal LocalPort As Long, ByVal RemoteIP As String, ByVal RemoteHost As String)
Event ServerStarted()
Event ServerStopped()
Event StartFailed()
Event SendComplete(idx As Integer)
'Default Property Values:
Const m_def_State = "Closed"
'Property Variables:
Dim m_State As String

Public Property Get ServerPort() As Long
    ServerPort = Winsock(1).LocalPort
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function StartServer(LocalPort As Long, Optional LocalIP As String) As Boolean
Attribute StartServer.VB_Description = "Starts server process"
    On Error GoTo errhandle
    
    If LocalIP = "" Then LocalIP = Winsock(1).LocalIP
    
    'Open the server port
    Winsock(1).Close
    'Winsock(1).LocalPort = LocalPort
    Winsock(1).Bind LocalPort
    Winsock(1).Listen
    
    'Wait until it is open
    Do While Winsock(1).State <> sckListening: DoEvents
        If Winsock(1).State = sckError Then StartServer = False: RaiseEvent StartFailed: Exit Function
    Loop
    
    m_State = "Running"
    
    StartServer = True
    RaiseEvent ServerStarted
    Exit Function
errhandle:
    StartServer = False
    RaiseEvent StartFailed
End Function

Public Function GetStateByIdx(idx As Integer)
On Error GoTo err
GetStateByIdx = Winsock(idx).State
err:
End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0

Public Function StopServer() As Boolean
Attribute StopServer.VB_Description = "Stops server process"
    Winsock(1).Close
    
    On Error Resume Next
    Dim x As Integer
    'close each connection and unload the winsock control
    For x = 2 To Winsock.UBound
        Winsock(x).Close
        Unload Winsock(x)
    Next x
    On Error GoTo 0
    
    m_State = "Closed"
    RaiseEvent ServerStopped
End Function

Private Sub UserControl_Initialize()
    m_State = m_def_State
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 420
    UserControl.Width = 420
End Sub

Private Sub Winsock_Close(Index As Integer)
    If Index <> 0 Then
        With Winsock(Index)
        RaiseEvent SocketClosed(Index, .LocalPort, .RemoteHostIP, .RemoteHost)
        End With
        'unload when the connection is closed
        Unload Winsock(Index)
    End If
End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    On Error GoTo errhandle
    
    Dim nIndex As Long
    If Index = 1 Then
        'set nIndex to the max winsock index, so we dont get errors
        nIndex = Winsock.UBound + 1
        'load a new winsock to accept the request
        '(this is what enables mutiple connections)
        Load Winsock(nIndex)
        Winsock(nIndex).Accept (requestID)
        
        With Winsock(nIndex)
        Dim lport As String
        Select Case Winsock(1).LocalPort
        Case 23
        lport = "telnet Port"
        Case 21
        lport = "ftp port"
        Case 80
        lport = "internet port"
        Case 25
        lport = "email port"
        Case Winsock(1).LocalPort
        lport = CStr(Winsock(1).LocalPort)
        End Select
        RaiseEvent SocketOpened(nIndex, .LocalPort, .RemoteHostIP, .RemoteHost, lport)
        End With
    End If
    
    Exit Sub
errhandle:
    
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=8,1,2,0
Public Property Get ConnectionCount() As Long
   ConnectionCount = Winsock.Count - 1
End Property


Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String
    Winsock(Index).GetData strData, , bytesTotal
    'pass the data to the control's dataarrival event
    RaiseEvent DataArrival(Index, strData, bytesTotal, Winsock(Index).RemoteHostIP, Winsock(Index).RemoteHost)
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub SendData(ByVal Data As Variant, ByVal idx As Long)
Attribute SendData.VB_Description = "Send data to a  remote computer."
    On Error Resume Next
    Winsock(idx).SendData Data
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,2,
Public Property Get State() As String
Attribute State.VB_Description = "Returns the state of the server."
Attribute State.VB_MemberFlags = "400"
    State = m_State
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub CloseSocket(idx As Integer)
     On Error Resume Next
        Winsock(idx).Close
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13
Public Function GetRemoteHost(idx As Integer) As String
Attribute GetRemoteHost.VB_Description = "Returns the remote host name of the specified socket."
    On Error Resume Next
        GetRemoteHost = Winsock(idx).RemoteHost
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13
Public Function GetRemoteIP(idx As Integer) As String
Attribute GetRemoteIP.VB_Description = "Returns the remote IP address of the specified socket."
    On Error Resume Next
        GetRemoteIP = Winsock(idx).RemoteHostIP
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8
Public Function GetBytesReceived(idx As Integer) As Long
Attribute GetBytesReceived.VB_Description = "Returns  total bytes recieved on specified socket."
    On Error Resume Next
        GetBytesReceived = Winsock(idx).BytesReceived
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8
Public Function GetLocalPort(idx As Integer) As Long
Attribute GetLocalPort.VB_Description = "Returns the local port of the specified socket."
    On Error Resume Next
        GetLocalPort = Winsock(idx).LocalPort
End Function

Public Property Get ServerIP() As String
    ServerIP = Winsock(1).LocalIP
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7
Public Function GetState(idx As Integer) As Integer
Attribute GetState.VB_Description = "Returns the state of the specified socket."
    On Error Resume Next
        GetState = Winsock(idx).State

End Function

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_State = m_def_State
End Sub

Private Sub Winsock_SendComplete(Index As Integer)
RaiseEvent SendComplete(Index)
End Sub
