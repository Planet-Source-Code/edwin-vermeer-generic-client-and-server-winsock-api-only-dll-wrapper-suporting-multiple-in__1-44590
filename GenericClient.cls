VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenericClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Purpose:
' The GenericClient object can be used to create any sort of client application.
' This object will hide the difficulties in the Winsock API calls for your.<BR><BR>
' You will probably write your own protocol handler (like SMTP, POP or ...) in the OnDataArive event.<BR>
Option Explicit

'Some variables that are used to store the state of this object.
Private m_intConnectionPort As Integer    'Just remember the port that we connected to
Private m_lngConnectionSocket As Long     'Just remember the socket where the connection is set up for
Private m_objConnection As Connection  'When we are connected to the server this object will be created for handeling this connection.

'We need a memory block for resolving the hostname. these variables will be used for handeling those.
Private lngMemoryHandle    As Long      'Handle of the allocated memory block object
Private lngMemoryPointer   As Long      'Address of the memory block

'We are going to subclass a generated window and add 2 new message numbers for the communication
Private objSuperClass As Subclass               'This is the subclass handler
Implements WinSubHook.iSubclass                  'This is the interface to the subclass handler


Private lngWinsockMessage As Long          'Winsock resolve host message
Private lngResolveHostMessage As Long   'General Winsock message
Private lngWindowHandle As Long            'Window handle

'Purpose: This event will be called when the connection is closed by the server or the client itself.
Public Event OnClose()
'Purpose: This event will be called when the connect method executed successfully.
Public Event OnConnect()
'Purpose: This event will be called when data arived. After this you have to call the .Connection.Recv method to get the data.
Public Event OnDataArrive()
'Purpose: This event will be called whenever there is a problem.
Public Event OnError(lngRetCode As Long, strDescription As String)



'Purpose:
' The moment an instance of this object is created we can set up a new socket for communication.
Private Sub Class_Initialize()
On Error GoTo ErrorHandler
Dim WSAInfo As WSAData
    
   'Create a blank, invisible window
0             lngWindowHandle = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)

   'Create 2 unique message numbers for our messages
1             lngWinsockMessage = RegisterWindowMessage(App.EXEName & ".WinsockMessage")
2             lngResolveHostMessage = RegisterWindowMessage(App.EXEName & ".ResolveHostMessage")
    
   'Start up the subclassing for the new blank invisible window. This will only be done for the 2 new messages
3             Set objSuperClass = New Subclass        'Create a cSuperClass instance
4             With objSuperClass
5               .AddMsg lngWinsockMessage, MSG_AFTER
6               .AddMsg lngResolveHostMessage, MSG_AFTER
7               .Subclass lngWindowHandle, Me
8             End With
   'Later on when we start the connection we will tell the created winsock to
   'post it's messages to the above subclassed window.
     
   'Start version 1.1 of the Microsoft Winsock API
9             WSAStartup &H101, WSAInfo
    
   'Since we only need one and the chance is big we are going to connect, just create the connection object now
10             Set m_objConnection = New Connection
    
11 Exit Sub
ErrorHandler:
12    Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in Class_Initialize on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub



'Purpose:
' When this object is closed we just make sure that we close the socket, clean up the connection and stop the subclassing.
Private Sub Class_Terminate()
On Error GoTo ErrorHandler
    
   'Cleanup Winsock
13             WSACleanup
        
   'clean up the connection object
14             Set m_objConnection = Nothing
    
   'Stop subclassing
15             Set objSuperClass = Nothing  'Destroy the subclasser
   
   'Also remove the window
16             DestroyWindow lngWindowHandle
   
17 Exit Sub
ErrorHandler:
18    Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in Class_Terminate on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub



'Purpose:
' This method will do the actual connect to the server.
' Before sending data make sure that te onConnect event was fired (and onDisconnect not).
' This method will return before the actuall connection was set up.
Public Function Connect(ByVal strRemoteHost As String, ByVal intRemotePort As Integer) As Long
On Error GoTo ErrorHandler
Dim lngSocket As Long

    'Create a new socket
19              lngSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
20              m_objConnection.Socket = lngSocket
    
    'Let the socket send the messages to the subclassed window handle (see class_initialize)
21              WSAAsyncSelect lngSocket, lngWindowHandle, lngWinsockMessage, FD_CONNECT Or FD_READ Or FD_CLOSE Or FD_ACCEPT
22              Connect = lngSocket
    
    'Resolve the host name and if that was successful then give the socket a connect command.
23              ResolveHost strRemoteHost
    
    'Save the connection parameters
24              m_intConnectionPort = intRemotePort
25              m_lngConnectionSocket = lngSocket
    
26 Exit Function
ErrorHandler:
27    Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in Connect on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpose:
' We can access the connection object from code with this property.
' This will let you use a sintax like GenericClient.Connection.Send
Public Property Get Connection() As Connection
On Error GoTo ErrorHandler

28            Set Connection = m_objConnection

29 Exit Property
ErrorHandler:
30    Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in Connection on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Property



'Purpuse:
' This is where the winsock messages are handeled.
Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
On Error GoTo ErrorHandler
'wParam is the socket
'lParam is the winsock message
'uMsg is the windows message type
Dim lngErrorCode As Long

   'was an error passed on
31             lngErrorCode = HiWord(lParam)
32             If lngErrorCode > 0 Then
33                RaiseEvent OnError(lngErrorCode, GetErrorDescription(lngErrorCode))
34             End If
   
   'Was it a winsock or resolvehost message?
35             Select Case uMsg
   Case lngWinsockMessage
      'What winsock message was posted
36                Select Case lParam
      Case FD_CONNECT     'The connection is established
37                   RaiseEvent OnConnect
38                   lReturn = 0
      Case FD_READ        'We are ready to read from the winsock buffer

            
            ' We needed this as a solution to handle faster incomming data than what we can process
39              Static intBufferCounter As Long
40              intBufferCounter = intBufferCounter + 1
41              If intBufferCounter > 1 Then Exit Sub
ContinueWithBuffer:
42              RaiseEvent OnDataArrive
43              intBufferCounter = intBufferCounter - 1
44              If intBufferCounter > 0 Then GoTo ContinueWithBuffer


45                   lReturn = 0
      Case FD_CLOSE       'The other end point wishes to close the connection
46                   RaiseEvent OnClose
47                   lReturn = 0
48                End Select
   Case lngResolveHostMessage
      'A resolve message was posted. This only happens if we connect to a name instead of an IP adress.
      Dim udtHost           As HOSTENT
      Dim lngIpAddrPtr      As Long
      Dim lngHostAddress    As Long
49                If Not lngErrorCode > 0 Then
         'Extract the host name from the memory block
50                   RtlMoveMemory udtHost, ByVal lngMemoryPointer, Len(udtHost)
51                   RtlMoveMemory lngIpAddrPtr, ByVal udtHost.hAddrList, 4
52                   RtlMoveMemory lngHostAddress, ByVal lngIpAddrPtr, 4
         'Free the allocated memory block
53                   Call GlobalUnlock(lngMemoryHandle)
54                   Call GlobalFree(lngMemoryHandle)
55                Else
56                   lngHostAddress = INADDR_NONE
57                End If
      'Now connect to the resolved host.
58                ResolveHostMessage lngHostAddress
59             End Select

60 Exit Sub
ErrorHandler:
61    Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in iSubclass_After on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub



'Purpose:
' When you subclass you can receive messages before or after the original windows handler was called.
Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)
  ' We only use the after.
End Sub




'Purpose:
'Connect to the host that was provided. If it was a name instead of an IP then it will be resolved first.
Private Sub ResolveHost(strHostName As String)
On Error GoTo ErrorHandler
Dim lngAddress As Long

    'Try and resolve the address. This will work if it was an IP we were given
62              lngAddress = inet_addr(strHostName)
    
    'We were unable to resolve it so we will have to go for the long way
63              If lngAddress = INADDR_NONE Then
        'Allocate 1Kb of fixed memory
64                  lngMemoryHandle = GlobalAlloc(GMEM_FIXED, 1024)
        
65                  If lngMemoryHandle > 0 Then
            'Lock the memory block just to get the address
66                      lngMemoryPointer = GlobalLock(lngMemoryHandle)

67                      If lngMemoryPointer = 0 Then
                'Memory allocation error
68                          Call GlobalFree(lngMemoryHandle)
69                          Exit Sub
70                      Else
                'Unlock the memory block
71                          GlobalUnlock (lngMemoryHandle)
72                      End If
73                  Else
            'Memory allocation error
74                      Exit Sub
75                  End If
        
        'Get the host by the name. This is an Asynchroneous call. This means
        'that the call will not freeze the app. It will post a message
        'to the WindowProc when it has finished.
76                  WSAAsyncGetHostByName lngWindowHandle, lngResolveHostMessage, strHostName, ByVal lngMemoryPointer, 1024
77              Else
        'A valid IP adress was suplied as a hostname so we can connect immediately.
78                  ResolveHostMessage (lngAddress)
79              End If

80 Exit Sub
ErrorHandler:
81    Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in ResolveHost on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub



'Purpuse:
' Give a connect message to the winsock.
Private Sub ResolveHostMessage(lngAddress As Long)
On Error GoTo ErrorHandler
Dim udtSocketAddress As SOCKADDR_IN

   ' We will only connect if we have a valid IP adress (converted to a long)
82             If lngAddress <> INADDR_NONE Then
83                With udtSocketAddress
84                   .sin_addr = lngAddress
85                   .sin_port = htons(UnsignedToInteger(CLng(m_intConnectionPort)))
86                   .sin_family = AF_INET
87                End With
      'Call the connect API function
88                Call WSAConnect(m_lngConnectionSocket, udtSocketAddress, LenB(udtSocketAddress))
89             End If
    
90 Exit Sub
ErrorHandler:
91    Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in ResolveHostMesssage on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub

























