VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenericServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Purpose:
' The GenericServer object can be used to create any sort of server application.
' This object will hide the difficulties in the Winsock API calls for your.<BR><BR>
' You will probably write your own protocol handler (like SMTP, POP or ...) in the OnDataArive event.<BR>
Option Explicit

'Some variables that are used to store the state of this object.
Private m_lngSocket As Long               'Just remember the socket where we listen on
Private m_objConnection As New Collection 'When a client is connected to this server then a new connection object will be added to this collection for handeling this connection.
Attribute m_objConnection.VB_VarHelpID = -1
Private m_intConnectionPort As Integer    'Just remember the port where we listen on

' We are going to subclass a generated window and add 2 new message numbers for the communication
Private objSuperClass As Subclass               'This is the subclass handler
Implements WinSubHook.iSubclass                  'This is the interface to the subclass handler
Private WinsockMessage As Long            'Winsock resolve host message
Private WindowHandle As Long              'Window handle

'Purpose: This event will be called when the connection is closed by the server or the client itself.
Public Event OnClose(lngSocket As Long)
'Purpose: This event will be called when the connect method executed successfully.
Public Event OnConnectRequest(lngSocket As Long)
'Purpose: This event will be called when data arived. After this you have to call the .Connection.Recv method to get the data.
Public Event OnDataArrive(lngSocket As Long)
'Purpose: This event will be called whenever there is a problem.
Public Event OnError(lngRetCode As Long, strDescription As String)




'Purpose:
' The moment an instance of this object is created we can set up a new socket for communication.
Private Sub Class_Initialize()
On Error GoTo ErrorHandler
Dim WSAInfo As WSAData
    
   'Create a blank, invisible window
270             WindowHandle = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)

   'Create a unique message number for our messages
271             WinsockMessage = RegisterWindowMessage(App.EXEName & ".WinsockMessage")
    
   'Start up the subclassing
272             Set objSuperClass = New Subclass        'Create a cSuperClass instance
273             With objSuperClass
274               .AddMsg WinsockMessage, MSG_AFTER
275               .Subclass WindowHandle, Me
276             End With
     
   'Start version 1.1 of the Microsoft Winsock API
277             WSAStartup &H101, WSAInfo
    
278 Exit Sub
ErrorHandler:
279    Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in Class_Initialize on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub




'Purpose:
' When this object is closed we just make sure that we close the socket, clean up all the connection and stop the subclassing.
Private Sub Class_Terminate()
On Error GoTo ErrorHandler
    
   'Close all clients that are connected
280             CloseAll
    
   'Cleanup Winsock
281             WSACleanup
        
   'Stop subclassing
282             Set objSuperClass = Nothing  'Destroy the subclasser
    
   'Also remove the window
283             DestroyWindow WindowHandle
       
284 Exit Sub
ErrorHandler:
285    Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in Class_Terminate on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub



'Purpose:
' This method will set the socket in listening mode.
' The moment a client requests a connection the event OnConnectRequest will be trigered.
' You must then call the .Accept method to accept this connection.
Public Function Listen(lngLocalPort As Long) As Long
On Error GoTo ErrorHandler
Dim udtSocketAddress As SOCKADDR_IN
Dim lngReturnValue   As Long
Dim lngAddress       As Long


    'Create a new socket
286              m_lngSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    
    'Let the socket send the messages to the subclassed window handle
287              WSAAsyncSelect m_lngSocket, WindowHandle, WinsockMessage, FD_CONNECT Or FD_READ Or FD_CLOSE Or FD_ACCEPT
288              Listen = m_lngSocket
289              m_intConnectionPort = lngLocalPort

    'If Not m_lngSocket > 0 Then Exit Function 'Check the socket handle
    'If Not lngLocalPort > 0 Then Exit Function 'Check the port number
    
290              With udtSocketAddress
291                  .sin_addr = inet_addr("0.0.0.0")                    'We will accept any IP
292                  .sin_port = htons(UnsignedToInteger(lngLocalPort))  'The port to listen on
293                  .sin_family = AF_INET                               'The internet family (TCP/IP, UDP etc)
294              End With
        
    'Call the bind method
295              Call WSABind(m_lngSocket, udtSocketAddress, LenB(udtSocketAddress))

    'Call the listen function
296              If WSAListen(m_lngSocket, SOMAXCONN) <> 0 Then
297                 RaiseEvent OnError(vbObjectError + 1, "Could not start listening on port " & lngLocalPort)
298              End If

299 Exit Function
ErrorHandler:
300    Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in Listen on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpose:
' You have to call this method to accept a client that wants to connect to this server.
Public Function Accept(ByVal lngSocketHandle As Long) As Long
On Error GoTo ErrorHandler
Dim udtSocketAddress As SOCKADDR_IN

    'Call the accept Winsock API function in order to create a new socket
301              Accept = WSAAccept(lngSocketHandle, udtSocketAddress, LenB(udtSocketAddress))
    
    Dim m_C As New Connection
302              m_C.Socket = Accept
303              m_objConnection.Add m_C

304 Exit Function
ErrorHandler:
305    Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in Accept on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpose:
'Make sure that all clients are disconnected.
Public Function CloseAll()
On Error GoTo ErrorHandler
Dim l As Long
    
    ' Close all connections
306              If m_objConnection.Count > 0 Then
307                 For l = m_objConnection.Count To 1 Step -1
308                    m_objConnection(l).CloseSocket
309                    m_objConnection.Remove l
310                 Next
311              End If
    
    ' Close the listening socket
312              WSACloseSocket m_lngSocket
313              m_lngSocket = 0

314 Exit Function
ErrorHandler:
315    Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in CloseAll on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpose:
' This property can be used to access the connection collection.<BR><BR>
' For instance if you want to send data to a specific client then you can call GenericServer.Connection(lngConnection).Send<BR>
' The variable lngConnection can be the socket number or the negative counter ID<BR>
' For instance if you have 1 connection with socket number 812 then Connection(812) = Connection(-1)<BR>
Public Property Get Connection(lngConnection As Long) As Connection
On Error GoTo ErrorHandler
316             If lngConnection < 0 Then
317                Set Connection = m_objConnection(-lngConnection)
318             Else
      Dim o As Connection
319                For Each o In m_objConnection
320                   If o.Socket = lngConnection Then
321                      Set Connection = o
322                      Exit For
323                   End If
324                Next
325             End If

326 Exit Property
ErrorHandler:
327    Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in Connection on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Property



'Purpose:
' This will give you the socket number where the server is listening on.
Public Property Get ListenSocket() As Long
On Error GoTo ErrorHandler

328            ListenSocket = m_lngSocket

329 Exit Sub
ErrorHandler:
330    Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in ListenSocket on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Property



'Purpose:
' This property will return the number of connected clients.
Public Property Get ConnectionCount() As Long
On Error GoTo ErrorHandler

331            ConnectionCount = m_objConnection.Count

332 Exit Property
ErrorHandler:
333    Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in ConnectionCount on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Property





'Purpuse:
' This is where the winsock messages are handeled.
Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)
On Error GoTo ErrorHandler
'We only use the _After
'wParam is the socket
'lParam is the winsock message
'uMsg is the windows message type
Dim lngErrorCode As Long

   'was an error passed on
334             lngErrorCode = HiWord(lParam)
335             If lngErrorCode > 0 Then
336                RaiseEvent OnError(lngErrorCode, GetErrorDescription(lngErrorCode))
337             End If
   
   'Was it a winsock message?
338             If uMsg = WinsockMessage Then
      'What winsock message
339                Select Case lParam
      Case FD_ACCEPT      'There is a connection request pending
340                   RaiseEvent OnConnectRequest(wParam)
341                   lReturn = 0
      Case FD_READ        'We are ready to read from the winsock buffer


            ' We needed this as a solution to handle faster incomming data than what we can process
342              Static intBufferCounter As Long
343              intBufferCounter = intBufferCounter + 1
344              If intBufferCounter > 1 Then Exit Sub
ContinueWithBuffer:
345              RaiseEvent OnDataArrive(wParam)
346              intBufferCounter = intBufferCounter - 1
347              If intBufferCounter > 0 Then GoTo ContinueWithBuffer

348                   lReturn = 0
      Case FD_CLOSE       'The other end point wishes to close the connection
349                   RaiseEvent OnClose(wParam)
350                   lReturn = 0
351                End Select
352             End If

353 Exit Sub
ErrorHandler:
354    Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in iSubclass_After on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub



'Purpose:
' When you subclass you can receive messages before or after the original windows handler was called.
Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)
'We only use the _After
End Sub


























