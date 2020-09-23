Attribute VB_Name = "modWinsock"
'Purpose
' For limiting memory load as much as possible functions are moved to this module.
' If there are multible instances of a class then there will be only one instance of this module loaded.
Option Explicit

'Winsock Initialization and termination
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long

'Server side Winsock API functions
Public Declare Function WSABind Lib "ws2_32.dll" Alias "bind" (ByVal S As Long, ByRef Name As SOCKADDR_IN, ByRef namelen As Long) As Long
Public Declare Function WSAListen Lib "ws2_32.dll" Alias "listen" (ByVal S As Long, ByVal backlog As Long) As Long
Public Declare Function WSAAccept Lib "ws2_32.dll" Alias "accept" (ByVal S As Long, ByRef addr As SOCKADDR_IN, ByRef addrlen As Long) As Long

'String functions
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

'Socket Functions
Public Declare Function WSAConnect Lib "ws2_32.dll" Alias "connect" (ByVal S As Long, ByRef Name As SOCKADDR_IN, ByVal namelen As Long) As Long
Public Declare Function Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Public Declare Function WSACloseSocket Lib "ws2_32.dll" Alias "closesocket" (ByVal S As Long) As Long
Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal S As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long

'Data transfer functions
Public Declare Function WSARecv Lib "ws2_32.dll" Alias "recv" (ByVal S As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function WSASend Lib "ws2_32.dll" Alias "send" (ByVal S As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long

'Network byte ordering functions
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer

'End point information
Public Declare Function getsockname Lib "ws2_32.dll" (ByVal S As Long, ByRef Name As SOCKADDR_IN, ByRef namelen As Long) As Long
Public Declare Function getpeername Lib "ws2_32.dll" (ByVal S As Long, ByRef Name As SOCKADDR_IN, ByRef namelen As Long) As Long

'Hostname resolving functions
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long

'ICMP functions
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Public Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Long

'Winsock API functions for resolving hostnames and IP's
Public Declare Function WSAAsyncGetHostByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal strHostName As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
'Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSADATA As WSAData) As Long
'Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Public Declare Function gethostbyaddr Lib "wsock32.dll" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Public Declare Function gethostname Lib "wsock32.dll" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal szHost As String) As Long

'Memory copy and move functions
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Any, ByVal cbCopy As Long)

'Window creation and destruction functions
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

'Messaging functions
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

'Memory allocation functions
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

'..
Private Const WSADESCRIPTION_LEN = 257
Private Const WSASYS_STATUS_LEN = 129

'Maximum queue length specifiable by listen.
Public Const SOMAXCONN = &H7FFFFFFF

'Windows Socket types
Public Const SOCK_STREAM = 1     'Stream socket

'Address family
Public Const AF_INET = 2          'Internetwork: UDP, TCP, etc.

'Socket Protocol
Public Const IPPROTO_TCP = 6     'tcp

'Data type conversion constants
Public Const OFFSET_4 = 4294967296#
Public Const MAXINT_4 = 2147483647
Public Const OFFSET_2 = 65536
Public Const MAXINT_2 = 32767

'Fixed memory flag for GlobalAlloc
Public Const GMEM_FIXED = &H0

'Winsock error offset
Public Const WSABASEERR = 10000

' Other constants
Public Const ERROR_SUCCESS              As Long = 0
Public Const WS_VERSION_REQD            As Long = &H101
Public Const WS_VERSION_MAJOR           As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR           As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD           As Long = 1
Public Const DATA_SIZE = 32
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128

Public Enum IP_STATUS
    IP_STATUS_BASE = 11000
    IP_SUCCESS = 0
    IP_BUF_TOO_SMALL = (11000 + 1)
    IP_DEST_NET_UNREACHABLE = (11000 + 2)
    IP_DEST_HOST_UNREACHABLE = (11000 + 3)
    IP_DEST_PROT_UNREACHABLE = (11000 + 4)
    IP_DEST_PORT_UNREACHABLE = (11000 + 5)
    IP_NO_RESOURCES = (11000 + 6)
    IP_BAD_OPTION = (11000 + 7)
    IP_HW_ERROR = (11000 + 8)
    IP_PACKET_TOO_BIG = (11000 + 9)
    IP_REQ_TIMED_OUT = (11000 + 10)
    IP_BAD_REQ = (11000 + 11)
    IP_BAD_ROUTE = (11000 + 12)
    IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
    IP_TTL_EXPIRED_REASSEM = (11000 + 14)
    IP_PARAM_PROBLEM = (11000 + 15)
    IP_SOURCE_QUENCH = (11000 + 16)
    IP_OPTION_TOO_BIG = (11000 + 17)
    IP_BAD_DESTINATION = (11000 + 18)
    IP_ADDR_DELETED = (11000 + 19)
    IP_SPEC_MTU_CHANGE = (11000 + 20)
    IP_MTU_CHANGE = (11000 + 21)
    IP_UNLOAD = (11000 + 22)
    IP_ADDR_ADDED = (11000 + 23)
    IP_GENERAL_FAILURE = (11000 + 50)
    MAX_IP_STATUS = 11000 + 50
    IP_PENDING = (11000 + 255)
    PING_TIMEOUT = 255
End Enum


'Winsock messages that will go to the window handler
Public Enum WSAMessage
    FD_READ = &H1&      'Data is ready to be read from the buffer
    FD_CONNECT = &H10&  'Connection esatblished
    FD_CLOSE = &H20&    'Connection closed
    FD_ACCEPT = &H8&    'Connection request pending
End Enum

Public Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Public Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Long
  ' Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type


'Winsock Data structure
Public Type WSAData
    wVersion       As Integer                       'Version
    wHighVersion   As Integer                       'High Version
    szDescription  As String * WSADESCRIPTION_LEN   'Description
    szSystemStatus As String * WSASYS_STATUS_LEN    'Status of system
    iMaxSockets    As Integer                       'Maximum number of sockets allowed
    iMaxUdpDg      As Integer                       'Maximum UDP datagrams
    lpVendorInfo   As Long                          'Vendor Info
End Type

'HostEnt Structure
Public Type HOSTENT
    hName     As Long       'Host Name
    hAliases  As Long       'Alias
    hAddrType As Integer    'Address Type
    hLength   As Integer    'Length
    hAddrList As Long       'Address List
End Type

'Socket Address structure
Public Type SOCKADDR_IN
    sin_family       As Integer 'Address familly
    sin_port         As Integer 'Port
    sin_addr         As Long    'Long address
    sin_zero(1 To 8) As Byte
End Type

'End Point of connection information
Public Enum IPEndPointFields
    LOCAL_HOST          'Local hostname
    LOCAL_HOST_IP       'Local IP
    LOCAL_PORT          'Local port
    REMOTE_HOST         'Remote hostname
    REMOTE_HOST_IP      'Remote IP
    REMOTE_PORT         'Remote port
End Enum

'Basic Winsock error results.
Public Enum WSABaseErrors
    INADDR_NONE = &HFFFF
    SOCKET_ERROR = -1
    INVALID_SOCKET = -1
End Enum

'Winsock error constants
Public Enum WSAErrorConstants
'Windows Sockets definitions of regular Microsoft C error constants
    WSAEINTR = (WSABASEERR + 4)
    WSAEBADF = (WSABASEERR + 9)
    WSAEACCES = (WSABASEERR + 13)
    WSAEFAULT = (WSABASEERR + 14)
    WSAEINVAL = (WSABASEERR + 22)
    WSAEMFILE = (WSABASEERR + 24)
'Windows Sockets definitions of regular Berkeley error constants
    WSAEWOULDBLOCK = (WSABASEERR + 35)
    WSAEINPROGRESS = (WSABASEERR + 36)
    WSAEALREADY = (WSABASEERR + 37)
    WSAENOTSOCK = (WSABASEERR + 38)
    WSAEDESTADDRREQ = (WSABASEERR + 39)
    WSAEMSGSIZE = (WSABASEERR + 40)
    WSAEPROTOTYPE = (WSABASEERR + 41)
    WSAENOPROTOOPT = (WSABASEERR + 42)
    WSAEPROTONOSUPPORT = (WSABASEERR + 43)
    WSAESOCKTNOSUPPORT = (WSABASEERR + 44)
    WSAEOPNOTSUPP = (WSABASEERR + 45)
    WSAEPFNOSUPPORT = (WSABASEERR + 46)
    WSAEAFNOSUPPORT = (WSABASEERR + 47)
    WSAEADDRINUSE = (WSABASEERR + 48)
    WSAEADDRNOTAVAIL = (WSABASEERR + 49)
    WSAENETDOWN = (WSABASEERR + 50)
    WSAENETUNREACH = (WSABASEERR + 51)
    WSAENETRESET = (WSABASEERR + 52)
    WSAECONNABORTED = (WSABASEERR + 53)
    WSAECONNRESET = (WSABASEERR + 54)
    WSAENOBUFS = (WSABASEERR + 55)
    WSAEISCONN = (WSABASEERR + 56)
    WSAENOTCONN = (WSABASEERR + 57)
    WSAESHUTDOWN = (WSABASEERR + 58)
    WSAETOOMANYREFS = (WSABASEERR + 59)
    WSAETIMEDOUT = (WSABASEERR + 60)
    WSAECONNREFUSED = (WSABASEERR + 61)
    WSAELOOP = (WSABASEERR + 62)
    WSAENAMETOOLONG = (WSABASEERR + 63)
    WSAEHOSTDOWN = (WSABASEERR + 64)
    WSAEHOSTUNREACH = (WSABASEERR + 65)
    WSAENOTEMPTY = (WSABASEERR + 66)
    WSAEPROCLIM = (WSABASEERR + 67)
    WSAEUSERS = (WSABASEERR + 68)
    WSAEDQUOT = (WSABASEERR + 69)
    WSAESTALE = (WSABASEERR + 70)
    WSAEREMOTE = (WSABASEERR + 71)
'Extended Windows Sockets error constant definitions
    WSASYSNOTREADY = (WSABASEERR + 91)
    WSAVERNOTSUPPORTED = (WSABASEERR + 92)
    WSANOTINITIALISED = (WSABASEERR + 93)
    WSAEDISCON = (WSABASEERR + 101)
    WSAENOMORE = (WSABASEERR + 102)
    WSAECANCELLED = (WSABASEERR + 103)
    WSAEINVALIDPROCTABLE = (WSABASEERR + 104)
    WSAEINVALIDPROVIDER = (WSABASEERR + 105)
    WSAEPROVIDERFAILEDINIT = (WSABASEERR + 106)
    WSASYSCALLFAILURE = (WSABASEERR + 107)
    WSASERVICE_NOT_FOUND = (WSABASEERR + 108)
    WSATYPE_NOT_FOUND = (WSABASEERR + 109)
    WSA_E_NO_MORE = (WSABASEERR + 110)
    WSA_E_CANCELLED = (WSABASEERR + 111)
    WSAEREFUSED = (WSABASEERR + 112)
    WSAHOST_NOT_FOUND = 11001
    WSATRY_AGAIN = 11002
    WSANO_RECOVERY = 11003
    WSANO_DATA = 11004
    FD_SETSIZE = 64
End Enum

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'Purpose:
' Convert an unsigned long to an integer.
Public Function UnsignedToInteger(Value As Long) As Integer
On Error GoTo ErrorHandler

355              If Value < 0 Or Value >= OFFSET_2 Then Error 6  'Overflow
    
356              If Value <= MAXINT_2 Then
357                  UnsignedToInteger = Value
358              Else
359                  UnsignedToInteger = Value - OFFSET_2
360              End If

361 Exit Function
ErrorHandler:
362    Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in UnsignedToInteger on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpose:
' Convert an integer to an unsigned long.
Public Function IntegerToUnsigned(Value As Integer) As Long
On Error GoTo ErrorHandler


363              If Value < 0 Then
364                  IntegerToUnsigned = Value + OFFSET_2
365              Else
366                  IntegerToUnsigned = Value
367              End If
    
368 Exit Function
ErrorHandler:
369    Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in IntegerToUnsigned on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpose:
' Create a string from a pointer
Public Function StringFromPointer(ByVal lngPointer As Long) As String
On Error GoTo ErrorHandler


  Dim strTemp As String
  Dim lRetVal As Long
    
370              strTemp = String$(lstrlen(ByVal lngPointer), 0)    'prepare the strTemp buffer
371              lRetVal = lstrcpy(ByVal strTemp, ByVal lngPointer) 'copy the string into the strTemp buffer
372              If lRetVal Then StringFromPointer = strTemp        'return the string

373 Exit Function
ErrorHandler:
374    Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in StringFromPointer on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpose:
' Return the Hi Word of a long value.
Public Function HiWord(lngValue As Long) As Long
On Error GoTo ErrorHandler

375              If (lngValue And &H80000000) = &H80000000 Then
376                  HiWord = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000&
377              Else
378                  HiWord = (lngValue And &HFFFF0000) \ &H10000
379              End If
    
380 Exit Function
ErrorHandler:
381    Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in HiWord on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpose:
' Get the received data from the socket and return it in the calling string.
Public Function mRecv(ByVal lngSocket As Long, ByRef strBuffer As String) As Long
On Error GoTo ErrorHandler

382            Const MAX_BUFFER_LENGTH As Long = 8192 'Normal= 8192  'MAX = 65536

  Dim arrBuffer(1 To MAX_BUFFER_LENGTH)   As Byte
  Dim lngBytesReceived                    As Long
  Dim strTempBuffer                       As String
    
    'Call the recv Winsock API function in order to read data from the buffer
383              lngBytesReceived = WSARecv(lngSocket, arrBuffer(1), MAX_BUFFER_LENGTH, 0&)

384              If lngBytesReceived > 0 Then
        'If we have received some data, convert it to the Unicode
        'string that is suitable for the Visual Basic String data type
385                  strTempBuffer = StrConv(arrBuffer, vbUnicode)

        'Remove unused bytes
386                  strBuffer = Left$(strTempBuffer, lngBytesReceived)
387              End If
        
388              mRecv = lngBytesReceived

389 Exit Function
ErrorHandler:
390    Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in mRecv on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpose:
' Send data to the specified socket.
Public Function mSend(ByVal lngSocket As Long, strData As String) As Long
On Error GoTo ErrorHandler

  Dim arrBuffer()     As Byte

    'Convert the data string to a byte array
391              arrBuffer() = StrConv(strData, vbFromUnicode)
    'Call the send Winsock API function in order to send data
392              mSend = WSASend(lngSocket, arrBuffer(0), Len(strData), 0&)

393 Exit Function
ErrorHandler:
394    Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in mSend on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function


'Purpose:
' Send data to the specified socket.
Public Function mSendByte(ByVal lngSocket As Long, byteData() As Byte) As Long
On Error GoTo ErrorHandler

    'Call the send Winsock API function in order to send data
395          mSendByte = WSASend(lngSocket, byteData(0), UBound(byteData), 0&)

396 Exit Function
ErrorHandler:
397    Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in mSendByte on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function

'Purpose:
' Get the IP adress of an endpoint (client or server).
Public Function GetIPEndPointField(ByVal lngSocket As Long, ByVal EndpointField As IPEndPointFields) As Variant
On Error GoTo ErrorHandler

  Dim udtSocketAddress    As SOCKADDR_IN
  Dim lngReturnValue      As Long
  Dim lngPtrToAddress     As Long
  Dim strIPAddress        As String
  Dim lngAddress          As Long

398              Select Case EndpointField
        Case LOCAL_HOST, LOCAL_HOST_IP, LOCAL_PORT

            'If the info of a local end-point of the connection is
            'requested, call the getsockname Winsock API function
399                      lngReturnValue = getsockname(lngSocket, udtSocketAddress, LenB(udtSocketAddress))
        Case REMOTE_HOST, REMOTE_HOST_IP, REMOTE_PORT
            
            'If the info of a remote end-point of the connection is
            'requested, call the getpeername Winsock API function
400                      lngReturnValue = getpeername(lngSocket, udtSocketAddress, LenB(udtSocketAddress))
401              End Select
    
    
402              If lngReturnValue = 0 Then
        'If no errors occurred, the getsockname or getpeername function returns 0.

403                  Select Case EndpointField
            Case LOCAL_PORT, REMOTE_PORT
                'Get the port number from the sin_port field and convert the byte ordering
404                          GetIPEndPointField = IntegerToUnsigned(ntohs(udtSocketAddress.sin_port))
            
            Case LOCAL_HOST_IP, REMOTE_HOST_IP
  
                'Get pointer to the string that contains the IP address
405                          lngPtrToAddress = inet_ntoa(udtSocketAddress.sin_addr)
                
                'Retrieve that string by the pointer
406                          GetIPEndPointField = StringFromPointer(lngPtrToAddress)
            Case LOCAL_HOST, REMOTE_HOST

                'The same procedure as for an IP address only using GetHostNameByAddress
407                          lngPtrToAddress = inet_ntoa(udtSocketAddress.sin_addr)
408                          strIPAddress = StringFromPointer(lngPtrToAddress)
409                          lngAddress = inet_addr(strIPAddress)
410                          GetIPEndPointField = GetHostNameByAddress(lngAddress)

411                  End Select
    'An error occured
412              Else
413                  GetIPEndPointField = SOCKET_ERROR
414              End If
    
415 Exit Function
ErrorHandler:
416    Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in GetIPEndPointField on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpose:
' Get the hostname of an endpoint (client or server).
Private Function GetHostNameByAddress(lngInetAdr As Long) As String
On Error GoTo ErrorHandler

  Dim lngPtrHostEnt As Long
  Dim udtHostEnt    As HOSTENT
  Dim strHostName   As String
  
    'Get the pointer to the HOSTENT structure
417              lngPtrHostEnt = gethostbyaddr(lngInetAdr, 4, AF_INET)
    
    'Copy data into the HOSTENT structure
418              RtlMoveMemory udtHostEnt, ByVal lngPtrHostEnt, LenB(udtHostEnt)
    
    'Prepare the buffer to receive a string
419              strHostName = String$(256, 0)
    
    'Copy the host name into the strHostName variable
420              RtlMoveMemory ByVal strHostName, ByVal udtHostEnt.hName, 256
    
    'Cut received string by first chr(0) character
421              GetHostNameByAddress = Left$(strHostName, InStr(1, strHostName, Chr(0)) - 1)

422 Exit Function
ErrorHandler:
423    Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in GetHostNameByAddress on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpose:
' Get the error description of a socket error.
Public Function GetErrorDescription(ByVal lngErrorCode As Long) As String
On Error GoTo ErrorHandler

  Dim strDesc As String
    
424              Select Case lngErrorCode
        Case WSAEACCES
425                      strDesc = "Permission denied."
        Case WSAEADDRINUSE
426                      strDesc = "Address already in use."
        Case WSAEADDRNOTAVAIL
427                      strDesc = "Cannot assign requested address."
        Case WSAEAFNOSUPPORT
428                      strDesc = "Address family not supported by protocol family."
        Case WSAEALREADY
429                      strDesc = "Operation already in progress."
        Case WSAECONNABORTED
430                      strDesc = "Software caused connection abort."
        Case WSAECONNREFUSED
431                      strDesc = "Connection refused."
        Case WSAECONNRESET
432                      strDesc = "Connection reset by peer."
        Case WSAEDESTADDRREQ
433                      strDesc = "Destination address required."
        Case WSAEFAULT
434                      strDesc = "Bad address."
        Case WSAEHOSTDOWN
435                      strDesc = "Host is down."
        Case WSAEHOSTUNREACH
436                      strDesc = "No route to host."
        Case WSAEINPROGRESS
437                      strDesc = "Operation now in progress."
        Case WSAEINTR
438                      strDesc = "Interrupted function call."
        Case WSAEINVAL
439                      strDesc = "Invalid argument."
        Case WSAEISCONN
440                      strDesc = "Socket is already connected."
        Case WSAEMFILE
441                      strDesc = "Too many open files."
        Case WSAEMSGSIZE
442                      strDesc = "Message too long."
        Case WSAENETDOWN
443                      strDesc = "Network is down."
        Case WSAENETRESET
444                      strDesc = "Network dropped connection on reset."
        Case WSAENETUNREACH
445                      strDesc = "Network is unreachable."
        Case WSAENOBUFS
446                      strDesc = "No buffer space available."
        Case WSAENOPROTOOPT
447                      strDesc = "Bad protocol option."
        Case WSAENOTCONN
448                      strDesc = "Socket is not connected."
        Case WSAENOTSOCK
449                      strDesc = "Socket operation on nonsocket."
        Case WSAEOPNOTSUPP
450                      strDesc = "Operation not supported."
        Case WSAEPFNOSUPPORT
451                      strDesc = "Protocol family not supported."
        Case WSAEPROCLIM
452                      strDesc = "Too many processes."
        Case WSAEPROTONOSUPPORT
453                      strDesc = "Protocol not supported."
        Case WSAEPROTOTYPE
454                      strDesc = "Protocol wrong type for socket."
        Case WSAESHUTDOWN
455                      strDesc = "Cannot send after socket shutdown."
        Case WSAESOCKTNOSUPPORT
456                      strDesc = "Socket type not supported."
        Case WSAETIMEDOUT
457                      strDesc = "Connection timed out."
        Case WSATYPE_NOT_FOUND
458                      strDesc = "Class type not found."
        Case WSAEWOULDBLOCK
459                      strDesc = "Resource temporarily unavailable."
        Case WSAHOST_NOT_FOUND
460                      strDesc = "Host not found."
        Case WSANOTINITIALISED
461                      strDesc = "Successful WSAStartup not yet performed."
        Case WSANO_DATA
462                      strDesc = "Valid name, no data record of requested type."
        Case WSANO_RECOVERY
463                      strDesc = "This is a nonrecoverable error."
        Case WSASYSCALLFAILURE
464                      strDesc = "System call failure."
        Case WSASYSNOTREADY
465                      strDesc = "Network subsystem is unavailable."
        Case WSATRY_AGAIN
466                      strDesc = "Nonauthoritative host not found."
        Case WSAVERNOTSUPPORTED
467                      strDesc = "Winsock.dll version out of range."
        Case WSAEDISCON
468                      strDesc = "Graceful shutdown in progress."
        Case Else
469                      strDesc = "Unknown error."
470              End Select
    
471              GetErrorDescription = strDesc
    
472 Exit Function
ErrorHandler:
473    Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in GetErrorDescription on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function


Public Sub SocketsCleanup()
On Error GoTo ErrorHandler
    
474             WSACleanup

475 Exit Sub
ErrorHandler:
476    Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in SocketsCleanup on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub


Public Function SocketsInitialize() As Boolean
On Error GoTo ErrorHandler
Dim WSAD            As WSAData

477             SocketsInitialize = False

478             If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then Exit Function
479             If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then Exit Function
  
480             SocketsInitialize = True

481 Exit Function
ErrorHandler:
482    Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in SocketsInitialize on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



























