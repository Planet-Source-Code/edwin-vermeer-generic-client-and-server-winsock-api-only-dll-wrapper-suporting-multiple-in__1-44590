VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Purpose:
' This object is created for a client when a connection is made or for a server for every connection that was made.
' You can get connection specific information or receive and send data.
Option Explicit

Dim m_myObject As Object 'used for a custom object reference.
Dim m_lngSocket As Long  'The socket number where we are connected to



'Purpose:
' You can get the socket number of this connection.
Public Property Get Socket() As Long
On Error GoTo ErrorHandler

218            Socket = m_lngSocket

219 Exit Property
ErrorHandler:
220    Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Get Socket on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Property

' It can only be set by other code within this dll.
Friend Property Let Socket(lngSocket As Long)
On Error GoTo ErrorHandler

221    m_lngSocket = lngSocket

222 Exit Property
ErrorHandler:
223    Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Let Socket on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Property


'Purpose:
' You can get or set a custom object. Remember that if you set this you also have to set it to nothing when you end your code. Otherwise this will leave an object reference open whicht will keep an instance of Site Skinner loaded.
' The main purpuse for this is so that you can attach an object of your own to every server connection that is established. This way you can easily access connection specific properties like username, connect duration and last activity.
' You can even add a Site Skinner client object to it so that you can create a proxy, gateway or port mapper with very litle code. Look at the Client Server sample to see how it can be used.
Public Property Get CustomObject() As Object
On Error GoTo ErrorHandler

224    Set CustomObject = m_myObject

225 Exit Property
ErrorHandler:
226    Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Get CustomObject on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Property


Public Property Set CustomObject(yourObject As Object)
On Error GoTo ErrorHandler

227    Set m_myObject = yourObject

228 Exit Property
ErrorHandler:
229    Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Let CustomObject on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Property



'Purpse:
' You can get the name of the client computer.
Public Function GetLocalHost() As String
On Error GoTo ErrorHandler
Dim lngSocket As Long

230     lngSocket = m_lngSocket
231     GetLocalHost = GetIPEndPointField(lngSocket, LOCAL_HOST)

232 Exit Function
ErrorHandler:
233    Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetLocalHost on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpse:
' You can get the IP adress of the client computer.
Public Function GetLocalIP() As String
On Error GoTo ErrorHandler
Dim lngSocket As Long

234     lngSocket = m_lngSocket
235     GetLocalIP = GetIPEndPointField(lngSocket, LOCAL_HOST_IP)

236 Exit Function
ErrorHandler:
237    Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetLocalIP on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpse:
' You can get the connection port on the client computer.
Public Function GetLocalPort() As Long
On Error GoTo ErrorHandler
Dim lngSocket As Long

238     lngSocket = m_lngSocket
239     GetLocalPort = GetIPEndPointField(lngSocket, LOCAL_PORT)

240 Exit Function
ErrorHandler:
241    Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetLocalPort on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function




'Purpse:
' You can get the name of the remote computer.
Public Function GetRemoteHost() As String
On Error GoTo ErrorHandler
Dim lngSocket As Long

242     lngSocket = m_lngSocket
243     GetRemoteHost = GetIPEndPointField(lngSocket, REMOTE_HOST)

244 Exit Function
ErrorHandler:
245    Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetRemoteHost on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpse:
' You can get the IP adress of the remote computer.
Public Function GetRemoteIP() As String
On Error GoTo ErrorHandler
Dim lngSocket As Long

246     lngSocket = m_lngSocket
247     GetRemoteIP = GetIPEndPointField(lngSocket, REMOTE_HOST_IP)

248 Exit Function
ErrorHandler:
249    Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetRemoteIP on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpse:
' You can get the port number of the remote computer.
Public Function GetRemotePort() As Long
On Error GoTo ErrorHandler
Dim lngSocket As Long

250     lngSocket = m_lngSocket
251     GetRemotePort = GetIPEndPointField(lngSocket, REMOTE_PORT)

252 Exit Function
ErrorHandler:
253    Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetRemotePort on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpose:
' Send data to the remote computer.
Public Function Send(strData As String) As Long
On Error GoTo ErrorHandler
Dim lngSocket As Long

254     lngSocket = m_lngSocket
255     mSend lngSocket, strData

256 Exit Function
ErrorHandler:
257    Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Send on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function

'Purpose:
' Send data to the remote computer.
Public Function SendByte(byteData() As Byte) As Long
On Error GoTo ErrorHandler
Dim lngSocket As Long

258     lngSocket = m_lngSocket
259     mSendByte lngSocket, byteData

260 Exit Function
ErrorHandler:
261    Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in SendByte on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpose:
' After receiving an OnDataArive event we can get the data with this.
Public Function Recv(ByRef strBuffer As String) As Long
On Error GoTo ErrorHandler
Dim lngSocket As Long

262     lngSocket = m_lngSocket
263     mRecv lngSocket, strBuffer

264 Exit Function
ErrorHandler:
265    Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Recv on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpose:
' Close the connection.
Public Sub CloseSocket()
On Error GoTo ErrorHandler

266      WSACloseSocket m_lngSocket
267      m_lngSocket = 0

268 Exit Sub
ErrorHandler:
269    Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in CloseSocket on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub



















