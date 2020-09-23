VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Selected Client"
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   11535
      Begin VB.CommandButton cmdSendAll 
         Caption         =   "Send all"
         Height          =   315
         Left            =   7680
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   315
         Left            =   7680
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDataRecv 
         Height          =   1575
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   1440
         Width           =   8535
      End
      Begin VB.TextBox txtSendData 
         Height          =   1080
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   7455
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         Height          =   315
         Left            =   7680
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblRemotePort 
         Caption         =   "Remote Port:"
         Height          =   330
         Left            =   8760
         TabIndex        =   11
         Top             =   2460
         Width           =   2640
      End
      Begin VB.Label lblRemoteIP 
         Caption         =   "Remote IP:"
         Height          =   330
         Left            =   8760
         TabIndex        =   10
         Top             =   2040
         Width           =   2640
      End
      Begin VB.Label lblRemoteHost 
         Caption         =   "Remote Host"
         Height          =   330
         Left            =   8760
         TabIndex        =   9
         Top             =   1620
         Width           =   2640
      End
      Begin VB.Label lblLocalPort 
         Caption         =   "Local Port:"
         Height          =   330
         Left            =   8760
         TabIndex        =   8
         Top             =   1200
         Width           =   2640
      End
      Begin VB.Label lblLocalIP 
         Caption         =   "Local IP:"
         Height          =   330
         Left            =   8760
         TabIndex        =   7
         Top             =   780
         Width           =   2640
      End
      Begin VB.Label lblLocalHost 
         Caption         =   "Local Host:"
         Height          =   330
         Left            =   8760
         TabIndex        =   6
         Top             =   360
         Width           =   2640
      End
   End
   Begin VB.CommandButton cmdDisconnectAll 
      Caption         =   "Disconnect all and quit"
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   2655
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2235
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   3942
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483641
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose:
' This demo will show how to create multiple instances of a generic client and server.
' Created by Edwin Vermeer
' Website http://siteskinner.com
'
'Credits:
' The (super) SubClass code is from Paul Canton [Paul_Caton@hotmail.com]
' see http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=42918&lngWId=1
' Most of the winsock stuff is based on the code from 'Coding Genius'
' see http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=39858&lngWId=1
' Most of the Exception hanler is from Thushan Fernando.
' see http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=41471&lngWId=1Option Explicit

'We will handle a server.
Private WithEvents cServer  As GenericServer    'Server class
Attribute cServer.VB_VarHelpID = -1

Dim lngSocket As Long  'The socket were we listen on



'Purpose:
' Create a new instance of the server object.
Private Sub Form_Load()
On Error GoTo ErrorHandler

32      Set cServer = New GenericServer
    ' Put in the listview headers
33      ListView1.ColumnHeaders.Add 1, , "Socket Handle"
34      ListView1.ColumnHeaders.Add 2, , "Remote Host"
35      ListView1.ColumnHeaders.Add 3, , "Remote IP"
36      ListView1.ColumnHeaders.Add 4, , "Remote Port"
37      ListView1.ColumnHeaders.Add 5, , "Start time"
38      ListView1.ColumnHeaders.Add 6, , "Data in"
39      ListView1.ColumnHeaders.Add 7, , "Data out"
40      ListView1.ColumnHeaders.Add 8, , "Last communication"

41  Exit Sub
ErrorHandler:
42    HandleTheException "frmServer :: Error in Form_Load() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in Form_Load()"
End Sub



'Purpose:
' Make sure that all clients are disconnected and unload the server object.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorHandler
   
   'Unload the client class - This MUST be done
43     cServer.CloseAll
44     Set cServer = Nothing

45  Exit Sub
ErrorHandler:
46    HandleTheException "frmServer :: Error in Form_Unload() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in Form_Unload()"
End Sub



'Purpose:
' Disconnect the active (click on one in the list) client.
Private Sub cmdDisconnect_Click()
On Error GoTo ErrorHandler
   
   ' You have to specify which connection to close
47     If lngSocket = 0 Then
48       MsgBox "First you have to select a connection!", vbCritical, "Close connection"
49     Else
     'Close the socket
50        cServer.Connection(lngSocket).CloseSocket
      'Clear data of active connection
51        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
52        ClearData lngSocket
53     End If

54  Exit Sub
ErrorHandler:
55    HandleTheException "frmServer :: Error in cmdDisconnect_Click() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdDisconnect_Click()"
End Sub



'Purpose:
' Send data to the active (click on one in the list) client.
Private Sub cmdSend_Click()
On Error GoTo ErrorHandler
Dim lngLoop As Long
   
   ' You have to specify which connection you want to use
56     If lngSocket = 0 Then
57       MsgBox "First you have to select a connection!", vbCritical, "Sending data"
58       Exit Sub
59     End If
   
   'Send the data
60     cServer.Connection(lngSocket).Send txtSendData

   ' Go to the coresponding listview item and update it
61     For lngLoop = 1 To ListView1.ListItems.Count
62        If CLng(ListView1.ListItems(lngLoop)) = lngSocket Then
63           ListView1.ListItems(lngLoop).SubItems(6) = ListView1.ListItems(lngLoop).SubItems(6) + Len(txtSendData)
64           ListView1.ListItems(lngLoop).SubItems(7) = Now
65           Exit For
66        End If
67     Next

68  Exit Sub
ErrorHandler:
69    HandleTheException "frmServer :: Error in cmdSend_Click() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdSend_Click()"
End Sub



'Purpose:
' Send data to the all connected clients.
Private Sub cmdSendAll_Click()
On Error GoTo ErrorHandler
Dim lngLoop As Long
   
   'Go through all connections
70     For lngLoop = 1 To ListView1.ListItems.Count
      'Send the data
71        cServer.Connection(CLng(ListView1.ListItems(lngLoop))).Send txtSendData
      'Update the listview
72        ListView1.ListItems(lngLoop).SubItems(6) = ListView1.ListItems(lngLoop).SubItems(6) + Len(txtSendData)
73        ListView1.ListItems(lngLoop).SubItems(7) = Now
74     Next

75  Exit Sub
ErrorHandler:
76    HandleTheException "frmServer :: Error in cmdSendAll_Click() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdSendAll_Click()"
End Sub



'Purpose:
' Just stop.
Private Sub cmdDisconnectAll_Click()
On Error GoTo ErrorHandler
   
77     Unload Me

78  Exit Sub
ErrorHandler:
79    HandleTheException "frmServer :: Error in cmdDisconnectAll_Click() on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdDisconnectAll_Click()"
End Sub



'Purpose:
' Set the server in listening mode on the specified port.
Public Sub Listen(strPort As String)
On Error GoTo ErrorHandler
   
80     cServer.Listen CInt(strPort)
81     Me.Caption = "Server listening at port " & CInt(strPort)

82  Exit Sub
ErrorHandler:
83    HandleTheException "frmServer :: Error in Listen(" & strPort & ") on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in Listen(" & strPort & ")"
End Sub


'Purpose:
' The connection where you click on will be the active connection.
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ErrorHandler
   
84     lngSocket = CLng(Item.Text)
    'Get end point information
85     Frame1.Caption = "Selected Client " & cServer.Connection(lngSocket).GetRemoteHost & " (" & cServer.Connection(lngSocket).GetRemoteIP & ") on port " & cServer.Connection(lngSocket).GetRemotePort
86     lblLocalHost.Caption = "Local Host: " & cServer.Connection(lngSocket).GetLocalHost
87     lblLocalIP.Caption = "Local IP: " & cServer.Connection(lngSocket).GetLocalIP
88     lblLocalPort.Caption = "Local Port: " & cServer.Connection(lngSocket).GetLocalPort
89     lblRemoteHost.Caption = "Remote Host: " & cServer.Connection(lngSocket).GetRemoteHost
90     lblRemoteIP.Caption = "Remote IP: " & cServer.Connection(lngSocket).GetRemoteIP
91     lblRemotePort.Caption = "Remote Port: " & cServer.Connection(lngSocket).GetRemotePort
92     txtDataRecv = ""

93  Exit Sub
ErrorHandler:
94    HandleTheException "frmServer :: Error in ListView1_ItemClick(..) on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in ListView1_ItemClick(..)"
End Sub



'Purpose:
' Whatever was set up as the active connection can not be active anymore.
Private Sub ClearData(lngSocketX As Long)
On Error GoTo ErrorHandler
Dim lngLoop As Long
   
   'Remove it from the list
95     For lngLoop = 1 To ListView1.ListItems.Count
96        If CLng(ListView1.ListItems(lngLoop)) = lngSocketX Then
97           ListView1.ListItems.Remove lngLoop
98           Exit For
99        End If
100     Next
   'Clear the Selected client info
101     If lngSocket = lngSocketX Then
102        lngSocket = 0
103        Frame1.Caption = "Selected Client"
104        lblLocalHost.Caption = "Local Host: "
105        lblLocalIP.Caption = "Local IP: "
106        lblLocalPort.Caption = "Local Port: "
107        lblRemoteHost.Caption = "Remote Host: "
108        lblRemoteIP.Caption = "Remote IP: "
109        lblRemotePort.Caption = "Remote Port: "
110        txtDataRecv = ""
111     End If

112  Exit Sub
ErrorHandler:
113    HandleTheException "frmServer :: Error in ClearData(" & lngSocketX & ") on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in ClearData(" & lngSocketX & ")"
End Sub






'----------------------------------------------------------
' The Server events
'----------------------------------------------------------

'Purpose:
' A client was closed
Private Sub cServer_OnClose(lngSocketX As Long)
On Error GoTo ErrorHandler
   
114     ClearData lngSocketX

115  Exit Sub
ErrorHandler:
116    HandleTheException "frmServer :: Error in cServer_OnClose(" & lngSocketX & ") on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cServer_OnClose(" & lngSocketX & ")"
End Sub



'Purpose:
' A client wants to connect. Accept it.
Private Sub cServer_OnConnectRequest(lngSocket As Long)
On Error GoTo ErrorHandler
Dim lngNewSocket As Long
    
    'Accept the connection and store the new socket handle
117      lngNewSocket = cServer.Accept(lngSocket)
        
    'We use the listbox to hold the info about the new client
    Dim ListHeader    As ListItem
118      Set ListHeader = ListView1.ListItems.Add(, , lngNewSocket)
119      ListHeader.SubItems(1) = cServer.Connection(lngNewSocket).GetRemoteHost
120      ListHeader.SubItems(2) = cServer.Connection(lngNewSocket).GetRemoteIP
121      ListHeader.SubItems(3) = cServer.Connection(lngNewSocket).GetRemotePort
122      ListHeader.SubItems(4) = Now
123      ListHeader.SubItems(5) = 0
124      ListHeader.SubItems(6) = 0
125      ListHeader.SubItems(7) = Now
    
    'Get end point information
126      Me.Caption = "Server " & cServer.Connection(lngNewSocket).GetLocalHost & " (" & cServer.Connection(lngNewSocket).GetLocalIP & ") is listening at port " & cServer.Connection(lngNewSocket).GetLocalPort
    
127  Exit Sub
ErrorHandler:
128    HandleTheException "frmServer :: Error in cServer_OnConnectRequest(" & lngSocket & ") on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cServer_OnConnectRequest(" & lngSocket & ")"
End Sub


'Purpose:
' This event will be triggered when data has arived.
' This is the location where you will write your server side protocol handler.
' In this case we just log the data and update the statistics.
Private Sub cServer_OnDataArrive(lngSocketX As Long)
On Error GoTo ErrorHandler
Dim strData As String
Dim lngLoop As Long
    
   'Recieve data on the server socket
129     cServer.Connection(lngSocketX).Recv strData
    
   ' Go to the coresponding listview item
130     For lngLoop = 1 To ListView1.ListItems.Count
131        If CLng(ListView1.ListItems(lngLoop)) = lngSocketX Then
132           ListView1.ListItems(lngLoop).SubItems(5) = ListView1.ListItems(lngLoop).SubItems(5) + Len(strData)
133           ListView1.ListItems(lngLoop).SubItems(7) = Now
134           Exit For
135        End If
136     Next
    
    ' Only show the data if it's the active/selected client
137      If lngSocket = lngSocketX Then
       'Log it
138         If Len(strData) > 0 Then
139            txtDataRecv.Text = txtDataRecv.Text & strData & vbCrLf
140            txtDataRecv.SelStart = Len(txtDataRecv.Text)
141         End If
142      End If
    
143  Exit Sub
ErrorHandler:
144    HandleTheException "frmServer :: Error in cServer_OnDataArrive(" & lngSocketX & ") on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cServer_OnDataArrive(" & lngSocketX & ")"
End Sub


'Purpose:
' This event is called whenever there was an error.
Private Sub cServer_OnError(lngRetCode As Long, strDescription As String)
On Error GoTo ErrorHandler
    
145      txtDataRecv.Text = txtDataRecv & "*** Error: " & strDescription & vbCrLf
146      txtDataRecv.SelStart = Len(txtDataRecv.Text)

147  Exit Sub
ErrorHandler:
148    HandleTheException "frmServer :: Error in cServer_OnError(..) on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description, "Error in cServer_OnError(..)"
End Sub









