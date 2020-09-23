VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Subclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------
'cSubclass - module-less, IDE safe, machine code window subclasser
'
'v1.00 20030107 First cut..........................................................................
'

Option Explicit

Private Const OFFSET_P1   As Long = 9                     'Callback gate address
Private Const OFFSET_P2   As Long = 22                    'Before table entry count
Private Const OFFSET_P3   As Long = 37                    'Before table address
Private Const OFFSET_P4   As Long = 51                    'In IDE ?
Private Const OFFSET_P5   As Long = 69                    'Owner object address for iSubclass_Before
Private Const OFFSET_P6   As Long = 141                   'Original WndProc address
Private Const OFFSET_P7   As Long = 146                   'CallWindowProc address
Private Const OFFSET_P8   As Long = 154                   'After table entry count
Private Const OFFSET_P9   As Long = 169                   'After table address
Private Const OFFSET_PA   As Long = 183                   'In IDE?
Private Const OFFSET_PB   As Long = 201                   'Owner object address for iSubclass_After
Private Const OFFSET_PC   As Long = 250                   'Original WndProc address
Private Const OFFSET_PD   As Long = 260                   'SetWindowLong address
Private Const ARRAY_LB    As Long = 1                     'Low bound of arrays

Private Type tCode
  buf(ARRAY_LB To 272)    As Byte                         'Code buffer
End Type

Private Type tCodeBuf
  Code                    As tCode                        'Subclass WndProc code
End Type

Private CodeBuf           As tCodeBuf                     'Subclass WndProc code instance
Private nBreakGate        As Long                         'Callback breakpoint gate
Private nMsgCntB          As Long                         'Before msg table entry count
Private nMsgCntA          As Long                         'After msg table entry count
Private aMsgTblB()        As WinSubHook.eMsg              'Before msg table array
Private aMsgTblA()        As WinSubHook.eMsg              'After msg table array
Private hWndSubclass      As Long                         'Handle of the window being subclassed
Private nWndProcSubclass  As Long                         'The address of our WndProc
Private nWndProcOriginal  As Long                         'The address of the existing WndProc

'-----------------------------
'Class creation/destruction

'Called automatically when the class instance is created.
Private Sub Class_Initialize()
On Error GoTo ErrorHandler

92  Const OPS As String = "558BEC83C4F85756BE_patch1_33C08945FC8945F8B90000000083F900746183F9FF740CBF000000008B450CF2AF755033C03D_patch4_740B833E007542C70601000000BA_patch5_8B0283F8000F84A50000008D4514508D4510508D450C508D4508508D45FC508D45F8508B0252FF5020C706000000008B45F883F8007570FF7514FF7510FF750CFF750868_patch6_E8_patch7_8945FCB90000000083F900744D83F9FF740CBF000000008B450CF2AF753C33C03D_patchA_740B833E00752EC70601000000BA_patchB_8B0283F8007425FF7514FF7510FF750CFF75088D45FC508B0252FF501CC706000000005E5F8B45FCC9C2100068_patchC_6AFCFF7508E8_patchD_33C08945FCEBE190"
Dim i     As Long, _
    j     As Long, _
    nIDE  As Long

'Convert the string from hexadecimal pairs to bytes and store in the opcode buffer
93    With CodeBuf.Code
94      j = 1                                                 'Set the character index to the start of the opcode string
95      For i = ARRAY_LB To UBound(.buf)                      'For each byte of the code buffer
96        .buf(i) = Val("&H" & Mid$(OPS, j, 2))               'Pull a pair of hex characters and convert to a byte
97        j = j + 2                                           'Bump to the next pair of characters
98      Next i                                                'Next byte of the code buffer
99      nWndProcSubclass = VarPtr(.buf(ARRAY_LB))             'Address of the cSubclass WndProc entry point
100    End With
  
101    nIDE = InIDE                                            'Determine whether we're running in the IDE or not

'Patch the WndProc thunk with runtime values
102    Call PatchVal(OFFSET_P1, VarPtr(nBreakGate))            'Breakpoint gate address
103    Call PatchVal(OFFSET_P4, nIDE)                          'Wether we need check the breakpoint gate and the vtable
104    Call PatchRel(OFFSET_P7, AddrFunc("CallWindowProcA"))   'Address of the CallWindowProc api function
105    Call PatchVal(OFFSET_PA, nIDE)                          'Whether we need check the breakpoint gate and the vtable
106    Call PatchRel(OFFSET_PD, AddrFunc("SetWindowLongA"))    'Address of the SetWindowLong api function

107 Exit Sub
ErrorHandler:
108    Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in Class_Initialize on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub

'Called automatically when the class instance is destroyed.
Private Sub Class_Terminate()
On Error GoTo ErrorHandler
109    If hWndSubclass <> 0 Then                               'If the Subclass thunk is active
110      Call UnSubclass                                       'UnSubclass
111    End If

112 Exit Sub
ErrorHandler:
113    Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in Class_Terminate on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub

'-----------------------------
'Public interface

'Call this method to add a message to the msg callback table. NB This method can be called at any time
Public Sub AddMsg(uMsg As WinSubHook.eMsg, When As WinSubHook.eMsgWhen)
On Error GoTo ErrorHandler
114    If When = WinSubHook.MSG_BEFORE Then                    'If before
115      Call AddMsgSub(uMsg, aMsgTblB, nMsgCntB, When)        'Add the message, pass the before table and before table message count variables
116    Else                                                    'Else after
117      Call AddMsgSub(uMsg, aMsgTblA, nMsgCntA, When)        'Add the message, pass the after table and after table message count variables
118    End If

119 Exit Sub
ErrorHandler:
120    Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in AddMsg on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub

'Allow the user to arbitarily call the original WndProc
Public Function CallOrigWndProc(ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo ErrorHandler

121    If hWndSubclass <> 0 Then
122      CallOrigWndProc = WinSubHook.CallWindowProc( _
    nWndProcOriginal, hWndSubclass, uMsg, wParam, lParam) 'Call the original WndProc
123    Else
124      Debug.Assert False                                    'LOGIC ERROR: The subclasser isn't active!
125    End If

126 Exit Function
ErrorHandler:
127    Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in CallOrigWndProc on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function

'Call this method to delete a message from the msg table. NB This method can be called at any time
Public Sub DelMsg(uMsg As WinSubHook.eMsg, When As WinSubHook.eMsgWhen)
On Error GoTo ErrorHandler
128    If When = WinSubHook.MSG_BEFORE Then                    'If before
129      Call DelMsgSub(uMsg, aMsgTblB, nMsgCntB, When)        'Delete the message, pass the before table and before message count variables
130    Else                                                    'Else after
131      Call DelMsgSub(uMsg, aMsgTblA, nMsgCntA, When)        'Delete the message, pass the after table and after message count variables
132    End If

133 Exit Sub
ErrorHandler:
134    Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in DelMsg on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub

'Call this method to subclass the passed window handle
Public Sub Subclass(hWnd As Long, Owner As WinSubHook.iSubclass)
On Error GoTo ErrorHandler
135    Debug.Assert (hWndSubclass = 0)                         'LOGIC ERROR: The subclasser is already active!
136    Debug.Assert IsWindow(hWnd)                             'LOGIC ERROR: Invalid window handle
  
137    hWndSubclass = hWnd                                     'Store the window handle
138    nWndProcOriginal = WinSubHook.SetWindowLong( _
                                  hWnd, _
                                  WinSubHook.GWL_WNDPROC, _
                                  nWndProcSubclass)       'Set our WndProc in place of the original
139    Debug.Assert nWndProcOriginal                           '??? You can't subclass a window outside of the current process
  
140    Call PatchVal(OFFSET_P5, ObjPtr(Owner))                 'Owner object address for iSubclass_Before
141    Call PatchVal(OFFSET_P6, nWndProcOriginal)              'Original WndProc address for CallWindowProc
142    Call PatchVal(OFFSET_PB, ObjPtr(Owner))                 'Owner object address for iSubclass_After
143    Call PatchVal(OFFSET_PC, nWndProcOriginal)              'Original WndProc address for SetWindowLong

144 Exit Sub
ErrorHandler:
145    Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in SubClass on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub

'Call this method to stop subclassing the window
Public Sub UnSubclass()
On Error GoTo ErrorHandler
146    If hWndSubclass <> 0 Then
147      Call PatchVal(OFFSET_P2, 0)                           'Patch the code to ensure no further iSubclass_Before callbacks
148      Call PatchVal(OFFSET_P8, 0)                           'Patch the code to ensure no further iSubclass_After callbacks
149      Call WinSubHook.SetWindowLong(hWndSubclass, _
                                  WinSubHook.GWL_WNDPROC, _
                                  nWndProcOriginal)       'Restore the original WndProc
150      hWndSubclass = 0                                      'Indicate the subclasser is inactive
151      nMsgCntB = 0                                          'Message before count equals zero
152      nMsgCntA = 0                                          'Message after count equals zero
153    End If

154 Exit Sub
ErrorHandler:
155    Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in UnSubclass on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub

'-----------------------------
' Private subroutines

'Worker sub for AddMsg
Private Sub AddMsgSub(uMsg As WinSubHook.eMsg, aMsgTbl() As WinSubHook.eMsg, nMsgCnt As Long, When As WinSubHook.eMsgWhen)
On Error GoTo ErrorHandler
Dim nEntry  As Long, _
    nOff1   As Long, _
    nOff2   As Long
  
156    If uMsg = WinSubHook.ALL_MESSAGES Then                  'If ALL_MESSAGES
157      nMsgCnt = -1                                          'Indicates that all messages are to callback
158    Else                                                    'Else a specific message number
159      For nEntry = ARRAY_LB To nMsgCnt                      'For each existing entry. NB will skip if nMsgCnt = 0
160        Select Case aMsgTbl(nEntry)                         'Select on the message number stored in this table entry
      Case -1                                             'This msg table slot is a deleted entry
161          aMsgTbl(nEntry) = uMsg                            'Re-use this entry
162          Exit Sub                                          'Bail
      Case uMsg                                           'The msg is already in the table!
163          Exit Sub                                          'Bail
164        End Select
165      Next nEntry                                           'Next entry

'Make space for the new entry
166      ReDim Preserve aMsgTbl(ARRAY_LB To nEntry)            'Increase the size of the table. NB nEntry = nMsgCnt + 1
167      nMsgCnt = nEntry                                      'Bump the entry count
168      aMsgTbl(nEntry) = uMsg                                'Store the message number in the table
169    End If
  
170    If When = WinSubHook.MSG_BEFORE Then                    'If before
171      nOff1 = OFFSET_P2                                     'Patch the Before table entry count
172      nOff2 = OFFSET_P3                                     'Patch the Before table address
173    Else                                                    'Else after
174      nOff1 = OFFSET_P8                                     'Patch the After table entry count
175      nOff2 = OFFSET_P9                                     'Patch the After table address
176    End If

'Patch the appropriate table entries
177    Call PatchVal(nOff1, nMsgCnt)                           'Patch the appropriate table entry count
178    Call PatchVal(nOff2, AddrMsgTbl(aMsgTbl))               'Patch the appropriate table address. We need do this because there's no guarantee that the table existed at SubClass time, the table only gets created if a specific message number is added.

179 Exit Sub
ErrorHandler:
180    Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in AddMsgSub on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub

'Worker sub for DelMsg
Private Sub DelMsgSub(uMsg As WinSubHook.eMsg, aMsgTbl() As WinSubHook.eMsg, nMsgCnt As Long, When As WinSubHook.eMsgWhen)
On Error GoTo ErrorHandler
Dim nEntry As Long
  
181    If uMsg = WinSubHook.ALL_MESSAGES Then                  'If deleting all messages (specific or ALL_MESSAGES)
182      nMsgCnt = 0                                           'Message count is now zero
183      If When = WinSubHook.MSG_BEFORE Then                  'If before
184        nEntry = OFFSET_P2                                  'Patch the before table message count location
185      Else                                                  'Else after
186        nEntry = OFFSET_P8                                  'Patch the after table message count location
187      End If
188      Call PatchVal(nEntry, 0)                              'Patch the table message count
189    Else                                                    'Else deleteting a specific message
190      For nEntry = ARRAY_LB To nMsgCnt                      'For each table entry
191        If aMsgTbl(nEntry) = uMsg Then                      'If this entry is the message we wish to delete
192          aMsgTbl(nEntry) = -1                              'Mark the table slot as available
193          Exit For                                          'Bail
194        End If
195      Next nEntry                                           'Next entry
196    End If

197 Exit Sub
ErrorHandler:
198    Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in DelMsgSub on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub

'Return the address of the passed user32.dll api function
Private Function AddrFunc(sProc As String) As Long
On Error GoTo ErrorHandler

199    AddrFunc = WinSubHook.GetProcAddress(WinSubHook.GetModuleHandle("user32"), sProc)

200 Exit Function
ErrorHandler:
201    Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in AddrFunc on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function

'Return the address of the low bound of the passed table array
Private Function AddrMsgTbl(aMsgTbl() As WinSubHook.eMsg) As Long
On Error Resume Next                                    'The table may not be dimensioned yet so we need protection
202      AddrMsgTbl = VarPtr(aMsgTbl(ARRAY_LB))                'Get the address of the first element of the passed message table
On Error GoTo 0                                         'Switch off error protection

End Function

'Patch the code offset with the passed value
Private Sub PatchVal(nOffset As Long, nValue As Long)
On Error GoTo ErrorHandler

203    Call WinSubHook.CopyMemory(ByVal (nWndProcSubclass + nOffset), nValue, 4)

204 Exit Sub
ErrorHandler:
205    Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in PatchVal on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub

'Patch the code offset with the relative address to the target address
Private Sub PatchRel(nOffset As Long, nTargetAddr As Long)
On Error GoTo ErrorHandler

206    Call WinSubHook.CopyMemory(ByVal (nWndProcSubclass + nOffset), nTargetAddr - nWndProcSubclass - nOffset - 4, 4)


' If you wan to see the error handeling in action just put the line below (unmarked) somewhere (Subclass.cls is a good one)in this code.
' Err.Raise 999, "raised by me", "This is just a test error"

207 Exit Sub
ErrorHandler:
208    Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in PatchRel on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub

'Return -1 if we're running in the IDE or 0 if were running compiled.
Private Function InIDE() As Long
On Error GoTo ErrorHandler
209  Static Value As Long
  
210    If Value = 0 Then
211      Value = 1
212      Debug.Assert True Or InIDE()                          'This line won't exist in the compiled app
213      InIDE = Value - 1
214    End If
  
215    Value = 0

216 Exit Function
ErrorHandler:
217    Err.Raise vbObjectError Or Err, "SubClass", "SubClass :: Error in InIDE on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function






