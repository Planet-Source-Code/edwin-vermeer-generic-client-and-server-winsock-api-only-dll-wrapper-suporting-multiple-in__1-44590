Attribute VB_Name = "modExceptionHandler"

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As EXCEPTION_RECORD, ByVal LPEXCEPTION_RECORD As Long, ByVal lngBytes As Long)
Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
Private Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)

Public Enum enumExceptionType
    enumExceptionType_AccessViolation = &HC0000005
    enumExceptionType_DataTypeMisalignment = &H80000002
    enumExceptionType_Breakpoint = &H80000003
    enumExceptionType_SingleStep = &H80000004
    enumExceptionType_ArrayBoundsExceeded = &HC000008C
    enumExceptionType_FaultDenormalOperand = &HC000008D
    enumExceptionType_FaultDivideByZero = &HC000008E
    enumExceptionType_FaultInexactResult = &HC000008F
    enumExceptionType_FaultInvalidOperation = &HC0000090
    enumExceptionType_FaultOverflow = &HC0000091
    enumExceptionType_FaultStackCheck = &HC0000092
    enumExceptionType_FaultUnderflow = &HC0000093
    enumExceptionType_IntegerDivisionByZero = &HC0000094
    enumExceptionType_IntegerOverflow = &HC0000095
    enumExceptionType_PriviledgedInstruction = &HC0000096
    enumExceptionType_InPageError = &HC0000006
    enumExceptionType_IllegalInstruction = &HC000001D
    enumExceptionType_NoncontinuableException = &HC0000025
    enumExceptionType_StackOverflow = &HC00000FD
    enumExceptionType_InvalidDisposition = &HC0000026
    enumExceptionType_GuardPageViolation = &H80000001
    enumExceptionType_InvalidHandle = &HC0000008
    enumExceptionType_ControlCExit = &HC000013A
End Enum


Private Const EXCEPTION_MAXIMUM_PARAMETERS = 15

Private Type CONTEXT
    FltF0        As Double
    FltF1        As Double
    FltF2        As Double
    FltF3        As Double
    FltF4        As Double
    FltF5        As Double
    FltF6        As Double
    FltF7        As Double
    FltF8        As Double
    FltF9        As Double
    FltF10       As Double
    FltF11       As Double
    FltF12       As Double
    FltF13       As Double
    FltF14       As Double
    FltF15       As Double
    FltF16       As Double
    FltF17       As Double
    FltF18       As Double
    FltF19       As Double
    FltF20       As Double
    FltF21       As Double
    FltF22       As Double
    FltF23       As Double
    FltF24       As Double
    FltF25       As Double
    FltF26       As Double
    FltF27       As Double
    FltF28       As Double
    FltF29       As Double
    FltF30       As Double
    FltF31       As Double
    IntV0        As Double
    IntT0        As Double
    IntT1        As Double
    IntT2        As Double
    IntT3        As Double
    IntT4        As Double
    IntT5        As Double
    IntT6        As Double
    IntT7        As Double
    IntS0        As Double
    IntS1        As Double
    IntS2        As Double
    IntS3        As Double
    IntS4        As Double
    IntS5        As Double
    IntFp        As Double
    IntA0        As Double
    IntA1        As Double
    IntA2        As Double
    IntA3        As Double
    IntA4        As Double
    IntA5        As Double
    IntT8        As Double
    IntT9        As Double
    IntT10       As Double
    IntT11       As Double
    IntRa        As Double
    IntT12       As Double
    IntAt        As Double
    IntGp        As Double
    IntSp        As Double
    IntZero      As Double
    Fpcr         As Double
    SoftFpcr     As Double
    Fir          As Double
    Psr          As Long
    ContextFlags As Long
    Fill(4)      As Long
End Type
Private Type EXCEPTION_RECORD
    ExceptionCode                                        As Long
    ExceptionFlags                                       As Long
    pExceptionRecord                                     As Long
    ExceptionAddress                                     As Long
    NumberParameters                                     As Long
    ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS)   As Long
End Type

Private Type EXCEPTION_POINTERS
    pExceptionRecord     As EXCEPTION_RECORD
    ContextRecord        As CONTEXT
End Type

Public blnIsHandlerInstalled As Boolean



Public Sub HandleTheException(strException As String, strProcedure As String)
Dim strEx As String
190          With frmException
191              If InStr(1, strException, "Script error") = 0 Then
192                 .lblWarning = "Vermeer Automatisering Apologizes for the inconvience but an Exception Error occured in the application. It may be possible for you to continue to run this application without issues but we recommend you only do so if your certain that its OK to do so. Click 'Continue' to resume ignoring the error or 'Exit' to terminate the application immediately."
193                 .lblToDo = "If the error continues, please contact Vermeer Automatisering at support@siteskinner.com with a detailed description of your system and how to reproduce the issue."
194                 .cmdExit.Visible = True
195              Else
196                 .lblWarning = "Vermeer Automatisering Apologizes for the inconvience but an Exception Error occured in the script that you are trying to execute."
197                 .lblToDo = "Please update the script and try running it again. If the error continues, then please contact the developer of the script with a detailed description how to reproduce the issue."
198                 .cmdExit.Visible = False
199              End If
200              strEx = Replace(strException, " - Error", vbCrLf & "Error")
                 .lblErrorTitle = strProcedure & " occured on " & Date & " " & Time
201              .txtException.Text = "####  Error in " & App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & " occured on " & Date & " " & Time & "  ####" & vbCrLf & strEx
202              .Show vbModal
203              If Not .bContinue Then
204                  Unload frmDemo
205                  DoEvents
206                  End '// End is not wize to use, but after unloading frmMain there could still be objects in memory
207              End If
208          End With
End Sub

Public Function ExceptionHandler(ByRef ExceptionPtrs As EXCEPTION_POINTERS) As Long
    On Error Resume Next
    
    Dim ExceptionRecord As EXCEPTION_RECORD
    Dim strExceptionDescriptiosn As String
  
209              ExceptionRecord = ExceptionPtrs.pExceptionRecord
  
210              Do Until ExceptionRecord.pExceptionRecord = 0
211                  CopyMemory ExceptionRecord, ExceptionRecord.pExceptionRecord, Len(ExceptionRecord)
212              Loop
  
213              strExceptionDescriptiosn = GetExceptionDescription(ExceptionRecord.ExceptionCode)
  
    On Error GoTo 0
214              Err.Raise vbObjectError, "ExceptionHandler", "Exception: " & strExceptionDescriptiosn & " [" & GetExceptionName(ExceptionRecord.ExceptionCode) & "]" & vbCrLf & "ExceptionAddress : " & ExceptionRecord.ExceptionAddress
End Function


Public Sub InstallExceptionHandler()
    On Error Resume Next
215              If Not blnIsHandlerInstalled Then
216                  Call SetUnhandledExceptionFilter(AddressOf ExceptionHandler)
217                  blnIsHandlerInstalled = True
218              End If
End Sub



Public Function GetExceptionDescription(ExceptionType As enumExceptionType) As String
    On Error Resume Next
    
    Dim strDescription As String
  
219              Select Case ExceptionType
        
        Case enumExceptionType_AccessViolation
220                      strDescription = "Access Violation"
        
        Case enumExceptionType_DataTypeMisalignment
221                      strDescription = "Data Type Misalignment"
        
        Case enumExceptionType_Breakpoint
222                      strDescription = "Breakpoint"
        
        Case enumExceptionType_SingleStep
223                      strDescription = "Single Step"
        
        Case enumExceptionType_ArrayBoundsExceeded
224                      strDescription = "Array Bounds Exceeded"
        
        Case enumExceptionType_FaultDenormalOperand
225                      strDescription = "Float Denormal Operand"
        
        Case enumExceptionType_FaultDivideByZero
226                      strDescription = "Divide By Zero"
        
        Case enumExceptionType_FaultInexactResult
227                      strDescription = "Floating Point Inexact Result"
        
        Case enumExceptionType_FaultInvalidOperation
228                      strDescription = "Invalid Operation"
        
        Case enumExceptionType_FaultOverflow
229                      strDescription = "Float Overflow"
        
        Case enumExceptionType_FaultStackCheck
230                      strDescription = "Float Stack Check"
        
        Case enumExceptionType_FaultUnderflow
231                      strDescription = "Float Underflow"
        
        Case enumExceptionType_IntegerDivisionByZero
232                      strDescription = "Integer Divide By Zero"
        
        Case enumExceptionType_IntegerOverflow
233                      strDescription = "Integer Overflow"
        
        Case enumExceptionType_PriviledgedInstruction
234                      strDescription = "Privileged Instruction"
        
        Case enumExceptionType_InPageError
235                      strDescription = "In Page Error"
        
        Case enumExceptionType_IllegalInstruction
236                      strDescription = "Illegal Instruction"
        
        Case enumExceptionType_NoncontinuableException
237                      strDescription = "Non Continuable Exception"
        
        Case enumExceptionType_StackOverflow
238                      strDescription = "Stack Overflow"
        
        Case enumExceptionType_InvalidDisposition
239                      strDescription = "Invalid Disposition"
        
        Case enumExceptionType_GuardPageViolation
240                      strDescription = "Guard Page Violation"
        
        Case enumExceptionType_InvalidHandle
241                      strDescription = "Invalid Handle"
        
        Case enumExceptionType_ControlCExit
242                      strDescription = "Control-C Exit"
        
        Case Else
243                      strDescription = "Unknown Exception Error"
    
244              End Select
    
245              GetExceptionDescription = strDescription
End Function


Public Function GetExceptionName(ExceptionType As enumExceptionType) As String
    On Error Resume Next
    
    Dim strDescription As String
  
246              Select Case ExceptionType
        
        Case enumExceptionType_AccessViolation
247                      strDescription = "EXCEPTION_ACCESS_VIOLATION"
        
        Case enumExceptionType_DataTypeMisalignment
248                      strDescription = "EXCEPTION_DATATYPE_MISALIGNMENT"
        
        Case enumExceptionType_Breakpoint
249                      strDescription = "EXCEPTION_BREAKPOINT"
        
        Case enumExceptionType_SingleStep
250                      strDescription = "EXCEPTION_SINGLE_STEP"
        
        Case enumExceptionType_ArrayBoundsExceeded
251                      strDescription = "EXCEPTION_ARRAY_BOUNDS_EXCEEDED"
        
        Case enumExceptionType_FaultDenormalOperand
252                      strDescription = "EXCEPTION_FLT_DENORMAL_OPERAND"
        
        Case enumExceptionType_FaultDivideByZero
253                      strDescription = "EXCEPTION_FLT_DIVIDE_BY_ZERO"
        
        Case enumExceptionType_FaultInexactResult
254                      strDescription = "EXCEPTION_FLT_INEXACT_RESULT"
        
        Case enumExceptionType_FaultInvalidOperation
255                      strDescription = "EXCEPTION_FLT_INVALID_OPERATION"
        
        Case enumExceptionType_FaultOverflow
256                      strDescription = "EXCEPTION_FLT_OVERFLOW"
        
        Case enumExceptionType_FaultStackCheck
257                      strDescription = "EXCEPTION_FLT_STACK_CHECK"
        
        Case enumExceptionType_FaultUnderflow
258                      strDescription = "EXCEPTION_FLT_UNDERFLOW"
        
        Case enumExceptionType_IntegerDivisionByZero
259                      strDescription = "EXCEPTION_INT_DIVIDE_BY_ZERO"
        
        Case enumExceptionType_IntegerOverflow
260                      strDescription = "EXCEPTION_INT_OVERFLOW"
        
        Case enumExceptionType_PriviledgedInstruction
261                      strDescription = "EXCEPTION_PRIVILEGED_INSTRUCTION"
        
        Case enumExceptionType_InPageError
262                      strDescription = "EXCEPTION_IN_PAGE_ERROR"
        
        Case enumExceptionType_IllegalInstruction
263                      strDescription = "EXCEPTION_ILLEGAL_INSTRUCTION"
        
        Case enumExceptionType_NoncontinuableException
264                      strDescription = "EXCEPTION_NONCONTINUABLE_EXCEPTION"
        
        Case enumExceptionType_StackOverflow
265                      strDescription = "EXCEPTION_STACK_OVERFLOW"
        
        Case enumExceptionType_InvalidDisposition
266                      strDescription = "EXCEPTION_INVALID_DISPOSITION"
        
        Case enumExceptionType_GuardPageViolation
267                      strDescription = "EXCEPTION_GUARD_PAGE_VIOLATION"
        
        Case enumExceptionType_InvalidHandle
268                      strDescription = "EXCEPTION_INVALID_HANDLE"
        
        Case enumExceptionType_ControlCExit
269                      strDescription = "EXCEPTION_CONTROL_C_EXIT"
        
        Case Else
270                      strDescription = "Unknown"
    
271              End Select
    
272              GetExceptionName = strDescription
End Function


Public Sub RaiseAnException(ExceptionType As enumExceptionType)
273              RaiseException ExceptionType, 0, 0, 0
End Sub

Public Sub UninstallExceptionHandler()
    On Error Resume Next
    
274              If blnIsHandlerInstalled Then
275                  Call SetUnhandledExceptionFilter(0&)
276                  blnIsHandlerInstalled = False
277              End If
End Sub

























