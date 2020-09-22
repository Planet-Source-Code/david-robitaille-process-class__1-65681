VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'class process
'By David Robitaille

'Init the class with setProcess()
'start process with Run()
'Check the state of execution with Finished()
'Then check if you get an error with inError()
'and get strOutput (with getstrOutput(TRUE)) or/and strError when needed
'use StdInput(strInput As String) to write to the process
'close the process with clean

'Dont forget that you need to Check the state of execution with Finished()
'before getting anyting from strOutput, strError or intError






Option Explicit

Private strCommand ' command line to execute
Private noProcess As PROCESS_INFORMATION ' pointer to the process
Public strError As String  ' error message
Private strOutput As String ' dos output of the process
Private intState As Integer ' State of the process
Private strAwnserFile As String

Const NONE = 0     ' not active
Const RUNNING = 1  ' still executing
Const COMPLETE = 2 ' Finished (but remain active)
Const ERROR = -1   ' Error
Const CODEPAGE = 850 ' CODE PAGE

Private over As OVERLAPPED 'Not used direcly

Private hReadPipeStdOutput As Long 'The handle used to read from the pipe.
Private hWritePipeStdOutput As Long 'The pipe end where StdOutput will be redirected to.
 
Private hReadPipeStdError As Long 'The handle used to read from the pipe.
Private hWritePipeStdError As Long 'The pipe end where StdErr will be redirected to.
  
 
Private hReadPipeStdInput As Long 'The pipe end where StdInputput will be redirected to.
Private hWritePipeStdInput As Long 'The handle used to write in the pipe.
  
 
'--- Déclarations AppI

Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
        ByVal CODEPAGE As Long, _
        ByVal dwFlags As Long, _
        ByVal lpMultiByteStr As String, _
        ByVal cchMultiByte As Long, _
        ByVal lpWideCharStr As String, _
        ByVal cchWideChar As Long) As Long

Private Type STARTUPINFO
         cb As Long
         lpReserved As String
         lpDesktop As String
         lpTitle As String
         dwX As Long
         dwY As Long
         dwXSize As Long
         dwYSize As Long
         dwXCountChars As Long
         dwYCountChars As Long
         dwFillAttribute As Long
         dwFlags As Long
         wShowWindow As Integer
         cbReserved2 As Integer
         lpReserved2 As Long
         hStdInput As Long
         hStdOutput As Long
         hStdError As Long
End Type

Private Type PROCESS_INFORMATION
         hProcess As Long
         hThread As Long
         dwProcessID As Long
         dwThreadID As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type



Private Declare Function CreatePipe Lib "kernel32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As Any, _
    ByVal nSize As Long) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    lpOverlapped As OVERLAPPED) As Long

Private Declare Function WriteFile Lib "kernel32" _
       (ByVal hFile As Long, _
        ByVal lpBuffer As String, _
        ByVal cToWrite As Long, _
        ByRef cWritten As Long, _
        Optional ByVal lpOverlapped As Long) As Long
    
Private Declare Function PeekNamedPipe Lib "kernel32" ( _
    ByVal hNamedPipe As Long, _
    lpBuffer As String, _
    ByVal nBufferSize As Long, _
    lpBytesRead As Long, _
    lpTotalBytesAvail As Long, _
    lpBytesLeftThisMessage As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
    lpApplicationName As Long, ByVal lpCommandLine As String, _
    lpProcessAttributes As Any, lpThreadAttributes As Any, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As Any, lpProcessInformation As Any) As Long
    
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_HIDE = 0
Private Const INFINITE = -1&

'init process
Public Sub setProcess(strCommandLine As String)
    strCommand = strCommandLine
    intState = NONE
    strError = ""
    strOutput = ""
End Sub


'démarre le process
Public Sub Run()
    Dim Start As STARTUPINFO
    Dim ReturnValue As Integer
    Dim secur As SECURITY_ATTRIBUTES
    
    'reset the output
    strOutput = ""
    
    'to redirect the output and input in PIPE
    secur.nLength = Len(secur)
    secur.bInheritHandle = True
    secur.lpSecurityDescriptor = 0
    
    'StdOutput Pipe
    ReturnValue = CreatePipe(hReadPipeStdOutput, hWritePipeStdOutput, secur, 0)
    If ReturnValue = 0 Then
        intState = ERROR
        strError = "Error Function CreatePipe (StdOutput) : (" & Err.LastDllError & ")"
        Exit Sub
    End If
    
    'StdError Pipe
    ReturnValue = CreatePipe(hReadPipeStdError, hWritePipeStdError, secur, 0)
    If ReturnValue = 0 Then
        intState = ERROR
        strError = "Error Function CreatePipe (StdError) : (" & Err.LastDllError & ")"
        Exit Sub
    End If
    
    'StdInput Pipe
    ReturnValue = CreatePipe(hReadPipeStdInput, hWritePipeStdInput, secur, 0)
    If ReturnValue = 0 Then
        intState = ERROR
        strError = "Error Function CreatePipe (StdInput) : (" & Err.LastDllError & ")"
        Exit Sub
    End If
    
    'create "process"
    Start.cb = Len(Start)
    Start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW

    Start.wShowWindow = SW_HIDE
    Start.hStdOutput = hWritePipeStdOutput
    Start.hStdError = hWritePipeStdError
    Start.hStdInput = hReadPipeStdInput
    ReturnValue = CreateProcessA(0&, strCommand, secur, secur, True, _
                  NORMAL_PRIORITY_CLASS, 0&, 0&, Start, noProcess)
    
    'check if ok
    If ReturnValue = 1 Then
        intState = RUNNING
    Else
        intState = ERROR
        strError = "Error Function CreateProcessA : (" & Err.LastDllError & ")"
    End If
    
    CloseHandle hWritePipeStdOutput
    CloseHandle hWritePipeStdError
    CloseHandle hReadPipeStdInput
    over.offset = 0
    over.OffsetHigh = 0
    over.hEvent = 0
End Sub

'clean the process handles
Public Sub Clean()
    intState = NONE
    strError = ""
    CloseHandle hWritePipeStdInput
    CloseHandle hReadPipeStdOutput
    CloseHandle hReadPipeStdError
    CloseHandle noProcess.hProcess
    CloseHandle noProcess.hThread
End Sub

Public Function Finished() As Boolean
    'update state before rerurnig if finished
    updateState
    Finished = ((intState = COMPLETE) Or (intState = ERROR))
End Function

Public Function inError() As Boolean
    inError = (intState = ERROR)
End Function


Private Function updateState()
    Dim ReturnValue As Long
    'need to cleen the out pipes before
    UpdateOutput
    'if executed
    If intState = RUNNING Then
        'check if finished
        ReturnValue = WaitForSingleObject(noProcess.hProcess, 0)
        If ReturnValue = 0 Then
            'finished, check the ERORRLEVEL DOS
            GetExitCodeProcess noProcess.hProcess, ReturnValue
            If ReturnValue = 0 Then
                intState = COMPLETE
            Else
                intState = ERROR
                If strError = "" Then strError = strError & vbCrLf
                strError = strError & "Error number " & CStr(ReturnValue) & " at the end of exec"
            End If
            'CloseHandle hReadPipe
        Else
            'not finished, but in error; else, do nothing
            If ReturnValue <> 258 Then
               'hang or crash
               intState = ERROR
               If strError = "" Then strError = strError & vbCrLf
               strError = "Error in exec"
               'CloseHandle hReadPipe
            End If
        End If
    End If
End Function

Private Sub UpdateOutput()
    Dim ReturnValue As Long
    Dim lngBytesRead As Long
    Dim lngByteToRead As Long
    Dim lngByteLeft As Long
    Dim sbuffer As String * 256
    Dim sbufferUnicode As String * 512
    Dim stemp As String * 1
    
    'transfer the output form pipe at the end of the string strOutput.
    
    Do
        'check if there something to reed before (solve a haging problem)
        ReturnValue = PeekNamedPipe(hReadPipeStdOutput, stemp, 0, lngBytesRead, lngByteToRead, lngByteLeft)
        If (lngByteToRead <> 0) Then
            If (lngByteToRead > 256) Then
                lngByteToRead = 256
            End If
            ReturnValue = ReadFile(hReadPipeStdOutput, sbuffer, lngByteToRead, lngBytesRead, over)
            ReturnValue = MultiByteToWideChar(CODEPAGE, 0, sbuffer, lngBytesRead, sbufferUnicode, 512)
            strOutput = strOutput & Left(StrConv(sbufferUnicode, vbFromUnicode), lngBytesRead)
        End If
    Loop While lngBytesRead <> 0 ' if lngBytesRead = 0 then there is no more characters to read

    'transfer the error form pipe at the end of the string strError.
    Do
        'check if there something to reed before (solve a haging problem)
        ReturnValue = PeekNamedPipe(hReadPipeStdError, stemp, 0, lngBytesRead, lngByteToRead, lngByteLeft)
        If (lngByteToRead <> 0) Then
            If (lngByteToRead > 256) Then
                lngByteToRead = 256
            End If
            ReturnValue = ReadFile(hReadPipeStdError, sbuffer, lngByteToRead, lngBytesRead, over)
            ReturnValue = MultiByteToWideChar(CODEPAGE, 0, sbuffer, lngBytesRead, sbufferUnicode, 512)
            strError = strError & Left(StrConv(sbufferUnicode, vbFromUnicode), lngBytesRead)
        End If
    Loop While lngBytesRead <> 0 ' if lngBytesRead = 0 then there is no more characters to read


End Sub

Public Function strGetOutput(blnComplet) As String
    If ((Len(strOutput) < 100) Or blnComplet) Then
                strGetOutput = strOutput
    Else
                strGetOutput = Left(strOutput, 50) + " [...] " + Right(strOutput, 50)
    End If
End Function

Public Sub StdInput(strInput As String)
    Dim ReturnValue As Long
    Dim lngBytesWrited As Long
    Dim ba() As Byte
    Dim i As Integer
    
    
    ReturnValue = WriteFile(hWritePipeStdInput, strInput, Len(strInput), lngBytesWrited)
    
    If ReturnValue = 0 Then
        intState = ERROR
        If strError = "" Then strError = strError & vbCrLf
        strError = "Error Function WriteFile : (" & Err.LastDllError & ")" & Err.Description
    End If
End Sub




