Attribute VB_Name = "M95"
Option Explicit
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'ShellAndWait 'Chip Pearson
    Option Compare Text
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Private Const SYNCHRONIZE = &H100000
    Public Enum ShellAndWaitResult: Success = 0: Failure = 1: Timeout = 2: InvalidParameter = 3: SysWaitAbandoned = 4: UserWaitAbandoned = 5: UserBreak = 6: End Enum
    Public Enum ActionOnBreak: IgnoreBreak = 0: AbandonWait = 1: PromptUser = 2: End Enum
    Private Const STATUS_ABANDONED_WAIT_0 As Long = &H80
    Private Const STATUS_WAIT_0 As Long = &H0
    Private Const WAIT_ABANDONED As Long = (STATUS_ABANDONED_WAIT_0 + 0)
    Private Const WAIT_OBJECT_0 As Long = (STATUS_WAIT_0 + 0)
    Private Const WAIT_TIMEOUT As Long = 258&
    Private Const WAIT_FAILED As Long = &HFFFFFFFF
    Private Const WAIT_INFINITE = -1&
'GetImageProperties
    Enum ImagePropertyTypeEnum: ImageWidth: ImageHeight: HorizontalResolution: VerticalResolution: IsAnimated: ActiveFrame: FrameCount: FileExtension: HasTransparency: PixelDepth: End Enum
'ShellAndWaitWithFlashingCmdWindow
    Private Type STARTUPINFO: cb As Long: lpReserved As String: lpDesktop As String: lpTitle As String: dwX As Long: dwY As Long: dwXSize As Long: dwYSize As Long: dwXCountChars As Long: dwYCountChars As Long: dwFillAttribute As Long: dwFlags As Long: wShowWindow As Integer: cbReserved2 As Integer: lpReserved2 As Long: hStdInput As Long: hStdOutput As Long: hStdError As Long: End Type
    Private Type PROCESS_INFORMATION: hProcess As Long: hThread As Long: dwProcessId As Long: dwThreadID As Long: End Type
    'Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
    'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Private Const NORMAL_PRIORITY_CLASS = &H20&
    Private Const INFINITE = -1&

Public Function ShellAndWait(ShellCommand$, TimeOutMs&, ShellWindowState As VbAppWinStyle, BreakKey As ActionOnBreak) As ShellAndWaitResult
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ShellAndWait
    '
    ' This function calls Shell and passes to it the command text in ShellCommand. The function
    ' then waits for TimeOutMs (in milliseconds) to expire.
    '
    '   Parameters:
    '       ShellCommand
    '           is the command text to pass to the Shell function.
    '
    '       TimeOutMs
    '           is the number of milliseconds to wait for the shell'd program to wait. If the
    '           shell'd program terminates before TimeOutMs has expired, the function returns
    '           ShellAndWaitResult.Success = 0. If TimeOutMs expires before the shell'd program
    '           terminates, the return value is ShellAndWaitResult.TimeOut = 2.
    '
    '       ShellWindowState
    '           is an item in VbAppWinStyle specifying the window state for the shell'd program.
    '
    '       BreakKey
    '           is an item in ActionOnBreak indicating how to handle the application's cancel key
    '           (Ctrl Break). If BreakKey is ActionOnBreak.AbandonWait and the user cancels, the
    '           wait is abandoned and the result is ShellAndWaitResult.UserWaitAbandoned = 5.
    '           If BreakKey is ActionOnBreak.IgnoreBreak, the cancel key is ignored. If
    '           BreakKey is ActionOnBreak.PromptUser, the user is given a ?Continue? message. If the
    '           user selects "do not continue", the function returns ShellAndWaitResult.UserBreak = 6.
    '           If the user selects "continue", the wait is continued.
    '
    '   Return values:
    '            ShellAndWaitResult.Success = 0
    '               indicates the the process completed successfully.
    '            ShellAndWaitResult.Failure = 1
    '               indicates that the Wait operation failed due to a Windows error.
    '            ShellAndWaitResult.TimeOut = 2
    '               indicates that the TimeOutMs interval timed out the Wait.
    '            ShellAndWaitResult.InvalidParameter = 3
    '               indicates that an invalid value was passed to the procedure.
    '            ShellAndWaitResult.SysWaitAbandoned = 4
    '               indicates that the system abandoned the wait.
    '            ShellAndWaitResult.UserWaitAbandoned = 5
    '               indicates that the user abandoned the wait via the cancel key (Ctrl+Break).
    '               This happens only if BreakKey is set to ActionOnBreak.AbandonWait.
    '            ShellAndWaitResult.UserBreak = 6
    '               indicates that the user broke out of the wait after being prompted with
    '               a ?Continue message. This happens only if BreakKey is set to
    '               ActionOnBreak.PromptUser.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim TaskID&, ProcHandle&, WaitRes&, Ms&, MsgRes As VbMsgBoxResult
    Dim SaveCancelKey As XlEnableCancelKey, ElapsedTime&, Quit As Boolean
    Const ERR_BREAK_KEY = 18: Const DEFAULT_POLL_INTERVAL = 500
    
    If Trim(ShellCommand) = vbNullString Then
        ShellAndWait = ShellAndWaitResult.InvalidParameter
        Exit Function
    End If
    
    If TimeOutMs < 0 Then
        ShellAndWait = ShellAndWaitResult.InvalidParameter
        Exit Function
    ElseIf TimeOutMs = 0 Then
        Ms = WAIT_INFINITE
    Else
        Ms = TimeOutMs
    End If
    
    Select Case BreakKey
        Case AbandonWait, IgnoreBreak, PromptUser
            ' valid
        Case Else
            ShellAndWait = ShellAndWaitResult.InvalidParameter
            Exit Function
    End Select
    
    Select Case ShellWindowState
        Case vbHide, vbMaximizedFocus, vbMinimizedFocus, vbMinimizedNoFocus, vbNormalFocus, vbNormalNoFocus
            ' valid
        Case Else
            ShellAndWait = ShellAndWaitResult.InvalidParameter
            Exit Function
    End Select
    
    On Error Resume Next
    Err.Clear
    TaskID = shell(ShellCommand, ShellWindowState)
    If (Err.Number <> 0) Or (TaskID = 0) Then
        ShellAndWait = ShellAndWaitResult.Failure
        Exit Function
    End If
    
    ProcHandle = OpenProcess(SYNCHRONIZE, False, TaskID)
    If ProcHandle = 0 Then
        ShellAndWait = ShellAndWaitResult.Failure
        Exit Function
    End If
    
    On Error GoTo errH:
    SaveCancelKey = Application.EnableCancelKey
    Application.EnableCancelKey = xlErrorHandler
    WaitRes = WaitForSingleObject(ProcHandle, DEFAULT_POLL_INTERVAL)
        Do Until WaitRes = WAIT_OBJECT_0
            DoEvents
            Select Case WaitRes
                Case WAIT_ABANDONED
                    ' Windows abandoned the wait
                    ShellAndWait = ShellAndWaitResult.SysWaitAbandoned
                    Exit Do
                Case WAIT_OBJECT_0
                    ' Successful completion
                    ShellAndWait = ShellAndWaitResult.Success
                    Exit Do
                Case WAIT_FAILED
                    ' attach failed
                    ShellAndWait = ShellAndWaitResult.Failure
                    Exit Do
                Case WAIT_TIMEOUT
                    ' Wait timed out. Here, this time out is on DEFAULT_POLL_INTERVAL.
                    ' See if ElapsedTime is greater than the user specified wait
                    ' time out. If we have exceed that, get out with a TimeOut status.
                    ' Otherwise, reissue as wait and continue.
                    ElapsedTime = ElapsedTime + DEFAULT_POLL_INTERVAL
                    If Ms > 0 Then
                        ' user specified timeout
                        If ElapsedTime > Ms Then
                            ShellAndWait = ShellAndWaitResult.Timeout
                            Exit Do
                        Else
                            ' user defined timeout has not expired.
                        End If
                    Else
                        ' infinite wait -- do nothing
                    End If
                    ' reissue the Wait on ProcHandle
                    WaitRes = WaitForSingleObject(ProcHandle, DEFAULT_POLL_INTERVAL)
                    
                Case Else
                    ' unknown result, assume failure
                    ShellAndWait = ShellAndWaitResult.Failure
                    Exit Do
                    Quit = True
            End Select
        Loop
        
        CloseHandle ProcHandle
        Application.EnableCancelKey = SaveCancelKey
        Exit Function
    
errH:
    Debug.Print "ErrH: Cancel: " & Application.EnableCancelKey
    If Err.Number = ERR_BREAK_KEY Then
        If BreakKey = ActionOnBreak.AbandonWait Then
            CloseHandle ProcHandle
            ShellAndWait = ShellAndWaitResult.UserWaitAbandoned
            Application.EnableCancelKey = SaveCancelKey
            Exit Function
        ElseIf BreakKey = ActionOnBreak.IgnoreBreak Then
            Err.Clear
            Resume
        ElseIf BreakKey = ActionOnBreak.PromptUser Then
            MsgRes = MsgBox("User Process Break." & vbCrLf & _
                "Continue to wait?", vbYesNo)
            If MsgRes = vbNo Then
                CloseHandle ProcHandle
                ShellAndWait = ShellAndWaitResult.UserBreak
                Application.EnableCancelKey = SaveCancelKey
            Else
                Err.Clear
                Resume Next
            End If
        Else
            CloseHandle ProcHandle
            Application.EnableCancelKey = SaveCancelKey
            ShellAndWait = ShellAndWaitResult.Failure
        End If
    Else
        ' some other error. assume failure
        CloseHandle ProcHandle
        ShellAndWait = ShellAndWaitResult.Failure
    End If
    
    Application.EnableCancelKey = SaveCancelKey
End Function

Public Sub ShellAndWaitWithFlashingCmdWindow(cmdline As String)
    'Shell&Wait-Variante, ShortPopUpOfCmdWindow
    Dim proc As PROCESS_INFORMATION, Start As STARTUPINFO, ReturnValue As Integer
    'Initialize the STARTUPINFO structure:
        Start.cb = Len(Start)
    'Start the shelled application:
        ReturnValue = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, Start, proc)
    'Wait for the shelled application to finish:
        Do
            ReturnValue = WaitForSingleObject(proc.hProcess, 0)
            DoEvents
        Loop Until ReturnValue <> 258
    ReturnValue = CloseHandle(proc.hProcess)
End Sub

Function GetImageProperties(ByVal file As Variant, ByVal PropertyType As ImagePropertyTypeEnum)
        Dim TargetImage As Object
    If TypeName(file) = "IImageFile" Then
        Set TargetImage = file
    ElseIf TypeName(file) = "String" Then
        If Len(Dir(file)) = 0 Then Exit Function
        ' Late-binding option
        Set TargetImage = CreateObject("WIA.ImageFile")
        ' Early-binding option
        ' Dim TargetImage As New WIA.ImageFile
        TargetImage.LoadFile file
    End If
    With TargetImage
        GetImageProperties = Trim(Array(.Width, .Height, .HorizontalResolution, .VerticalResolution, _
                                        CBool(.IsAnimated), .ActiveFrame, .FrameCount, .FileExtension, _
                                        CBool(.IsAlphaPixelFormat), .PixelDepth)(PropertyType))
    End With
    Set TargetImage = Nothing
End Function

Sub TEST_GetImageProperties(ByVal file As Variant)
    Dim PropertyHeadings  As Variant
    PropertyHeadings = Array("Width (pixels)", "Height (pixels)", "Horizontal Resolution", "Vertical Resolution", _
                             "Animated", "Active Frame", "Frame Count", "File Extension", _
                             "Is Alpha Pixel Format", "Pixel Depth")
    Dim Counter As Long
    Dim Heading As String * 22
    For Counter = LBound(PropertyHeadings) To UBound(PropertyHeadings)
        Heading = PropertyHeadings(Counter)
        Debug.Print Counter + 1, Heading, GetImageProperties(file, Counter)
    Next
End Sub



