Attribute VB_Name = "CallStack"

Option Explicit

'##SUMMARY API for performance timer
Public Declare Function GetTickCount Lib "kernel32" () As Long

'##SUMMARY API to send debug messages to a debugger
'##REMARKS The project needs to have a conditional compilation argument of DEBUGLEVEL set to a bitmasked value.
'##REMARKS Values for DEBUGLEVEL:  1 = Error Messages, 2 = Call Stack Messages, 3 = Log Messages
Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

'##SUMMARY The private local variable holding the call stack
Private m_Stack() As StackType

'##SUMMARY Current size of the call stack
Private m_intStackSize As Integer

'##SUMMARY True if there has been a call to read or write to the call stack
Private m_blnStackInitialized As Boolean

'##SUMMARY The CallStack definition
Private Type StackType
    strModuleName As String
    strProcedureName As String
    strParamsText As String
    lngStartTime As Long
End Type

'##SUMMARY Project Global Variable containing the error number of the last error that occurred
Public g_lngLastErrorNumber As Long

'##SUMMARY Project Global Variable containing the error description of the last error that occurred
Public g_strLastError As String

'##SUMMARY Project Global Variable containing the last dll error number that occurred
Public g_lngLastDLLError As Long

'##SUMMARY Project Global Variable containing the error source of the last error that occurred
Public g_strLastSource As String

'##SUMMARY Project Global Variable indicating the time in seconds that the last call took to execute
Public g_dblLastCallTime As Double

'##SUMMARY Project Global Variable that tells the RunTimeError method to disable writing errors to a log file
'##REMARKS <P>The log file will be created in a sub directory off the _
 app.path called LogFiles. If the directory doesn't exist, it will be _
 created. The log file name will be in the following format.</P><P>YYYYMMDD_APPNAME.log</P>
Public g_blnDisableTextLog As Boolean

'##SUMMARY Project Global Variable that tells the RunTimeError method _
 to disable writing errors to the NT Event Log
Public g_blnDisableEventLog As Boolean

'##SUMMARY The default path to the log files.
'##REMARKS If not set, uses app.path\LogFiles
Public g_strDefaultLogPath As String

Public g_lngExitCode As Long

Public Sub RuntimeError( _
       ByVal ModuleName As String, _
       ByVal ProcedureName As String, _
       ByVal objErr As ErrObject, _
       Optional ByVal ErrLine As Long = 0, _
       Optional ByVal UserDescription As String = "" _
     )
Attribute RuntimeError.VB_Description = "Function called from error handler to log error"

'--------------------------------------------------------------------------------
' Procedure  : RuntimeError
' Created by : Paul Welter
' Date-Time  : 2/20/2002 3:08:24 PM
' ##SUMMARY Function called from error handler to log error
' ##PARAM ModuleName Name of the module the error occurred in
' ##PARAM ProcedureName Name of the method the error occurred in
' ##PARAM objErr The error object
' ##PARAM ErrLine The line number that the error occurred on. User _
  ERL and add line numbers to method to get this information.
' ##PARAM UserDescription Optional developer description of error
'--------------------------------------------------------------------------------

    Dim strMsgText As String    '  text of error report
    Dim strDebugMsg As String

    g_lngLastErrorNumber = objErr.Number
    g_strLastError = objErr.Description
    g_strLastSource = objErr.Source
    g_lngLastDLLError = objErr.LastDllError

    strDebugMsg = App.EXEName & " ERROR: " & ProcedureName & " - " & g_lngLastErrorNumber & " - " & g_strLastError

    Debug.Print strDebugMsg
    Debug.Assert False    'force break point, ignored when compiled

    On Error Resume Next

    #If DEBUGLEVEL And 1 Then
        OutputDebugString strDebugMsg
    #End If

    'error message
    strMsgText = vbCrLf & vbCrLf & Now & "  RunTime Error in " & App.EXEName & vbCrLf & vbCrLf
    strMsgText = strMsgText & "RunTime Error: " & CStr(g_lngLastErrorNumber) & " - " & _
                 g_strLastError
    strMsgText = strMsgText & _
                 vbCrLf & vbTab & "Module Name    : " & ModuleName & _
                 vbCrLf & vbTab & "Procedure Name : " & ProcedureName & _
                 vbCrLf & vbTab & "Line Number    : " & CStr(ErrLine) & _
                 vbCrLf & vbTab & "Error Source   : " & g_strLastSource & _
                 vbCrLf & vbTab & "User Desc.     : " & UserDescription & _
                 vbCrLf

    strMsgText = strMsgText & vbCrLf & "Call Stack:" & vbCrLf & StackRead

    'writing to event log
    If Not g_blnDisableEventLog Then
        Call App.LogEvent(strMsgText, vbLogEventTypeError)    'only works compiled
    End If

    'writing to log file
    If Not g_blnDisableTextLog Then
        WriteLogFile strMsgText
    End If

End Sub

Public Sub StackAdd( _
       ByVal ModuleName As String, _
       ByVal ProcedureName As String, _
       ParamArray ParamList() _
     )
Attribute StackAdd.VB_Description = "Adds a procedure to the debug call stack"

'--------------------------------------------------------------------------------
' Procedure  : StackAdd
' Created by : Paul Welter
' Date-Time  : 2/20/2002 2:45:14 PM
' ##SUMMARY Adds a procedure to the debug call stack
' ##PARAM ModuleName The name of the class or module
' ##PARAM ProcedureName The name of the method that is currently being executing
' ##PARAM ParamList List of input parameter for the current method
'--------------------------------------------------------------------------------
    On Error Resume Next

    Dim intCounter As Integer
    Dim strParams As String

    If Not m_blnStackInitialized Then
        ReDim m_Stack(20) As StackType    'defaulting Stack to 20 elements
        m_blnStackInitialized = True
    End If

    If m_intStackSize > UBound(m_Stack) Then
        'to prevent excessive redims, increases stack array in 20 element chunks
        ReDim Preserve m_Stack(m_intStackSize + 19) As StackType
    End If

    strParams = ""
    For intCounter = 0 To UBound(ParamList)
        strParams = strParams & vbTab & "Param_" & Right$(String$(3, "0") & (intCounter + 1), 3) & " : " & _
                    VarToString(ParamList(intCounter)) & vbCrLf
    Next

    With m_Stack(m_intStackSize)
        .strModuleName = ModuleName
        .strProcedureName = ProcedureName
        .strParamsText = strParams
        .lngStartTime = GetTickCount

        #If DEBUGLEVEL And 2 Then
            OutputDebugString "BEGIN " & App.EXEName & "." & .strModuleName & "." & .strProcedureName
        #End If
    End With

    m_intStackSize = m_intStackSize + 1

End Sub

Public Sub StackRemove()
Attribute StackRemove.VB_Description = "Removes the last call from the stack"
'--------------------------------------------------------------------------------
' Procedure  : StackRemove
' Created by : Paul Welter
' Date-Time  : 2/20/2002 3:07:08 PM
' ##SUMMARY Removes the last call from the stack
'--------------------------------------------------------------------------------

    On Error Resume Next

    Dim strMessage As String
    Dim lngEndTime As Long

    If Not m_blnStackInitialized Then
        ReDim m_Stack(20) As StackType    'defaulting Stack to 20 elements
        m_blnStackInitialized = True
    End If

    m_intStackSize = m_intStackSize - 1    'shrinking stack by 1
    If m_intStackSize < 0 Then m_intStackSize = 0

    lngEndTime = GetTickCount    'getting time
    g_dblLastCallTime = (lngEndTime - m_Stack(m_intStackSize).lngStartTime) / 1000    'calculating run time

    With m_Stack(m_intStackSize)
        strMessage = "END " & App.EXEName & "." & .strModuleName & "." & .strProcedureName & " Time: " & _
                     g_dblLastCallTime & " sec"    'performance log message
    End With

    #If DEBUGLEVEL And 2 Then
        OutputDebugString strMessage
    #End If

    'Debug.Print strMessage

End Sub

Public Function StackRead() As String
Attribute StackRead.VB_Description = "Formats the CallStack to a string message for error logging"
'--------------------------------------------------------------------------------
' Procedure  : StackRead
' Created by : Paul Welter
' Date-Time  : 2/20/2002 3:32:57 PM
' ##SUMMARY Formats the CallStack to a string message for error logging
' ##RETURNS Returns a formatted string from the call stack
'--------------------------------------------------------------------------------

    On Error Resume Next

    Dim strTemp As String
    Dim intCounter As Integer

    If Not m_blnStackInitialized Then
        ReDim m_Stack(20) As StackType    'defaulting Stack to 20 elements
        m_blnStackInitialized = True
    End If

    strTemp = ""

    For intCounter = m_intStackSize - 1 To 0 Step -1    'creating stack message
        With m_Stack(intCounter)
            strTemp = strTemp & _
                      vbTab & "Module    : " & .strModuleName & vbCrLf & _
                      vbTab & "Procedure : " & .strProcedureName & "()" & vbCrLf & _
                      .strParamsText & vbCrLf
        End With
    Next

    StackRead = strTemp

End Function

Public Function WriteLogFile( _
       ByVal v_strMessage As String, _
       Optional ByVal v_strPath As String = "" _
     ) As Boolean
Attribute WriteLogFile.VB_Description = "Writes text to a log file"

'--------------------------------------------------------------------------------
' Procedure  : WriteLogFile
' Created by : Paul Welter
' Date-Time  : 7/29/2002 - 9:33:40 AM
' ##SUMMARY Writes text to a log file
' ##PARAM v_strMessage The message to write to the log file
' ##PARAM v_strPath The path of the log file. Defaults to App.Path\Logfiles.
'--------------------------------------------------------------------------------

    On Error GoTo ExitPoint:

    Dim intFileNumber As Integer
    Dim strLogFileName As String
    Dim strTempPath As String

    #If DEBUGLEVEL And 4 Then
        OutputDebugString Left$(v_strMessage, 256)    'preventing really long strings from being passed
    #End If

    If v_strPath = "" Then
        'if there is a default folder use it, else use app.path\LogFiles
        If g_strDefaultLogPath <> "" Then
            v_strPath = g_strDefaultLogPath
        Else
            v_strPath = App.path & "\LogFiles\"
        End If
    End If

    If Right(v_strPath, 1) <> "\" Then
        v_strPath = v_strPath & "\"
    End If

    strLogFileName = Right$(String$(2, "0") & Year(Now), 4) & _
                     Right$(String$(2, "0") & Month(Now), 2) & _
                     Right$(String$(2, "0") & Day(Now), 2) & _
                     "_" & App.EXEName & ".log"

    intFileNumber = FreeFile

    strTempPath = Dir$(v_strPath, vbDirectory)    'looking for folder
    If strTempPath <> "." Then
        Call MkDir(v_strPath)    'creating folder
        If Err.Number <> 0 Then v_strPath = App.path & "\"
    End If

    Open v_strPath & strLogFileName For Append Access Write As #intFileNumber    ' Create file name.
    Print #intFileNumber, time & " " & v_strMessage    ' Output text.
    Close #intFileNumber   ' Close file.

    WriteLogFile = True

    Exit Function
ExitPoint:
    App.LogEvent Err.Description, vbLogEventTypeError

End Function
Public Sub ClearError()
Attribute ClearError.VB_Description = "Clears global error properties"
'--------------------------------------------------------------------------------
' Procedure  : ClearError
' Created by : Paul Welter
' Date-Time  : 9/17/2002 - 2:22:52 PM
' ##SUMMARY Clears global error properties
'--------------------------------------------------------------------------------

    On Error Resume Next

    g_lngLastErrorNumber = 0
    g_strLastError = ""
    g_lngLastDLLError = 0
    g_strLastSource = ""

End Sub

Private Function VarToString(ByVal v_varVariable As Variant) As String
Attribute VarToString.VB_Description = "Converts a variable to a string value"
'--------------------------------------------------------------------------------
' Procedure  : VarToString
' Created by : Paul Welter
' Date-Time  : 7/29/2002 - 9:34:20 AM
' ##SUMMARY Converts a variable to a string value
'--------------------------------------------------------------------------------

    On Error Resume Next

    If IsArray(v_varVariable) Then
        VarToString = "{Array}"
    Else
        Select Case VarType(v_varVariable)
        Case vbInteger, vbLong, vbByte, vbSingle, vbDouble, vbCurrency, vbBoolean, vbDecimal
            VarToString = CStr(v_varVariable)
        Case vbDate
            VarToString = "'" & CStr(v_varVariable) & "'"
        Case vbError
            VarToString = ""
        Case vbEmpty
            VarToString = "{Empty}"
        Case vbNull
            VarToString = "{Null}"
        Case vbString
            VarToString = "'" & v_varVariable & "'"
        Case vbObject
            VarToString = "{" & TypeName(v_varVariable) & "}"    'Value of Nothing will be shown as "Nothing"
        Case Else
            VarToString = "{?}"
        End Select
    End If

End Function


