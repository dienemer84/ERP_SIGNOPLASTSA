Attribute VB_Name = "RegOCX"
Option Explicit

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" _
                                     (ByVal lpLibFileName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, _
                                                      ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, _
                                                      ByVal dwCreationFlags As Long, lpThreadID As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
                                                             ByVal dwMilliseconds As Long) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, _
                                                           lpExitCode As Long) As Long

Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Private Declare Function GetVersionExA Lib "kernel32" _
                                       (lpVersionInformation As OSVERSIONINFO) As Integer

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type


Public Sub RegistrarOCXs()
    If IsWin7OrVistaOr2003 Then Exit Sub

    Dim lLib As Long             ' Store handle of the control library
    Dim lpDLLEntryPoint As Long  ' Store the address of function called
    Dim lpThreadID As Long       ' Pointer that receives the thread identifier
    Dim lpExitCode As Long       ' Exit code of GetExitCodeThread
    Dim mThread
    Dim mresult

    Dim ocxs As New Collection
    Dim ocx As Variant

    ocxs.Add App.path & "\Codejock.CommandBars.v12.0.2.ocx"
    ocxs.Add App.path & "\Codejock.Controls.v12.0.2.ocx"
    ocxs.Add App.path & "\Codejock.ReportControl.v12.0.2.ocx"

    For Each ocx In ocxs
        If LenB(Dir(ocx)) > 0 Then
            lLib = LoadLibrary(ocx)
            If lLib <> 0 Then
                lpDLLEntryPoint = GetProcAddress(lLib, "DllRegisterServer")
                If lpDLLEntryPoint = vbNull Then GoTo prox

                mThread = CreateThread(ByVal 0, 0, ByVal lpDLLEntryPoint, ByVal 0, 0, lpThreadID)
                If mThread = 0 Then
                    FreeLibrary lLib
                    GoTo prox
                End If

                mresult = WaitForSingleObject(mThread, 10000)
                If mresult <> 0 Then
                    FreeLibrary lLib
                    lpExitCode = GetExitCodeThread(mThread, lpExitCode)
                    ExitThread lpExitCode
                    GoTo prox
                End If

                CloseHandle mThread
                FreeLibrary lLib

            End If
        End If
prox:
    Next ocx

End Sub


'http://msdn.microsoft.com/en-us/library/ms724834%28VS.85%29.aspx
Private Function IsWin7OrVistaOr2003() As Boolean
    Dim osInfo As OSVERSIONINFO
    Dim retvalue As Integer

    osInfo.dwOSVersionInfoSize = 148
    osInfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osInfo)

    IsWin7OrVistaOr2003 = (osInfo.dwMajorVersion = 6)
End Function
