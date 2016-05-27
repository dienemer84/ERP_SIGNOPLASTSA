Attribute VB_Name = "m_API"
'//**************************************************************************
'// ----------------- Module -----------------
'// Name        : --
'// Version     : --
'// Author      : Benoit Frigon
'// Created on  : 13-MAY-2002
'// Last update : 16-MAY-2002
'// File        : FrmTest
'// Desc.       : Drag controls API declaration
'//**************************************************************************
'// All rights reserved@Logiciels M.T.L enr. NEQ# 22-48153829(Québec)
'//**************************************************************************
Option Explicit


'//***************************************************************************************
'// Memory management API
'//***************************************************************************************
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function IsBadReadPtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Public Declare Function IsBadWritePtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Public Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long

Public Const GMEM_FIXED = &H0
Public Const GMEM_ZEROINIT = &H40
Public Const GMEM_MOVEABLE = &H2
Public Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
'//---------------------------------------------------------------------------------------



'//***************************************************************************************
'// Rect API
'//***************************************************************************************
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function EqualRect Lib "user32" (lpRect1 As RECT, lpRect2 As RECT) As Long
'//---------------------------------------------------------------------------------------



'//***************************************************************************************
'// Cursor API
'//***************************************************************************************
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Public Const IDC_ARROW = 32512&
Public Const IDC_IBEAM = 32513&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_SIZE = 32640&
Public Const IDC_CROSS = 32515&
Public Const IDC_APPSTARTING = 32650&
Public Const IDC_NO = 32648&
Public Const IDC_WAIT = 32514&
'//---------------------------------------------------------------------------------------



'//***************************************************************************************
'// Messaging API
'//***************************************************************************************
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'//--- Windows messages ---
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Public Const WM_MOUSEMOVE = &H200
Public Const WM_ERASEBKGND = &H14
Public Const WM_SETCURSOR = &H20
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31
Public Const WM_DESTROY = &H2
Public Const WM_LBUTTONUP = &H202
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORSTATIC = &H138
Public Const EM_SETSEL = &HB1
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_NCHITTEST = &H84
Public Const WM_PAINT = &HF
Public Const WM_PRINTCLIENT = &H318
Public Const WM_NCPAINT = &H85
Public Const WM_PRINT = &H317

'//--- WM_NCHITTEST return code ---
Public Const HTTRANSPARENT = (-1)
'//---------------------------------------------------------------------------------------



'//***************************************************************************************
'// Window API
'//***************************************************************************************
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWndParent As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long


'//--- SetWindowPos flags ---
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
'//--- GetWindowLong constants ---
Public Const GWL_WNDPROC = (-4)
'//--- Z-order constants ---
Public Const HWND_TOP = 0
'//--- Showwindow commands ---
Public Const SW_SHOW = 5
Public Const SW_HIDE = 0
'//--- Window styles ---
Public Const WS_CHILD = &H40000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_BORDER = &H800000
'//--- Editbox styles ---
Public Const ES_MULTILINE = &H4&
'//---------------------------------------------------------------------------------------




'//***************************************************************************************
'// Class API
'//***************************************************************************************
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long


Public Type WNDCLASS
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type
'//---------------------------------------------------------------------------------------



'//***************************************************************************************
'// Drawing API
'//***************************************************************************************
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hrgn As Long) As Long
Public Declare Function GetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Public Declare Function GetUpdateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function ValidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Any, ByVal fuRedraw As Long) As Long


'//--- RedrawWindow flags ---
Public Const RDW_ERASE = &H4
Public Const RDW_ERASENOW = &H200
Public Const RDW_INVALIDATE = &H1
Public Const RDW_UPDATENOW = &H100
'//--- Raster ops ---
Public Const R2_COPYPEN = 13    '  P
Public Const R2_NOTXORPEN = 10  '  DPxn
'//--- Pen styles ---
Public Const PS_SOLID = 0
'//--- Color codes ---
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14


Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type PAINTSTRUCT
    hdc As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(32) As Byte
End Type
'//---------------------------------------------------------------------------------------



'//***************************************************************************************
'// Version info
'//***************************************************************************************
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'//--- Platform ID ---
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32s = 0

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
'//---------------------------------------------------------------------------------------




'//***************************************************************************************
'// Misc functions
'//***************************************************************************************
Public Function GetAddress(Address As Long)
    GetAddress = Address
End Function




'//***************************************************************************************
'// Metric conversion
'//***************************************************************************************
Public Function ScreenRectToClient(hWnd As Long, lpRect As RECT)
    '//**** Convert Left and Top positions ****
    Dim Pt As POINTAPI
    Pt.x = lpRect.Left: Pt.y = lpRect.Top
    Call ScreenToClient(hWnd, Pt)

    Call OffsetRect(lpRect, (Pt.x - lpRect.Left), (Pt.y - lpRect.Top))
End Function
Public Function ClientRectToScreen(hWnd As Long, lpRect As RECT)
    '//**** Convert Left and Top positions ****
    Dim Pt As POINTAPI
    Pt.x = lpRect.Left: Pt.y = lpRect.Top
    Call ClientToScreen(hWnd, Pt)

    Call OffsetRect(lpRect, (Pt.x - lpRect.Left), (Pt.y - lpRect.Top))
End Function
Public Function PointInRect(lpPoint As POINTAPI, lpRect As RECT) As Boolean
    PointInRect = ((lpPoint.x >= lpRect.Left) And (lpPoint.x <= lpRect.Right) And _
                   (lpPoint.y >= lpRect.Top) And (lpPoint.y <= lpRect.Bottom))
End Function
Public Sub NormalizeRect(lpRect As RECT)
    If (lpRect.Right < lpRect.Left) Then Call Swap(lpRect.Right, lpRect.Left)
    If (lpRect.Bottom < lpRect.Top) Then Call Swap(lpRect.Bottom, lpRect.Top)
End Sub
Private Sub Swap(Num1 As Long, Num2 As Long)
    Dim Temp As Long
    Temp = Num1

    Num1 = Num2
    Num2 = Temp
End Sub



'//***************************************************************************************
'// Window functions
'//***************************************************************************************
Public Function GetClassNameEx(hWnd As Long) As String
    Dim sBuffer As String
    sBuffer = String(256, Chr(0))

    Dim Length As Long
    Length = GetClassName(hWnd, sBuffer, 256)

    GetClassNameEx = Left(sBuffer, Length)
End Function
Public Function GetWindowTextEx(hWnd As Long) As String
    Dim sBuffer As String
    sBuffer = String(256, Chr(0))

    Dim Length As Long
    Length = GetWindowText(hWnd, sBuffer, 256)

    GetWindowTextEx = Left(sBuffer, Length)
End Function



'//***************************************************************************************
'// API macro
'//***************************************************************************************
Public Function HIWORD(LongIn As Long) As Integer
    CopyMemory HIWORD, ByVal VarPtr(LongIn) + 2, 2
End Function
Public Function LOWORD(LongIn As Long) As Integer
    CopyMemory LOWORD, LongIn, 2
End Function
Public Function GET_X_LPARAM(lParam As Long) As Long
    GET_X_LPARAM = LOWORD(lParam)
End Function
Public Function GET_Y_LPARAM(lParam As Long) As Long
    GET_Y_LPARAM = HIWORD(lParam)
End Function


'//***************************************************************************************
'// Version info functions
'//***************************************************************************************
Public Function AreLargePatternSupported() As Boolean
    Dim osInfo As OSVERSIONINFO
    osInfo.dwOSVersionInfoSize = Len(osInfo)

    Call GetVersionEx(osInfo)
    Debug.Print osInfo.dwMinorVersion

    '//**** Is the OS version is Win98+ or Win2k+ ****
    If ((osInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS) And ((osInfo.dwMajorVersion >= 4) And (osInfo.dwMinorVersion > 0))) Or _
       ((osInfo.dwPlatformId = VER_PLATFORM_WIN32_NT) And ((osInfo.dwMajorVersion >= 4) And (osInfo.dwMinorVersion >= 0))) Then

        AreLargePatternSupported = True
    End If
End Function


