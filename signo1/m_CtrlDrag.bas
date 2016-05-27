Attribute VB_Name = "m_CtrlDrag"
'//**************************************************************************
'// ----------------- Module -----------------
'// Name        : --
'// Version     : 2.0 (BETA)
'// Author      : Benoit Frigon
'// Created on  : 13-MAY-2002
'// Last update : 10-JUL-2002
'// File        : m_CtrlDrag.bas
'// Desc.       : Drag controls at run-time (2.0 rev 0005)
'//**************************************************************************
'// All rights reserved@Logiciels M.T.L enr. NEQ# 22-48153829(Québec)
'//**************************************************************************
'//
'//                  = = = = =  Report bugs  = = = = =
'//
'// To report bugs, send an email to mtlsoftware@idz.net
'//
'// Include : - EXACTLY what it does
'//           - The OS version you're using
'//           - What you do to recreate the situation
'//
'// The most details you give me, the easier it is for me to correct the bug.
'// I can't do anything with messages like : "It doesn't work...".
'//**************************************************************************
Option Explicit


'//**************************************************************************
'// Constants
'//**************************************************************************
'//--- Class names ---
Private Const ClassName_GrabBox = "MTLSOFT_GrabBox20"
'//--- Stored properties name ---
Private Const PropName_PrevWndProc = "PrevWndProc"
Private Const PropName_DragEnabled = "DragEnabled"
Private Const PropName_HwndGrab = "HwndGrab"
Private Const PropName_GrabBoxID = "GrabBoxID"
Private Const PropName_SelectedHwnd = "SelectedHwnd"
Private Const PropName_AcceptDragDrop = "AcceptDragDrop"
Private Const PropName_AllowEdit = "AllowEdit"
Private Const PropName_ClassPtr = "ClassPtr"
Private Const PropName_ShowGrid = "ShowGrid"
Private Const PropName_SnapToGrid = "SnapToGrid"
Private Const PropName_GridSize = "GridSize"
Private Const PropName_GridBrush = "GridBrush"
Private Const PropName_GridBrushBMP = "GridBrushBMP"
Private Const PropName_ObjPtr = "ObjectPtr"
'//--- Enumeration actions ---
Private Const EnumMode_EnableDrag = 1
Private Const EnumMode_DisableDrag = 2
Private Const EnumMode_UnSubclass = 3
'//--- Metrics ---
Private Const Metrics_GrabBoxWidth = 7
'//--- Drag mode ---
Private Const DragMode_Move = 0
Private Const DragMode_SizeNW = 1
Private Const DragMode_SizeN = 2
Private Const DragMode_SizeNE = 3
Private Const DragMode_SizeW = 4
Private Const DragMode_SizeE = 5
Private Const DragMode_SizeSW = 6
Private Const DragMode_SizeS = 7
Private Const DragMode_SizeSE = 8
'//--- Default properties ---
Private Const Default_GridSize = 8



'//**************************************************************************
'// Variable
'//**************************************************************************
Private ContainerList As String
Private m_GrabBoxInit As Boolean
Private m_hdcScreen As Long
Private m_DragRc As RECT
Private m_hDragPen As Long
Private m_hOldPen As Long
Private m_DrawStatus As Long
Private DragOriginPt As POINTAPI
Private m_OnDrag As Boolean
Private m_DropContainerHwnd As Long
Private m_DragMode As Long
Private m_EditboxHwnd As Long
Private m_ActiveContainer As Long
Private m_ActiveObject As Long
Private m_SnapRc As RECT
Private m_InvalidMove As Boolean


'//**************************************************************************
'// Properties
'//**************************************************************************
Property Let GridSize(Container As Object, GridSize As Long)
    '//**** Get the handle of the container ****
    Dim hWndContainer As Long
    hWndContainer = GetContainerHwnd(Container)
    If (hWndContainer = 0) Then Exit Property

    If (GridSize < 3) Then GridSize = 3
    If (GridSize > 256) Then GridSize = 256

    '//**** If the OS is windows 95, restrict the grid size to 8x8 ****
    If (Not AreLargePatternSupported) Then
        If (GridSize > 8) Then GridSize = 8
    End If

    Call SetProp(hWndContainer, PropName_GridSize, GridSize)

    '//**** Delete the previous brush ****
    Dim hBrush As Long
    hBrush = GetProp(hWndContainer, PropName_GridBrush)
    If (hBrush <> 0) Then
        Call DeleteObject(hBrush)
        Call DeleteObject(GetProp(hWndContainer, PropName_GridBrushBMP))
    End If

    '//**** Create a new grid brush ****
    Call SetProp(hWndContainer, PropName_GridBrush, CreateGridBrush(GridSize))

    '//**** Refresh the container window ****
    Call RefreshContainer(hWndContainer)
End Property
Property Get GridSize(Container As Object) As Long
    '//**** Get the handle of the container ****
    Dim hWndContainer As Long
    hWndContainer = GetContainerHwnd(Container)
    If (hWndContainer = 0) Then Exit Property

    GridSize = GetProp(hWndContainer, PropName_GridSize)
End Property
Property Let ShowGrid(Container As Object, ShowGrid As Boolean)
    '//**** Get the handle of the container ****
    Dim hWndContainer As Long
    hWndContainer = GetContainerHwnd(Container)
    If (hWndContainer = 0) Then Exit Property


    If (GetProp(hWndContainer, PropName_GridSize) < 2) Then
        GridSize(Container) = 8
    End If

    Call SetProp(hWndContainer, PropName_ShowGrid, IIf(ShowGrid, 1, 0))

    '//**** Refresh the container window ****
    Call RefreshContainer(hWndContainer)
End Property
Property Get ShowGrid(Container As Object) As Boolean
    '//**** Get the handle of the container ****
    Dim hWndContainer As Long
    hWndContainer = GetContainerHwnd(Container)
    If (hWndContainer = 0) Then Exit Property

    ShowGrid = IIf(GetProp(hWndContainer, PropName_ShowGrid) <> 0, True, False)
End Property
Property Let SnapToGrid(Container As Object, SnapToGrid As Boolean)
    '//**** Get the handle of the container ****
    Dim hWndContainer As Long
    hWndContainer = GetContainerHwnd(Container)
    If (hWndContainer = 0) Then Exit Property

    Call SetProp(hWndContainer, PropName_SnapToGrid, IIf(SnapToGrid, 1, 0))
End Property
Property Get SnapToGrid(Container As Object) As Boolean
    '//**** Get the handle of the container ****
    Dim hWndContainer As Long
    hWndContainer = GetContainerHwnd(Container)
    If (hWndContainer = 0) Then Exit Property

    SnapToGrid = IIf(GetProp(hWndContainer, PropName_SnapToGrid) <> 0, True, False)
End Property
Property Let AcceptDragDrop(Container As Object, Accept As Boolean)
    '//**** Get the handle of the container ****
    Dim hWndContainer As Long
    hWndContainer = GetContainerHwnd(Container)
    If (hWndContainer = 0) Then Exit Property

    Call SetProp(hWndContainer, PropName_AcceptDragDrop, IIf(Accept, 1, 0))
End Property
Property Get AcceptDragDrop(Container As Object) As Boolean
    '//**** Get the handle of the container ****
    Dim hWndContainer As Long
    hWndContainer = GetContainerHwnd(Container)
    If (hWndContainer = 0) Then Exit Property

    AcceptDragDrop = IIf(GetProp(hWndContainer, PropName_AcceptDragDrop) <> 0, True, False)
End Property
Property Let AllowEdit(Container As Object, Allow As Boolean)
    '//**** Get the handle of the container ****
    Dim hWndContainer As Long
    hWndContainer = GetContainerHwnd(Container)
    If (hWndContainer = 0) Then Exit Property

    Call SetProp(hWndContainer, PropName_AllowEdit, IIf(Allow, 1, 0))
End Property
Property Get AllowEdit(Container As Object) As Boolean
    '//**** Get the handle of the container ****
    Dim hWndContainer As Long
    hWndContainer = GetContainerHwnd(Container)
    If (hWndContainer = 0) Then Exit Property

    AllowEdit = IIf(GetProp(hWndContainer, PropName_AllowEdit), True, False)
End Property
Private Function GetContainerHwnd(Container As Object) As Long
    '//**** Get the handle of the container ****
    On Error Resume Next
    GetContainerHwnd = Container.hWnd
    On Local Error GoTo 0
End Function



'//**************************************************************************
'// Container functions
'//**************************************************************************
Private Sub RefreshContainer(hWnd As Long)


    Call RedrawWindow(hWnd, ByVal 0&, ByVal 0&, RDW_ERASE Or RDW_ERASENOW Or RDW_INVALIDATE)
End Sub
Public Function InitializeContainer(Container As Object, Optional InitializeAllChild As Boolean = True, Optional AcceptDragDrop, Optional AllowEdit, Optional EventObject As ClsEvents) As Boolean
    '//**** Get the handle of the container ****
    On Error Resume Next
    Dim hWnd As Long
    hWnd = Container.hWnd
    On Local Error GoTo 0
    If (hWnd = 0) Then Exit Function

    '//**** Check if the type of container is a form or picture box (can't handle other type of container) ****
    If Not ((TypeOf Container Is Form) Or (TypeOf Container Is PictureBox)) Then
        Exit Function
    End If

    '//**** This control is already subclassed ****
    If (GetProp(hWnd, PropName_PrevWndProc) <> 0) Then
        Exit Function
    End If

    '//**** Get the current window procedure address ****
    Dim prevWndProc As Long
    prevWndProc = GetWindowLong(hWnd, GWL_WNDPROC)

    '//**** Store this address ****
    Call SetProp(hWnd, PropName_PrevWndProc, prevWndProc)

    '//**** Set the new window procedure address ****
    Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProcContainer)

    '//**** Add this window to the container list ****
    Call AddToContainerList(hWnd)

    '//**** Get the pointer to the event object ****
    If Not (EventObject Is Nothing) Then

        Dim ClassPtr As Long
        ClassPtr = ObjPtr(EventObject)
    End If

    '//**** Set properties ****
    Call SetProp(hWnd, PropName_ClassPtr, ClassPtr)
    Call SetProp(hWnd, PropName_ObjPtr, ObjPtr(Container))
    If (Not IsMissing(AcceptDragDrop)) Then
        Call SetProp(hWnd, PropName_AcceptDragDrop, IIf(AcceptDragDrop, 1, 0))
    End If
    If (Not IsMissing(AllowEdit)) Then
        Call SetProp(hWnd, PropName_AllowEdit, IIf(AllowEdit, 1, 0))
    End If

    Call EnableContainerDrag(hWnd, True)

    '//**** Create grab box ****
    Call RegisterGrabBoxes
    Call CreateGrabBoxes(hWnd)

    If (InitializeAllChild) Then
        Call EnumChildWindows(hWnd, AddressOf EnumChildProc, EnumMode_EnableDrag)
    End If

    InitializeContainer = True
End Function
Private Sub AddToContainerList(hWnd As Long)
    ContainerList = ContainerList & Chr(1) & hWnd & Chr(2)
End Sub
Private Sub RemoveFromContainerList(hWnd As Long)
    Dim lStart As Long
    lStart = InStr(ContainerList, Chr(1) & hWnd & Chr(2))
    If (lStart = 0) Then Exit Sub

    Dim lEnd As Long
    lEnd = InStr(lStart, ContainerList, Chr(2))

    ContainerList = Left(ContainerList, lStart - 1) & Mid(ContainerList, lEnd + 1)
End Sub
Public Sub UnInitializeContainer(Container As Object)
    '//**** Get the handle of the container ****
    On Error Resume Next
    Dim hWnd As Long
    hWnd = Container.hWnd
    On Local Error GoTo 0
    If (hWnd = 0) Then Exit Sub

    Call UnInitializeContainerEx(hWnd)
End Sub
Private Sub UnInitializeContainerEx(hWnd As Long)
    Dim prevWndProc As Long
    prevWndProc = GetProp(hWnd, PropName_PrevWndProc)
    If (prevWndProc = 0) Then Exit Sub

    '//**** Restore the old procedure ****
    Call SetWindowLong(hWnd, GWL_WNDPROC, prevWndProc)

    '//**** Remove properties ****
    Call RemoveProp(hWnd, PropName_PrevWndProc)
    Call RemoveProp(hWnd, PropName_AcceptDragDrop)
    Call RemoveProp(hWnd, PropName_ClassPtr)
    Call RemoveProp(hWnd, PropName_ObjPtr)
    Call RemoveProp(hWnd, PropName_GridSize)
    Call RemoveProp(hWnd, PropName_SnapToGrid)
    Call RemoveProp(hWnd, PropName_ShowGrid)
    Call RemoveProp(hWnd, PropName_GridBrush)

    '//**** Remove this container from the list ****
    Call RemoveFromContainerList(hWnd)

    '//**** Delete the grid brush ****
    Dim hBrush As Long
    hBrush = GetProp(hWnd, PropName_GridBrush)
    If (hBrush <> 0) Then
        Call DeleteObject(hBrush)
        Call DeleteObject(GetProp(hWnd, PropName_GridBrushBMP))
    End If

    '//**** Unsubclass all children ****
    Call UnSubclassAllChild(hWnd)

    '//**** Destroy the grab boxes ****
    Call DestroyGrabBoxes(hWnd)
    If (ContainerList = "") Then UnRegisterGrabBoxes

    If (GetParent(m_EditboxHwnd) = hWnd) Then
        Call EndEditMode(True)
    End If
End Sub
Public Function UnInitializeAllContainer()

    Do
        Dim lEnd As Long
        lEnd = InStr(ContainerList, Chr(2))
        If (lEnd = 0) Then Exit Do

        Dim hWnd As Long
        hWnd = Val(Mid(ContainerList, 2, (lEnd - 2)))

        Call UnInitializeContainerEx(hWnd)
    Loop Until (ContainerList = "")
End Function
Public Function EnableContainerDrag(hWndContainer As Long, Enabled As Boolean)
    Call SetProp(hWndContainer, PropName_DragEnabled, IIf(Enabled, 1, 0))

    If (Not Enabled) Then
        Call HideGrabBoxes(hWndContainer)
    End If
End Function
Private Function isContainerSupportEvents(hWndContainer As Long) As Boolean

    isContainerSupportEvents = (GetProp(hWndContainer, PropName_ClassPtr) <> 0)
End Function
Private Function GetEventObject(hWndContainer As Long) As ClsEvents
    Dim ClassPtr As Long
    ClassPtr = GetProp(hWndContainer, PropName_ClassPtr)

    If (ClassPtr <> 0) Then

        Dim ObjTemp As ClsEvents
        CopyMemory ObjTemp, ClassPtr, 4

        Set GetEventObject = ObjTemp

        CopyMemory ObjTemp, 0&, 4
    End If
End Function



'//**************************************************************************
'// Edit functions
'//**************************************************************************
Private Sub BeginEditMode(hWndContainer As Long, hWnd As Long)
    If (GetProp(hWndContainer, PropName_AllowEdit) = 0) Then Exit Sub

    '//**** Get the caption of this window ****
    Dim sCaption As String
    sCaption = GetWindowTextEx(hWnd)

    '//**** Check if there is an event handler ****
    If (isContainerSupportEvents(hWndContainer)) Then

        '//**** If the user cancel this action, exit ****
        If (GetEventObject(hWndContainer).EventBeforeEdit(hWndContainer, hWnd)) Then
            Exit Sub
        End If
    End If

    '//**** Destroy the previous edit box ****
    If (m_EditboxHwnd <> 0) Then Call EndEditMode(True)
    m_ActiveObject = hWnd
    m_ActiveContainer = hWndContainer


    '//**** Get the font of the window to be edited ****
    Dim hFont As Long
    hFont = SendMessage(hWnd, WM_GETFONT, ByVal 0&, ByVal 0&)

    '//**** Get the rect of the window to be edited ****
    Dim WindowRc As RECT
    Call GetWindowRect(hWnd, WindowRc)
    Call ScreenRectToClient(hWndContainer, WindowRc)

    '//**** Create the edit box ****
    m_EditboxHwnd = CreateWindowEx(0, "EDIT", sCaption, WS_CHILD Or WS_BORDER Or ES_MULTILINE, WindowRc.Left, WindowRc.Top, (WindowRc.Right - WindowRc.Left), (WindowRc.Bottom - WindowRc.Top), hWndContainer, 0, 0, 0)
    If (m_EditboxHwnd = 0) Then Exit Sub

    '//**** Apply the font to the edit box ****
    If (hFont <> 0) Then
        Call SendMessage(m_EditboxHwnd, WM_SETFONT, hFont, True)
    End If

    '//**** Show the edit box and set it on top ****
    Call SetWindowPos(m_EditboxHwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Call ShowWindow(m_EditboxHwnd, SW_SHOW)

    Call SetFocusAPI(m_EditboxHwnd)
    Call SendMessage(m_EditboxHwnd, EM_SETSEL, 0, SendMessage(m_EditboxHwnd, WM_GETTEXTLENGTH, 0, 0))
End Sub
Private Sub EndEditMode(Cancel As Boolean)
    If (m_EditboxHwnd = 0) Then Exit Sub

    If (m_ActiveObject <> 0) Then
        Dim sCaption As String
        sCaption = GetWindowTextEx(m_EditboxHwnd)

        '//**** Check if there is an event handler ****
        If (isContainerSupportEvents(m_ActiveContainer)) Then

            '//**** If the user cancel this action, exit ****
            Cancel = (GetEventObject(m_ActiveContainer).EventAfterEdit(m_ActiveContainer, m_ActiveObject, sCaption))
        End If

        If (Not Cancel) Then
            Call SetWindowText(m_ActiveObject, sCaption)
        End If
    End If

    m_ActiveObject = 0
    m_ActiveContainer = 0
    Call DestroyWindow(m_EditboxHwnd)
End Sub




'//**************************************************************************
'// Controls functions
'//**************************************************************************
Private Sub UnSubclassAllChild(hWndContainer As Long)
    Call EnumChildWindows(hWndContainer, AddressOf EnumChildProc, EnumMode_UnSubclass)
End Sub
Private Function UnSubclassChild(hWnd As Long) As Boolean
    '//**** Get the address of the previous procedure ****
    Dim prevWndProc As Long
    prevWndProc = GetProp(hWnd, PropName_PrevWndProc)

    '//**** The control was not subclassed ****
    If (prevWndProc = 0) Then Exit Function

    Call SetWindowLong(hWnd, GWL_WNDPROC, prevWndProc)
    UnSubclassChild = True
End Function
Public Sub EnableAllControlDrag(hWndContainer As Long, Enabled As Boolean)
    Call EnumChildWindows(hWndContainer, AddressOf EnumChildProc, IIf(Enabled, EnumMode_EnableDrag, EnumMode_DisableDrag))
End Sub
Public Function EnableControlDrag(hWnd As Long, Enabled As Boolean)
    If (GetClassNameEx(hWnd) = ClassName_GrabBox) Then Exit Function

    '//**** Get the address of the previous procedure ****
    Dim prevWndProc As Long
    prevWndProc = GetProp(hWnd, PropName_PrevWndProc)

    '//**** The control is not subclassed, subclass it ****
    If (prevWndProc = 0) Then

        '//**** Get the current window procedure address ****
        prevWndProc = GetWindowLong(hWnd, GWL_WNDPROC)

        '//**** Store this address ****
        Call SetProp(hWnd, PropName_PrevWndProc, prevWndProc)

        '//**** Set the new window procedure address ****
        Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProcChild)
    End If

    EnableControlDrag = (SetProp(hWnd, PropName_DragEnabled, IIf(Enabled, 1, 0)) <> 0)
End Function




'//**************************************************************************
'// Callbacks
'//**************************************************************************
Private Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Select Case lParam

        Case EnumMode_EnableDrag
            EnumChildProc = EnableControlDrag(hWnd, True)

        Case EnumMode_DisableDrag
            EnumChildProc = EnableControlDrag(hWnd, False)

        Case EnumMode_UnSubclass
            EnumChildProc = UnSubclassChild(hWnd)
    End Select
End Function
Private Function WindowProcContainer(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '//**** Get the address of the previous procedure ****
    Dim prevWndProc As Long
    prevWndProc = GetProp(hWnd, PropName_PrevWndProc)

    '//**** We dont have the address... then call the default procedure ****
    If (prevWndProc = 0) Then
        WindowProcContainer = DefWindowProc(hWnd, uMsg, wParam, lParam)
        Exit Function
    End If

    Select Case uMsg
        Case WM_ERASEBKGND
            '//**** Get the grid brush ****
            Dim hBrush As Long
            hBrush = GetProp(hWnd, PropName_GridBrush)

            Dim ShowGrid As Boolean
            ShowGrid = IIf(GetProp(hWnd, PropName_ShowGrid) <> 0, 1, 0)

            If (hBrush <> 0) And (ShowGrid) Then
                Dim ClientRc As RECT
                Call GetClientRect(hWnd, ClientRc)

                '//**** Validate the update region to keep vb from drawing over. VB is so well made :) ****
                Dim UpdateRc As RECT
                Call GetUpdateRect(hWnd, UpdateRc, True)
                Call ValidateRect(hWnd, UpdateRc)

                '//**** Ask vb wich color to use ****
                Call SendMessage(hWnd, WM_CTLCOLORSTATIC, wParam, ByVal hWnd)

                Dim GridSize As Long
                GridSize = GetProp(hWnd, PropName_GridSize)

                '//**** Set the brush draw origin to (-1,-1) ****
                Dim PrevOrg As POINTAPI
                Call SetBrushOrgEx(wParam, -1, -1, PrevOrg)

                '//**** Swap background and foreground colors ****
                Call SwapBkColors(wParam)

                '//**** Fill the background ****
                Call FillRect(wParam, UpdateRc, hBrush)

                WindowProcContainer = True
                Exit Function
            End If

        Case WM_PAINT
            '//**** Get the grid brush ****
            hBrush = GetProp(hWnd, PropName_GridBrush)

            ShowGrid = IIf(GetProp(hWnd, PropName_ShowGrid) <> 0, 1, 0)

            If (hBrush <> 0) And (ShowGrid) Then
                Dim ps As PAINTSTRUCT
                Call BeginPaint(hWnd, ps)
                Call EndPaint(hWnd, ps)

                Exit Function
            End If

        Case WM_DESTROY
            Call UnSubclassAllChild(hWnd)

        Case WM_LBUTTONDOWN
            Call onButtonDown(hWnd, 1, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam))

        Case WM_LBUTTONUP
            Call EndControlDrag(hWnd)

        Case WM_LBUTTONDBLCLK
            Call onButtonDblClk(hWnd, 1, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam))

        Case WM_MOUSEMOVE
            Call DragMove

        Case WM_SETCURSOR
            If onSetCursor() Then Exit Function

        Case WM_CTLCOLOREDIT
            Dim hdc As Long
            hdc = wParam

            If (lParam = m_EditboxHwnd) Then
                Call SetTextColor(hdc, GetSysColor(COLOR_WINDOWTEXT))

                WindowProcContainer = GetSysColorBrush(COLOR_WINDOW)
                Exit Function
            End If
    End Select

    '//**** Call the previous window procedure ****
    WindowProcContainer = CallWindowProc(prevWndProc, hWnd, uMsg, wParam, lParam)
End Function
Private Function WindowProcChild(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '//**** Get the address of the previous procedure ****
    Dim prevWndProc As Long
    prevWndProc = GetProp(hWnd, PropName_PrevWndProc)

    '//**** We dont have the address... then call the default procedure ****
    If (prevWndProc = 0) Then
        WindowProcChild = DefWindowProc(hWnd, uMsg, wParam, lParam)
        Exit Function
    End If

    Dim DragEnabled As Boolean
    DragEnabled = (GetProp(hWnd, PropName_DragEnabled) <> 0) And (GetProp(GetParent(hWnd), PropName_DragEnabled) <> 0)


    Dim hParent As Long
    hParent = GetParent(hWnd)

    Dim SelectedHwnd As Long
    SelectedHwnd = GetProp(hParent, PropName_SelectedHwnd)

    Select Case uMsg
        Case WM_NCHITTEST

            If (DragEnabled) Then
                WindowProcChild = HTTRANSPARENT
                Exit Function
            End If

        Case WM_MOVE, WM_SIZE
            If (SelectedHwnd = hWnd) Then
                Call ShowGrabBoxes(hParent, hWnd)
            End If

        Case WM_DESTROY
            If (SelectedHwnd = hWnd) Then
                Call SelectControl(hParent, 0)
                Call HideGrabBoxes(hParent)

                Exit Function
            End If
    End Select

    '//**** Call the previous window procedure ****
    WindowProcChild = CallWindowProc(prevWndProc, hWnd, uMsg, wParam, lParam)
End Function
Private Function onButtonDblClk(hWnd As Long, Button As Long, x As Long, y As Long)
    If (m_OnDrag) Then Exit Function

    Dim hwndUnder As Long
    hwndUnder = ChildWindowFromPoint(hWnd, x, y)

    '//**** Cant edit the grab boxes ****
    If (GetClassNameEx(hwndUnder) = ClassName_GrabBox) Then Exit Function

    '//**** Cant edit the container itself ****
    If (hwndUnder = hWnd) Then Exit Function

    Call BeginEditMode(hWnd, hwndUnder)
End Function
Private Function onButtonDown(hWnd As Long, Button As Long, x As Long, y As Long)
    If (m_OnDrag) Then Exit Function

    Dim hwndUnder As Long
    hwndUnder = ChildWindowFromPoint(hWnd, x, y)

    '//**** Cant drag the grab boxes ****
    If (GetClassNameEx(hwndUnder) = ClassName_GrabBox) Then
        Exit Function
    End If

    '//**** Cancel edit mode ****
    Call EndEditMode(False)

    '//**** Cant drag the container ****
    If (hwndUnder = hWnd) Then
        Call SelectControl(hWnd, 0)
        Call HideGrabBoxes(hWnd)

        Exit Function
    End If

    Call SelectControl(hWnd, hwndUnder)
    Call HideGrabBoxes(hWnd)
    DoEvents

    Call BeginControlDrag(hWnd, hwndUnder, DragMode_Move)
End Function
Private Function onSetCursor() As Boolean
    Dim hIcon As Long
    Select Case m_DragMode
        Case 1, 8: hIcon = LoadCursor(ByVal 0&, IDC_SIZENWSE)
        Case 2, 7: hIcon = LoadCursor(ByVal 0&, IDC_SIZENS)
        Case 3, 6: hIcon = LoadCursor(ByVal 0&, IDC_SIZENESW)
        Case 4, 5: hIcon = LoadCursor(ByVal 0&, IDC_SIZEWE)
        Case Else

            Exit Function
    End Select



    Call SetCursor(hIcon)
    onSetCursor = True
End Function








'//**************************************************************************
'// Drag functions
'//**************************************************************************
Private Function BeginControlDrag(hWndContainer As Long, hWnd As Long, DragMode As Long) As Boolean
    If (m_OnDrag) Then Exit Function


    '//**** Check if there is an event handler ****
    If (isContainerSupportEvents(hWndContainer)) Then
        Dim Cancel As Boolean
        Cancel = GetEventObject(hWndContainer).EventBeginDrag(hWndContainer, hWnd)

        '//**** If the user cancel this action, restore the grab boxes and exit ****
        If (Cancel) Then
            Call ShowGrabBoxes(hWndContainer, hWnd)
            Exit Function
        End If
    End If


    m_OnDrag = True
    m_ActiveContainer = hWndContainer
    m_ActiveObject = hWnd

    Call SetCapture(hWndContainer)

    '//**** Get the handle to screen dc ****
    m_hdcScreen = GetDC(ByVal 0&)

    '//**** Set mix mode to invert ****
    Call SetROP2(m_hdcScreen, R2_NOTXORPEN)

    '//**** Create the pen used to draw around the selection *****
    m_hDragPen = CreatePen(PS_SOLID, 3, vbBlack)
    m_hOldPen = SelectObject(m_hdcScreen, m_hDragPen)

    '//**** Get the rect of the control to be dragged ****
    Call GetWindowRect(hWnd, m_DragRc)
    Let m_SnapRc = m_DragRc

    '//**** Get the current position ****
    Call GetCursorPos(DragOriginPt)

    m_DragMode = DragMode
    Call onSetCursor

    m_DrawStatus = 0
    Call DrawDragRect(True, m_SnapRc)
    DoEvents
End Function
Private Sub EndControlDrag(hWndContainer As Long)
    If (Not m_OnDrag) Then Exit Sub

    '//**** Erase the drag rectangle ****
    Call DrawDragRect(False, m_SnapRc)

    '//**** Delete the drag pen ****
    Call SelectObject(m_hdcScreen, m_hOldPen)
    Call DeleteObject(m_hDragPen)


    '//**** Restore the mix mode to default ****
    Call SetROP2(m_hdcScreen, R2_COPYPEN)

    '//**** Release the screen dc ****
    Call ReleaseDC(0, m_hdcScreen)

    '//**** Release mouse capture ****
    Call ReleaseCapture
    m_OnDrag = False

    '//**** Normalize the rectangle ****
    Let m_DragRc = m_SnapRc
    Call NormalizeRect(m_DragRc)

    '//**** Get the hwnd of the selected control ****
    Dim hWnd As Long
    hWnd = GetProp(hWndContainer, PropName_SelectedHwnd)
    If (hWnd <> 0) Then

        '//**** Check if there is an event handler ****
        If (isContainerSupportEvents(hWndContainer)) Then

            Dim Width As Long, Height As Long
            Width = (m_DragRc.Right - m_DragRc.Left)
            Height = (m_DragRc.Bottom - m_DragRc.Top)

            Dim Cancel As Boolean
            Cancel = GetEventObject(hWndContainer).EventStopDrag(hWndContainer, hWnd, m_DragRc.Left, m_DragRc.Top, Width, Height)

            m_DragRc.Right = (m_DragRc.Left + Width)
            m_DragRc.Bottom = (m_DragRc.Top + Height)
        End If

        Dim NewContainer As Long
        NewContainer = hWndContainer
        If (Not Cancel) Then
            '//**** Get the window rect of the container ****
            Dim WindowRc As RECT
            Call GetWindowRect(NewContainer, WindowRc)

            '//**** check if the cursor is into this rectangle ****
            Dim CurPos As POINTAPI
            Call GetCursorPos(CurPos)
            If ((PointInRect(CurPos, WindowRc) = 0) And (m_DragMode = DragMode_Move)) Then

                '//**** Find a container that accept drag & drop ****
                NewContainer = WindowFromPoint(CurPos.x, CurPos.y)

                Do
                    If (GetProp(NewContainer, PropName_AcceptDragDrop) = 1) Then Exit Do
                    NewContainer = GetParent(NewContainer)
                Loop Until (NewContainer = 0)
            Else

                '//**** Keep the same container ****
                NewContainer = hWndContainer
            End If

            '//**** Send an object drop message to the destination container ****
            If (NewContainer <> hWndContainer) And (NewContainer <> 0) Then
                If (isContainerSupportEvents(NewContainer)) Then
                    If (GetEventObject(NewContainer).ObjectDrop(NewContainer, hWndContainer, hWnd)) Then
                        NewContainer = 0
                    End If
                End If
            End If

            If (NewContainer <> 0) Then
                Call LockWindowUpdate(NewContainer)

                '//**** Set the new container ****
                If (NewContainer <> hWndContainer) Then
                    Call SetParent(hWnd, NewContainer)
                End If

                '//**** Set the new controls position ****
                Call ScreenRectToClient(NewContainer, m_DragRc)
                Call SetWindowPos(hWnd, 0, m_DragRc.Left, m_DragRc.Top, (m_DragRc.Right - m_DragRc.Left), (m_DragRc.Bottom - m_DragRc.Top), SWP_NOZORDER Or SWP_NOACTIVATE)

                If ((m_EditboxHwnd <> 0) And (m_ActiveObject = hWnd)) Then
                    Call SetWindowPos(m_EditboxHwnd, 0, m_DragRc.Left, m_DragRc.Top, (m_DragRc.Right - m_DragRc.Left), (m_DragRc.Bottom - m_DragRc.Top), SWP_NOZORDER Or SWP_NOACTIVATE)
                End If

                Call LockWindowUpdate(ByVal 0&)
            End If
        End If

        If (NewContainer = 0) Then NewContainer = hWndContainer

        '//**** Show the grab boxes around the control ****
        Call SelectControl(NewContainer, hWnd)
        Call ShowGrabBoxes(NewContainer, hWnd)
    End If

    m_ActiveContainer = 0
    m_ActiveObject = 0
    m_DragMode = 0
    Call onSetCursor

End Sub
Private Sub DragMove()
    If (Not m_OnDrag) Then Exit Sub

    '//**** Get the current cursor position ****
    Dim NewOriginPt As POINTAPI
    Call GetCursorPos(NewOriginPt)

    '//**** Get the window handle under the cursor ****
    Dim hwndUnder As Long
    hwndUnder = WindowFromPoint(NewOriginPt.x, NewOriginPt.y)
    If (hwndUnder = m_EditboxHwnd) Then hwndUnder = GetParent(m_EditboxHwnd)

    '//**** Check if this window is a valid container ****
    Dim IsValid As Boolean
    If (hwndUnder = m_ActiveContainer) Then
        IsValid = True
    Else
        IsValid = GetProp(hwndUnder, PropName_AcceptDragDrop)
    End If

    If (IsValid) Then
        Dim GridSize As Long, SnapToGrid As Boolean
        GridSize = GetProp(hwndUnder, PropName_GridSize)
        SnapToGrid = IIf(GetProp(hwndUnder, PropName_SnapToGrid) <> 0, True, False)

        '//**** Get the client position ****
        Dim ClientPT As POINTAPI
        Let ClientPT = NewOriginPt
        Call ScreenToClient(hwndUnder, ClientPT)
    End If

    '//**** Get the diference beetween the old and new cursor position ****
    Dim OffsetX As Long, OffSetY As Long
    OffsetX = (NewOriginPt.x - DragOriginPt.x)
    OffSetY = (NewOriginPt.y - DragOriginPt.y)
    Let DragOriginPt = NewOriginPt

    '//**** Move the drag rect ****
    Select Case m_DragMode
        Case DragMode_Move
            Call OffsetRect(m_DragRc, OffsetX, OffSetY)

        Case DragMode_SizeNW
            m_DragRc.Left = m_DragRc.Left + OffsetX
            m_DragRc.Top = m_DragRc.Top + OffSetY

        Case DragMode_SizeN
            m_DragRc.Top = m_DragRc.Top + OffSetY

        Case DragMode_SizeNE
            m_DragRc.Right = m_DragRc.Right + OffsetX
            m_DragRc.Top = m_DragRc.Top + OffSetY

        Case DragMode_SizeW
            m_DragRc.Left = m_DragRc.Left + OffsetX

        Case DragMode_SizeE
            m_DragRc.Right = m_DragRc.Right + OffsetX

        Case DragMode_SizeSW
            m_DragRc.Left = m_DragRc.Left + OffsetX
            m_DragRc.Bottom = m_DragRc.Bottom + OffSetY

        Case DragMode_SizeS
            m_DragRc.Bottom = m_DragRc.Bottom + OffSetY

        Case DragMode_SizeSE
            m_DragRc.Right = m_DragRc.Right + OffsetX
            m_DragRc.Bottom = m_DragRc.Bottom + OffSetY
    End Select


    Dim OldSnapRc As RECT
    Let OldSnapRc = m_SnapRc

    If ((Not IsValid) And (m_DragMode = DragMode_Move)) Then
        m_InvalidMove = False
    Else
        m_InvalidMove = True
    End If
    Call onSetCursor

    If ((IsValid) And (SnapToGrid) And (GridSize > 2)) Then
        '//**** Convert the drag rect to client rect ****
        Dim DragRc As RECT
        Let DragRc = m_DragRc
        Call ScreenRectToClient(hwndUnder, DragRc)

        '//**** Get the nearest snap points ****
        DragRc.Left = Round((DragRc.Left / GridSize), 0) * GridSize - 1
        DragRc.Top = Round((DragRc.Top / GridSize), 0) * GridSize - 1
        DragRc.Right = Round((DragRc.Right / GridSize), 0) * GridSize
        DragRc.Bottom = Round((DragRc.Bottom / GridSize), 0) * GridSize

        '//**** Convert the rectangle back to screen rect ****
        Call ClientRectToScreen(hwndUnder, DragRc)

        '//**** Apply the new position ****
        Select Case m_DragMode
            Case DragMode_Move
                Call OffsetRect(m_SnapRc, (DragRc.Left - m_SnapRc.Left), (DragRc.Top - m_SnapRc.Top))

            Case DragMode_SizeE
                m_SnapRc.Right = DragRc.Right

            Case DragMode_SizeSE
                m_SnapRc.Right = DragRc.Right
                m_SnapRc.Bottom = DragRc.Bottom

            Case DragMode_SizeS
                m_SnapRc.Bottom = DragRc.Bottom

            Case DragMode_SizeSW
                m_SnapRc.Left = DragRc.Left
                m_SnapRc.Bottom = DragRc.Bottom

            Case DragMode_SizeW
                m_SnapRc.Left = DragRc.Left

            Case DragMode_SizeNW
                m_SnapRc.Left = DragRc.Left
                m_SnapRc.Top = DragRc.Top

            Case DragMode_SizeN
                m_SnapRc.Top = DragRc.Top

            Case DragMode_SizeNE
                m_SnapRc.Top = DragRc.Top
                m_SnapRc.Right = DragRc.Right
        End Select
    Else
        Let m_SnapRc = m_DragRc
    End If

    '//**** Check if there is an event handler ****
    If (isContainerSupportEvents(m_ActiveContainer)) Then
        Dim Width As Long, Height As Long
        Width = (m_DragRc.Right - m_DragRc.Left)
        Height = (m_DragRc.Bottom - m_DragRc.Top)

        If (m_DragMode = DragMode_Move) Then
            Call GetEventObject(m_ActiveContainer).EventDragMove(m_ActiveContainer, m_ActiveObject, m_DragRc.Left, m_DragRc.Top, Width, Height)

        Else
            Call GetEventObject(m_ActiveContainer).EventDragResize(m_ActiveContainer, m_ActiveObject, m_DragRc.Left, m_DragRc.Top, Width, Height, m_DragMode)
        End If

        m_DragRc.Right = (m_DragRc.Left + Width)
        m_DragRc.Bottom = (m_DragRc.Top + Height)
    End If

    '//**** If no change was made, exit ****
    If (EqualRect(m_SnapRc, OldSnapRc) <> 0) Then
        Exit Sub
    End If

    '//**** Undraw the drag rect ****
    Call DrawDragRect(False, OldSnapRc)

    '//**** Draw the new drag rect ****
    Call DrawDragRect(True, m_SnapRc)
End Sub
Private Sub DrawDragRect(Draw As Boolean, lpRect As RECT)
    If ((Draw) And (m_DrawStatus <> 0)) Then Exit Sub
    If ((Not Draw) And (m_DrawStatus = 0)) Then Exit Sub



    Call Rectangle(m_hdcScreen, lpRect.Left, lpRect.Top, lpRect.Right, lpRect.Bottom)
    m_DrawStatus = (Not m_DrawStatus)
End Sub



'//**************************************************************************
'// Grab box functions
'//**************************************************************************
Private Sub RegisterGrabBoxes()
    If (m_GrabBoxInit) Then Exit Sub

    Dim Wc As WNDCLASS
    Wc.lpszClassName = ClassName_GrabBox
    Wc.hInstance = App.hInstance
    Wc.lpfnwndproc = GetAddress(AddressOf WindowProcGrab)

    '//**** Register the class ****
    m_GrabBoxInit = (RegisterClass(Wc) <> 0)
End Sub
Private Sub UnRegisterGrabBoxes()
    If (Not m_GrabBoxInit) Then Exit Sub

    '//**** Unregister the class ****
    Call UnregisterClass(ClassName_GrabBox, App.hInstance)
    m_GrabBoxInit = False
End Sub
Private Function CreateGrabBoxes(hWndContainer As Long) As Boolean
    Dim i As Long
    For i = 1 To 8
        Dim hWnd As Long
        hWnd = CreateWindowEx(0, ClassName_GrabBox, "", WS_CHILD, 0, 0, Metrics_GrabBoxWidth, Metrics_GrabBoxWidth, hWndContainer, 0, 0, 0)

        Call SetProp(hWndContainer, PropName_HwndGrab & i, hWnd)
        Call SetProp(hWnd, PropName_GrabBoxID, i)
    Next

    CreateGrabBoxes = True
End Function
Private Function DestroyGrabBoxes(hWndContainer As Long) As Boolean

    Dim i As Long
    For i = 1 To 8
        Dim hWnd As Long
        hWnd = GetProp(hWndContainer, PropName_HwndGrab & i)
        Call RemoveProp(hWndContainer, PropName_HwndGrab & i)

        Call DestroyWindow(hWnd)
    Next

    DestroyGrabBoxes = True
End Function

'//--- Call back ---
Private Function WindowProcGrab(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim id As Long
    id = GetProp(hWnd, PropName_GrabBoxID)

    Select Case uMsg

        Case WM_ERASEBKGND
            Dim hdc As Long
            hdc = wParam


            Dim hBrush As Long
            hBrush = GetSysColorBrush(COLOR_HIGHLIGHT)
            Call SelectObject(hdc, hBrush)

            Dim hPen As Long, hOldPen As Long
            hPen = CreatePen(PS_SOLID, 0, GetSysColor(COLOR_HIGHLIGHTTEXT))
            hOldPen = SelectObject(hdc, hPen)

            Dim ClientRc As RECT
            Call GetClientRect(hWnd, ClientRc)
            Call Rectangle(hdc, 0, 0, ClientRc.Right, ClientRc.Bottom)

            Call SelectObject(hdc, hOldPen)
            Call DeleteObject(hPen)

        Case WM_SETCURSOR
            Dim hIcon As Long
            Select Case id
                Case 1, 8: hIcon = LoadCursor(ByVal 0&, IDC_SIZENWSE)
                Case 2, 7: hIcon = LoadCursor(ByVal 0&, IDC_SIZENS)
                Case 3, 6: hIcon = LoadCursor(ByVal 0&, IDC_SIZENESW)
                Case 4, 5: hIcon = LoadCursor(ByVal 0&, IDC_SIZEWE)
            End Select

            Call SetCursor(hIcon)
            Exit Function


        Case WM_LBUTTONDOWN
            Dim hWndSelected As Long
            hWndSelected = GetProp(GetParent(hWnd), PropName_SelectedHwnd)

            If (hWndSelected <> 0) Then
                Call HideGrabBoxes(GetParent(hWnd))
                DoEvents

                Call BeginControlDrag(GetParent(hWnd), hWndSelected, id)
            End If
    End Select
    WindowProcGrab = DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function
Private Sub HideGrabBoxes(hWndContainer As Long)
    Dim i As Long
    For i = 1 To 8
        Dim hWnd As Long
        hWnd = GetProp(hWndContainer, PropName_HwndGrab & i)

        Call ShowWindow(hWnd, SW_HIDE)
    Next
    DoEvents
End Sub
Private Sub ShowGrabBoxes(hWndContainer As Long, hWnd As Long)

    '//**** Hide all boxes ***
    Dim hwndGrab(8) As Long
    Dim i As Long
    For i = 1 To 8
        hwndGrab(i) = GetProp(hWndContainer, PropName_HwndGrab & i)
        Call ShowWindow(hwndGrab(i), SW_HIDE)
    Next


    If (GetProp(hWndContainer, PropName_DragEnabled) = 0) Then Exit Sub


    '//**** Get the control rect and convert it to client related position ****
    Dim WindowRc As RECT
    Call GetWindowRect(hWnd, WindowRc)
    Call ScreenRectToClient(hWndContainer, WindowRc)

    '//**** Move all grab boxes ****
    Call SetWindowPos(hwndGrab(1), HWND_TOP, WindowRc.Left - Metrics_GrabBoxWidth, WindowRc.Top - Metrics_GrabBoxWidth, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE)
    Call SetWindowPos(hwndGrab(2), HWND_TOP, WindowRc.Left + Int((WindowRc.Right - WindowRc.Left) / 2) - Int(Metrics_GrabBoxWidth / 2), WindowRc.Top - Metrics_GrabBoxWidth, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE)
    Call SetWindowPos(hwndGrab(3), HWND_TOP, WindowRc.Right, WindowRc.Top - Metrics_GrabBoxWidth, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE)
    Call SetWindowPos(hwndGrab(4), HWND_TOP, (WindowRc.Left - Metrics_GrabBoxWidth), WindowRc.Top + Int((WindowRc.Bottom - WindowRc.Top) / 2) - Int(Metrics_GrabBoxWidth / 2), 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE)
    Call SetWindowPos(hwndGrab(5), HWND_TOP, WindowRc.Right, WindowRc.Top + Int((WindowRc.Bottom - WindowRc.Top) / 2) - Int(Metrics_GrabBoxWidth / 2), 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE)
    Call SetWindowPos(hwndGrab(6), HWND_TOP, WindowRc.Left - Metrics_GrabBoxWidth, WindowRc.Bottom, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE)
    Call SetWindowPos(hwndGrab(7), HWND_TOP, WindowRc.Left + Int((WindowRc.Right - WindowRc.Left) / 2) - Int(Metrics_GrabBoxWidth / 2), WindowRc.Bottom, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE)
    Call SetWindowPos(hwndGrab(8), HWND_TOP, WindowRc.Right, WindowRc.Bottom, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE)

    For i = 1 To 8
        Call ShowWindow(hwndGrab(i), SW_SHOW)
    Next
    DoEvents
End Sub
Private Sub SelectControl(hWndContainer As Long, hWnd As Long)
    Call SetProp(hWndContainer, PropName_SelectedHwnd, hWnd)
End Sub



'//**************************************************************************
'// Drawing functions
'//**************************************************************************
Private Function CreateGridBrush(Size As Long) As Long
    Dim nBytes As Long
    nBytes = Int((Size * Size))

    '//**** Define pattern bits ****
    Dim bits() As Integer: ReDim bits(1 To nBytes)
    bits(1) = &H80    '//&H80 = 128 = [1000 0000 0000 0000]

    '//**** Create the pattern bitmap ****
    Dim hBmp As Long
    hBmp = CreateBitmap(Size, Size, 1, 1, bits(1))
    If (hBmp = 0) Then Exit Function

    '//**** Create a brush from the bitmap ****
    CreateGridBrush = CreatePatternBrush(hBmp)
End Function
Private Function SwapBkColors(hdc As Long)
    Dim TempBkColor As Long
    TempBkColor = GetBkColor(hdc)

    Call SetBkColor(hdc, GetTextColor(hdc))
    Call SetTextColor(hdc, TempBkColor)
End Function
