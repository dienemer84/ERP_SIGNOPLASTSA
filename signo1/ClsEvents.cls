VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//**************************************************************************
'// ----------------- Module -----------------
'// Name        : --
'// Version     : --
'// Author      : Benoit Frigon
'// Created on  : 13-MAY-2002
'// Last update : 14-MAY-2002
'// File        : ClsEvents.cls
'// Desc.       : Event handler
'//**************************************************************************
'// All rights reserved@Logiciels M.T.L enr. NEQ# 22-48153829(Québec)
'//**************************************************************************
Option Explicit



'//**************************************************************************
'// Events
'//**************************************************************************
Public Event BeginDrag(hWndContainer As Long, hWndObject As Long, Cancel As Boolean)
Public Event StopDrag(hWndContainer As Long, hWndObject As Long, x As Long, y As Long, Width As Long, Height As Long, Cancel As Boolean)
Public Event DragMove(hWndContainer As Long, hWndObject As Long, x As Long, y As Long, Width As Long, Height As Long)
Public Event DragResize(hWndContainer As Long, hWndObject As Long, x As Long, y As Long, Width As Long, Height As Long, SizeMode As Long)
Public Event BeforeEdit(hWndContainer As Long, hWndObject As Long, Cancel As Boolean)
Public Event AfterEdit(hWndContainer As Long, hWndObject As Long, WindowText As String, Cancel As Boolean)
Public Event ObjectDrop(hWndContainerSource As Long, hWndContainer As Long, hWndObject As Long, Reject As Boolean)



'//**************************************************************************
'// Events raiser (We cant raise an event from outside an object)
'//**************************************************************************
Public Function EventBeginDrag(hWndContainer As Long, hWndObject As Long) As Boolean
    Dim l_HwndContainer As Long
    Dim l_HwndObject As Long
    l_HwndContainer = hWndContainer
    l_HwndObject = hWndObject

    RaiseEvent BeginDrag(l_HwndContainer, l_HwndObject, EventBeginDrag)
End Function
Public Function EventStopDrag(hWndContainer As Long, hWndObject As Long, x As Long, y As Long, Width As Long, Height As Long) As Boolean
    Dim l_HwndContainer As Long
    Dim l_HwndObject As Long
    l_HwndContainer = hWndContainer
    l_HwndObject = hWndObject

    RaiseEvent StopDrag(l_HwndContainer, l_HwndObject, x, y, Width, Height, EventStopDrag)
End Function
Public Sub EventDragMove(hWndContainer As Long, hWndObject As Long, x As Long, y As Long, Width As Long, Height As Long)
    Dim l_HwndContainer As Long
    Dim l_HwndObject As Long
    l_HwndContainer = hWndContainer
    l_HwndObject = hWndObject

    RaiseEvent DragMove(l_HwndContainer, l_HwndObject, x, y, Width, Height)
End Sub
Public Sub EventDragResize(hWndContainer As Long, hWndObject As Long, x As Long, y As Long, Width As Long, Height As Long, SizeMode As Long)
    Dim l_HwndContainer As Long
    Dim l_HwndObject As Long
    l_HwndContainer = hWndContainer
    l_HwndObject = hWndObject

    RaiseEvent DragResize(l_HwndContainer, l_HwndObject, x, y, Width, Height, SizeMode)
End Sub
Public Function EventBeforeEdit(hWndContainer As Long, hWndObject As Long) As Boolean
    Dim l_HwndContainer As Long
    Dim l_HwndObject As Long
    l_HwndContainer = hWndContainer
    l_HwndObject = hWndObject

    RaiseEvent BeforeEdit(l_HwndContainer, l_HwndObject, EventBeforeEdit)
End Function
Public Function EventAfterEdit(hWndContainer As Long, hWndObject As Long, WindowText As String) As Boolean
    Dim l_HwndContainer As Long
    Dim l_HwndObject As Long
    l_HwndContainer = hWndContainer
    l_HwndObject = hWndObject

    RaiseEvent AfterEdit(l_HwndContainer, l_HwndObject, WindowText, EventAfterEdit)
End Function
Public Function ObjectDrop(hWndContainerSource As Long, hWndContainer As Long, hWndObject As Long) As Boolean
    Dim l_HwndContainer As Long
    Dim l_HwndContainerSource As Long
    Dim l_HwndObject As Long
    l_HwndContainer = hWndContainer
    l_HwndObject = hWndObject
    l_HwndContainerSource = hWndContainerSource

    RaiseEvent ObjectDrop(l_HwndContainerSource, l_HwndContainer, l_HwndObject, ObjectDrop)
End Function
