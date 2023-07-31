VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAsigacionRecursos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Recursos para OT Nº"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   Icon            =   "frmAsigacionRecursos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8055
   Begin XtremeSuiteControls.GroupBox grpPersonal 
      Height          =   2865
      Left            =   3720
      TabIndex        =   2
      Top             =   3435
      Width           =   4170
      _Version        =   786432
      _ExtentX        =   7355
      _ExtentY        =   5054
      _StockProps     =   79
      Caption         =   "Personal designado"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ListBox lstPersonal 
         Height          =   2400
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   3810
         _Version        =   786432
         _ExtentX        =   6720
         _ExtentY        =   4233
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   1
      End
   End
   Begin XtremeSuiteControls.GroupBox grpPiezas 
      Height          =   6210
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   3435
      _Version        =   786432
      _ExtentX        =   6059
      _ExtentY        =   10954
      _StockProps     =   79
      Caption         =   "Piezas Seleccionadas de OT"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX gridPiezas 
         Height          =   5790
         Left            =   150
         TabIndex        =   1
         Top             =   270
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   10213
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         GroupByBoxVisible=   0   'False
         ColumnHeaders   =   0   'False
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   1
         Column(1)       =   "frmAsigacionRecursos.frx":000C
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAsigacionRecursos.frx":0124
         FormatStyle(2)  =   "frmAsigacionRecursos.frx":025C
         FormatStyle(3)  =   "frmAsigacionRecursos.frx":030C
         FormatStyle(4)  =   "frmAsigacionRecursos.frx":03C0
         FormatStyle(5)  =   "frmAsigacionRecursos.frx":0498
         FormatStyle(6)  =   "frmAsigacionRecursos.frx":0550
         ImageCount      =   0
         PrinterProperties=   "frmAsigacionRecursos.frx":0630
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   3270
      Left            =   3720
      TabIndex        =   3
      Top             =   90
      Width           =   4170
      _Version        =   786432
      _ExtentX        =   7355
      _ExtentY        =   5768
      _StockProps     =   79
      Caption         =   "Tareas de piezas"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX gridTareas 
         Height          =   2790
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   4921
         Version         =   "2.0"
         HoldSortSettings=   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         GroupByBoxVisible=   0   'False
         ColumnHeaders   =   0   'False
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   1
         Column(1)       =   "frmAsigacionRecursos.frx":0808
         SortKeysCount   =   1
         SortKey(1)      =   "frmAsigacionRecursos.frx":0918
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAsigacionRecursos.frx":0980
         FormatStyle(2)  =   "frmAsigacionRecursos.frx":0AB8
         FormatStyle(3)  =   "frmAsigacionRecursos.frx":0B68
         FormatStyle(4)  =   "frmAsigacionRecursos.frx":0C1C
         FormatStyle(5)  =   "frmAsigacionRecursos.frx":0CF4
         FormatStyle(6)  =   "frmAsigacionRecursos.frx":0DAC
         ImageCount      =   0
         PrinterProperties=   "frmAsigacionRecursos.frx":0E8C
      End
   End
End
Attribute VB_Name = "frmAsigacionRecursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_col As Collection
Dim deta As DetalleOrdenTrabajo
Dim detadto As DetalleOTConjuntoDTO
Private tareas As Collection
Private Tarea As clsTarea
Private empl As clsEmpleado
Private ProcesosId As Dictionary

Private empleados As Collection

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.gridPiezas
    GridEXHelper.CustomizeGrid Me.gridTareas
    Me.gridPiezas.ItemCount = 0
    Me.gridTareas.ItemCount = 0
End Sub

Public Sub llenar(idOt As Long, col As Collection)
    Me.caption = "Asignación de Recursos para OT Nº " & idOt

    Set m_col = col
    Me.gridPiezas.ItemCount = 0
    Me.gridPiezas.ItemCount = m_col.count

    Dim detaVariant As Variant

    Set tareas = New Collection

    Dim TiempoProceso As PlaneamientoTiempoProceso
    Dim tiemposProcesos As New Collection

    Set ProcesosId = New Dictionary

    For Each detaVariant In m_col
        Set deta = Nothing
        Set detadto = Nothing
        If TypeName(detaVariant) = "DetalleOrdenTrabajo" Then
            Set deta = detaVariant
            'Set tiemposProcesos = DAOTiemposProceso.FindAllByDetallePedidoIdAndPiezaId(deta.Id, deta.pieza.Id)
            Set tiemposProcesos = DAOTiemposProceso.FindAllByDetallePedidoId(deta.Id)
        Else
            Set detadto = detaVariant
            'Set tiemposProcesos = DAOTiemposProceso.FindAllByDetallePedidoIdAndPiezaId(detadto.Id, detadto.pieza.Id)
            Set tiemposProcesos = DAOTiemposProceso.FindAllByDetallePedidoId(0, detadto.Id)
        End If

        For Each TiempoProceso In tiemposProcesos
            ProcesosId.Add CStr(TiempoProceso.Id), TiempoProceso.Tarea.Id
            If Not funciones.BuscarEnColeccion(tareas, CStr(TiempoProceso.Tarea.Id)) Then
                tareas.Add TiempoProceso.Tarea, CStr(TiempoProceso.Tarea.Id)
            End If
        Next TiempoProceso

    Next detaVariant

    Me.gridTareas.ItemCount = 0
    Me.gridTareas.ItemCount = tareas.count

End Sub


Private Sub gridPiezas_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    If rowIndex > 0 And m_col.count > 0 Then
        Set deta = Nothing
        Set detadto = Nothing
        If TypeName(m_col.item(rowIndex)) = "DetalleOrdenTrabajo" Then
            Set deta = m_col.item(rowIndex)
            Values(1) = deta.Pieza.nombre
        Else
            Set detadto = m_col.item(rowIndex)
            Values(1) = detadto.Pieza.nombre
        End If
    End If
End Sub

Private Sub gridTareas_SelectionChange()
    If Me.gridTareas.rowIndex(Me.gridTareas.row) > 0 Then
        Set Tarea = tareas.item(Me.gridTareas.rowIndex(Me.gridTareas.row))
        Me.lstPersonal.Clear
        Set empleados = DAOEmpleados.GetEmpleadosByTareaId(Tarea.Id)
        For Each empl In empleados
            Me.lstPersonal.AddItem empl.NombreCompleto & " (" & empl.legajo & ")"
            Me.lstPersonal.ItemData(Me.lstPersonal.NewIndex) = empl.Id
            Me.lstPersonal.Checked(Me.lstPersonal.NewIndex) = FindAllAsignedByEmpleadoAndTiempoProcesoAndTareaId(empl.Id, ProcesosId, Tarea.Id).count > 0
        Next empl
    End If
End Sub

Private Sub gridTareas_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And tareas.count > 0 Then
        Set Tarea = tareas.item(rowIndex)
        Values(1) = Tarea.Description
    End If
End Sub

Private Sub lstPersonal_ItemCheck(ByVal item As Long)
    On Error GoTo E
    Set Tarea = tareas.item(Me.gridTareas.rowIndex(Me.gridTareas.row))

    conectar.BeginTransaction

    'delete antes por si las moscas
    Dim q As String
    q = "DELETE FROM PlaneamientoTiemposProcesosDetalle WHERE idTiemposProcesos IN (" & funciones.JoinDictionaryKeyValues(ProcesosId, ", ") & ") AND inico = '0000-00-00 00:00:00' AND legajo = " & empleados.item(CStr(Me.lstPersonal.ItemData(item))).Id
    If Not conectar.execute(q) Then GoTo E

    If Me.lstPersonal.Checked(item) Then
        'agrego
        Dim procId As Variant
        Dim tpd As PlaneamientoTiempoProcesoDetalle
        For Each procId In ProcesosId
            If ProcesosId.item(procId) = Tarea.Id Then
                Set tpd = New PlaneamientoTiempoProcesoDetalle
                tpd.IdPlaneamientoTiempoProceso = procId
                Set tpd.Empleado = empleados.item(CStr(Me.lstPersonal.ItemData(item)))
                If Not DAOTiemposProcesosDetalles.Save(tpd, True) Then GoTo E
            End If
        Next
    Else
        'borro, delete ya hecho
    End If

    conectar.CommitTransaction

    Exit Sub
E:
    conectar.RollBackTransaction
    MsgBox "Ocurrió un error.", vbCritical + vbOKOnly
End Sub


