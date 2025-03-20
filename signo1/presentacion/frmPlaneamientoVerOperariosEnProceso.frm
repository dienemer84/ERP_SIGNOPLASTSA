VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoVerOperariosEnProceso 
   Caption         =   "Empleados en Proceso"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   Icon            =   "frmPlaneamientoVerOperariosEnProceso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   11100
   Begin VB.Timer Timer1 
      Interval        =   15000
      Left            =   10590
      Top             =   4680
   End
   Begin XtremeSuiteControls.PushButton cmdActualizar 
      Height          =   495
      Left            =   45
      TabIndex        =   1
      Top             =   4635
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "F5 - Actualizar Lista"
      UseVisualStyle  =   -1  'True
   End
   Begin GridEX20.GridEX grilla 
      Height          =   4560
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   8043
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   5
      Column(1)       =   "frmPlaneamientoVerOperariosEnProceso.frx":000C
      Column(2)       =   "frmPlaneamientoVerOperariosEnProceso.frx":012C
      Column(3)       =   "frmPlaneamientoVerOperariosEnProceso.frx":0218
      Column(4)       =   "frmPlaneamientoVerOperariosEnProceso.frx":0308
      Column(5)       =   "frmPlaneamientoVerOperariosEnProceso.frx":0424
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmPlaneamientoVerOperariosEnProceso.frx":0574
      FormatStyle(2)  =   "frmPlaneamientoVerOperariosEnProceso.frx":06AC
      FormatStyle(3)  =   "frmPlaneamientoVerOperariosEnProceso.frx":075C
      FormatStyle(4)  =   "frmPlaneamientoVerOperariosEnProceso.frx":0810
      FormatStyle(5)  =   "frmPlaneamientoVerOperariosEnProceso.frx":08E8
      FormatStyle(6)  =   "frmPlaneamientoVerOperariosEnProceso.frx":09A0
      ImageCount      =   0
      PrinterProperties=   "frmPlaneamientoVerOperariosEnProceso.frx":0A80
   End
End
Attribute VB_Name = "frmPlaneamientoVerOperariosEnProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim operarios As New Collection
Dim ptp As PlaneamientoTiempoProcesoDetalle



Private Sub cmdActualizar_Click()
    llenarGrilla
End Sub

Private Sub Form_Activate()
    llenarGrilla
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KeyCodeConstants.vbKeyF5 Then
        llenarGrilla
    End If
End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.grilla, True, False
    llenarGrilla
End Sub

Private Sub llenarGrilla()
    Set operarios = DAOTiemposProcesosDetalles.FindAllWithoutFinish()
    Me.grilla.ItemCount = 0
    Me.grilla.ItemCount = operarios.count
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.grilla.Width = Me.ScaleWidth
    Me.grilla.Height = Me.ScaleHeight - 600
    Me.cmdActualizar.Top = Me.ScaleHeight - Me.cmdActualizar.Height - 60
End Sub

Private Sub grilla_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    ColumnHeaderClick Me.grilla, Column
End Sub

Private Sub grilla_DblClick()
    If Me.grilla.row > 0 Then
        If Me.grilla.rowIndex(Me.grilla.row) > 0 Then
            Dim ptp As PlaneamientoTiempoProcesoDetalle
            Set ptp = operarios.item(Me.grilla.rowIndex(Me.grilla.row))
            Dim F As New frmPlaneamientoSeguimientoRutas3
            F.txtOTNro.Text = ptp.PlaneamientoTiempoProceso.idpedido
            F.cmdBuscar_Click
            F.Show
            Dim found As Boolean
            Explore F.ReportControl.rows, ptp, F, found
        End If
    End If
End Sub

Private Sub Explore(rows As ReportRows, ptp As PlaneamientoTiempoProcesoDetalle, F As frmPlaneamientoSeguimientoRutas3, found As Boolean)
    If found Then GoTo E

    Dim row As ReportRow
    Dim row2 As ReportRow
    Dim row3 As ReportRow

    For Each row In rows
        If row.record.Tag > 0 Then
            'Debug.Print row.record.Tag, ptp.PlaneamientoTiempoProceso.idDetallePedido, ptp.PlaneamientoTiempoProceso.idDetallePedidoConj
            If (row.record.Tag = ptp.PlaneamientoTiempoProceso.idDetallePedido And ptp.PlaneamientoTiempoProceso.idDetallePedidoConj = 0) Or _
               (row.record.Tag = ptp.PlaneamientoTiempoProceso.idDetallePedidoConj) _
               Then
                found = True

                ExpandParents row
                row.EnsureVisible


                For Each row2 In row.Childs
                    If (row2.record.Tag * -1) = ptp.IdPlaneamientoTiempoProceso Then

                        F.ReportControl.SelectedRows.DeleteAll
                        F.ReportControl.SelectedRows.Add row2
                        F.ReportControl_SelectionChanged
                        row2.EnsureVisible

                        F.ReportControlDetalles.SelectedRows.DeleteAll
                        For Each row3 In F.ReportControlDetalles.rows
                            If row3.record.Tag = ptp.Id Then
                                F.ReportControlDetalles.SelectedRows.Add row3
                                row3.EnsureVisible
                                Exit For
                            End If
                        Next row3
                        Exit For
                    End If
                Next row2
                Exit For

                Exit Sub

            Else
                If row.Childs.count > 0 Then
                    Explore row.Childs, ptp, F, found
                End If
            End If
        End If
    Next row

E:
End Sub

Private Sub ExpandParents(rowToExpand As ReportRow)
    If IsSomething(rowToExpand) Then
        If IsSomething(rowToExpand.ParentRow) Then
            ExpandParents rowToExpand.ParentRow
        End If
        rowToExpand.Expanded = True
    End If
End Sub

Private Sub grilla_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set ptp = operarios.item(rowIndex)
    Values(1) = ptp.Empleado.LegajoAndNombreCompleto
    Values(2) = ptp.PlaneamientoTiempoProceso.Tarea.Description
    Values(3) = ptp.PlaneamientoTiempoProceso.idpedido & " - " & ptp.PlaneamientoTiempoProceso.item
    Values(4) = ptp.FechaInicioTarea
    Values(5) = funciones.Hours2HourMinute(RedondearDecimales(DateDiff("n", ptp.FechaInicioTarea, Now) / 60, 2))
    Exit Sub
err1:
End Sub


Private Sub Timer1_Timer()
    llenarGrilla
End Sub
