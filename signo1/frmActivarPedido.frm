VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#12.0#0"; "CODEJO~3.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmActivarPedido 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activar OT Nº"
   ClientHeight    =   7410
   ClientLeft      =   6960
   ClientTop       =   750
   ClientWidth     =   13860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmActivarPedido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   13860
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Contaduría: Faltantes de facturar"
      Height          =   390
      Left            =   2910
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1140
      Width           =   2655
   End
   Begin XtremeSuiteControls.PushButton PushButton4 
      Height          =   480
      Left            =   4155
      TabIndex        =   24
      Top             =   6360
      Width           =   1545
      _Version        =   786432
      _ExtentX        =   2725
      _ExtentY        =   847
      _StockProps     =   79
      Caption         =   "Etiquetas Selec"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdActivar 
      Height          =   375
      Left            =   2790
      TabIndex        =   21
      Top             =   4425
      Width           =   2895
      _Version        =   786432
      _ExtentX        =   5106
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Activar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   420
      Left            =   3015
      TabIndex        =   18
      Top             =   5850
      Width           =   2520
      _Version        =   786432
      _ExtentX        =   4445
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Materialización"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   420
      Left            =   3015
      TabIndex        =   17
      Top             =   5370
      Width           =   2520
      _Version        =   786432
      _ExtentX        =   4445
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Resúmen de Materiales"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdChekAll 
      Height          =   480
      Left            =   2715
      TabIndex        =   16
      Top             =   6375
      Width           =   1410
      _Version        =   786432
      _ExtentX        =   2487
      _ExtentY        =   847
      _StockProps     =   79
      Caption         =   "Marcar Todas"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Frame fraDetalles 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Detalles de la Orden de Trabajo"
      Height          =   7335
      Left            =   5805
      TabIndex        =   13
      Top             =   30
      Width           =   7980
      Begin XtremeReportControl.ReportControl ReportControl 
         Height          =   7065
         Left            =   60
         TabIndex        =   14
         Top             =   150
         Width           =   7845
         _Version        =   786432
         _ExtentX        =   13838
         _ExtentY        =   12462
         _StockProps     =   64
         BorderStyle     =   3
         PreviewMode     =   -1  'True
         AllowColumnRemove=   0   'False
         AllowColumnReorder=   0   'False
         AllowColumnSort =   0   'False
         AllowEdit       =   -1  'True
         ShowHeaderRows  =   -1  'True
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Rutas"
      Height          =   855
      Left            =   4320
      TabIndex        =   3
      Top             =   3495
      Width           =   1350
      Begin VB.CheckBox chkPortada 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Principal"
         Height          =   255
         Left            =   165
         TabIndex        =   5
         Top             =   255
         Width           =   900
      End
      Begin VB.CheckBox chkRuta 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ruta"
         Height          =   255
         Left            =   165
         TabIndex        =   4
         Top             =   495
         Value           =   1  'Checked
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Impresión"
      Height          =   3390
      Left            =   2790
      TabIndex        =   8
      Top             =   30
      Width           =   2895
      Begin VB.CommandButton cmdRutasPorSector 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rutas ( Todo por sector )"
         Height          =   390
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2415
         Width           =   2655
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Etiquetas Formato Viejo"
         Height          =   390
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Se imprimen las etiquetas que se chequearon"
         Top             =   2850
         Width           =   2655
      End
      Begin VB.CommandButton cmdRutas2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rutas ( Seleccionado )"
         Height          =   390
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1980
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Portada"
         Height          =   390
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Planeamiento"
         Height          =   390
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1545
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contaduría"
         Height          =   390
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   675
         Width           =   2655
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sectores"
      Height          =   7320
      Left            =   75
      TabIndex        =   6
      Top             =   30
      Width           =   2595
      Begin MSComctlLib.ListView lstSectores 
         Height          =   7080
         Left            =   45
         TabIndex        =   7
         Top             =   195
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   12488
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sector"
            Object.Width           =   2788
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Portada"
      Height          =   855
      Left            =   2820
      TabIndex        =   0
      Top             =   3495
      Width           =   1380
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Principal"
         Height          =   255
         Left            =   165
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   885
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Tiempos"
         Height          =   255
         Left            =   165
         TabIndex        =   1
         Top             =   480
         Value           =   1  'Checked
         Width           =   930
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   420
      Left            =   3015
      TabIndex        =   20
      Top             =   4890
      Width           =   2520
      _Version        =   786432
      _ExtentX        =   4445
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Resúmen de Fabricación"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnTareasSector 
      Height          =   480
      Left            =   2715
      TabIndex        =   22
      Top             =   6870
      Width           =   1410
      _Version        =   786432
      _ExtentX        =   2487
      _ExtentY        =   847
      _StockProps     =   79
      Caption         =   "Marcar Tareas de  Sector:"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboSector 
      Height          =   315
      Left            =   4185
      TabIndex        =   23
      Top             =   6945
      Width           =   1545
      _Version        =   786432
      _ExtentX        =   2725
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuChequearTareas 
         Caption         =   "Chequear tareas"
      End
      Begin VB.Menu mnuDeschequearTareas 
         Caption         =   "Deschequear tareas"
      End
   End
End
Attribute VB_Name = "frmActivarPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim idpedido As Long
Dim cod As Long
Dim procDef As Boolean
Dim baseP As New classPlaneamiento
Dim baseSP As New classSignoplast
Dim stock As New classStock
Dim rs As Recordset
Dim vpedido As OrdenTrabajo
Private porSector As Boolean
Private parentRow2Check As ReportRow
Private sectores As New Collection
Private sectoresInvolucrados As New Collection
Public Property Set Pedido(T As OrdenTrabajo)
    Set vpedido = T
    Set vpedido.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(vpedido.Id)
    idpedido = T.Id
    Me.caption = "Activar OT Nº " & T.IdFormateado
End Property

Private Sub btnTareasSector_Click()
    If cboSector.ListIndex = -1 Then
        MsgBox "Seleccione un sector.", vbExclamation
    Else

        Dim repRow As ReportRow
        Dim tareasSector As Collection
        Set tareasSector = DAOTareas.FindAll("t.id_sector = " & Me.cboSector.ItemData(Me.cboSector.ListIndex))

        For Each repRow In Me.ReportControl.rows
            ProcessRowForSector repRow, tareasSector
        Next repRow
        Me.ReportControl.Redraw
    End If
End Sub

Private Sub ProcessRowForSector(row As ReportRow, tareasSector As Collection)
    Dim rowChild As ReportRow
    If row.record.Tag < 0 Then
        row.record.item(1).Checked = funciones.BuscarEnColeccion(tareasSector, Mid$(row.record.item(1).caption, 8, InStr(8, row.record.item(1).caption, " ") - 8))
    End If

    For Each rowChild In row.Childs
        ProcessRowForSector rowChild, tareasSector
    Next rowChild
End Sub


Private Sub cmdActivar_Click()
    Dim A As VbMsgBoxResult
    A = MsgBox("¿Desea darle curso a este pedido?", vbYesNo + vbQuestion, "Confirmación")

    If A = vbYes Then
        If baseP.procesos_definidos(cod) Then
            If DAOOrdenTrabajo.PonerEnProduccion(vpedido) Then

                Dim EVENTO As New clsEventoObserver
                Set EVENTO.Elemento = vpedido
                EVENTO.EVENTO = modificar_
                Set EVENTO.Originador = Me
                Channel.Notificar EVENTO, ordenesTrabajo
                Unload Me
            End If

        Else
            MsgBox "Debe tener definidos los procesos para poder activar el pedido!", vbCritical, "Error"
        End If
    End If

End Sub

Private Sub cmdChekAll_Click()
    Dim repRow As ReportRow
    For Each repRow In Me.ReportControl.rows
        repRow.record.item(1).Checked = Not repRow.record.item(1).Checked
    Next repRow
    Me.ReportControl.Redraw
End Sub

Private Sub cmdRutas2_Click()
    On Error GoTo E
    If Not porSector Then Me.CommonDialog1.ShowPrinter

    Dim rec As ReportRecord
    Dim rec1 As ReportRecord
    Dim rec2 As ReportRecord
    Dim rec3 As ReportRecord
    Dim rec4 As ReportRecord
    Dim tmpDetalle As DetalleOrdenTrabajo
    Dim tmpDetalleConj1 As DetalleOTConjuntoDTO
    Dim tmpDetalleConj2 As DetalleOTConjuntoDTO
    Dim tmpDetalleConj3 As DetalleOTConjuntoDTO
    Dim tmpDetalleConj4 As DetalleOTConjuntoDTO
    Dim tmpCol As Collection

    For Each rec In Me.ReportControl.Records
        Set tmpDetalle = vpedido.Detalles(CStr(rec.Tag))

        If rec.item(1).Checked Then
            If tmpDetalle.Pieza.EsConjunto Then
                DAOPieza.informePiezaArbol tmpDetalle.Id, 0    'rec.Item(2).value
            End If

            DAOOrdenTrabajo.ImprimirRuta vpedido, tmpDetalle
        End If

        For Each rec1 In rec.Childs
            Set tmpDetalleConj1 = Nothing

            If rec1.Tag > 0 Then
                Set tmpCol = DAODetalleOrdenTrabajo.FindAllConjunto(, , "dpc.Id = " & rec1.Tag)
                If tmpCol.count > 0 Then
                    Set tmpDetalleConj1 = tmpCol(1)
                End If
            End If

            If rec1.item(1).Checked Then
                If rec1.Tag < 0 Then
                    ImprimirHojaTarea rec1.Tag, tmpDetalle.Pieza, tmpDetalle
                Else

                    DAOOrdenTrabajo.ImprimirRuta vpedido, tmpDetalle, tmpDetalleConj1

                End If
            End If

            For Each rec2 In rec1.Childs
                Set tmpDetalleConj2 = Nothing

                If rec2.Tag > 0 Then
                    Set tmpCol = DAODetalleOrdenTrabajo.FindAllConjunto(, , "dpc.Id = " & rec2.Tag)
                    If tmpCol.count > 0 Then
                        Set tmpDetalleConj2 = tmpCol(1)
                    End If
                End If

                If rec2.item(1).Checked Then
                    If rec2.Tag < 0 Then    'es tarea
                        ImprimirHojaTarea rec2.Tag, tmpDetalleConj1.Pieza, tmpDetalle
                    Else
                        DAOOrdenTrabajo.ImprimirRuta vpedido, tmpDetalle, tmpDetalleConj2
                    End If
                End If

                For Each rec3 In rec2.Childs
                    Set tmpDetalleConj3 = Nothing

                    If rec3.Tag > 0 Then
                        Set tmpCol = DAODetalleOrdenTrabajo.FindAllConjunto(, , "dpc.Id = " & rec3.Tag)
                        If tmpCol.count > 0 Then
                            Set tmpDetalleConj3 = tmpCol(1)
                        End If
                    End If

                    If rec3.item(1).Checked Then
                        If rec3.Tag < 0 Then
                            ImprimirHojaTarea rec3.Tag, tmpDetalleConj2.Pieza, tmpDetalle
                        Else
                            DAOOrdenTrabajo.ImprimirRuta vpedido, tmpDetalle, tmpDetalleConj3
                        End If
                    End If

                    For Each rec4 In rec3.Childs
                        Set tmpDetalleConj4 = Nothing

                        If rec4.Tag > 0 Then
                            Set tmpCol = DAODetalleOrdenTrabajo.FindAllConjunto(, , "dpc.Id = " & rec4.Tag)
                            If tmpCol.count > 0 Then
                                Set tmpDetalleConj4 = tmpCol(1)
                            End If
                        End If

                        If rec4.item(1).Checked Then
                            If rec4.Tag < 0 Then
                                ImprimirHojaTarea rec4.Tag, tmpDetalleConj3.Pieza, tmpDetalle
                            Else
                                DAOOrdenTrabajo.ImprimirRuta vpedido, tmpDetalle, tmpDetalleConj4
                            End If
                        End If
                    Next


                Next
            Next
        Next
    Next

    Exit Sub
E:
    If Err.Source = "CommonDialog" And Err.Number = 32755 Then Exit Sub
    MsgBox Err.Description, vbCritical, "Error"
    'Resume
End Sub


Private Sub ImprimirHojaTarea(ptpId As Long, P As Pieza, d As DetalleOrdenTrabajo)
'Exit Sub  '-> para q no las imprima ahora
    On Error GoTo E

    Dim ptp As PlaneamientoTiempoProceso
    Set ptp = DAOTiemposProceso.FindById(ptpId * -1)

    informe_tareas.Sections("Sección4").Controls("lblOT").caption = "Orden de trabajo " & vpedido.IdFormateado & " | Item: " & d.item
    informe_tareas.Sections("Sección4").Controls("lblCliente2").caption = vpedido.ClienteFacturar.razon
    informe_tareas.Sections("Sección4").Controls("lblReferenciaOT").caption = vpedido.descripcion
    informe_tareas.Sections("Sección4").Controls("lblReferencia").caption = P.nombre
    informe_tareas.Sections("Sección4").Controls("lblfechaEntrega").caption = d.FechaEntrega
    informe_tareas.Sections("Sección4").Controls("barCode").caption = "*" & Format(ptp.Id, "00000000") & "*"
    informe_tareas.Sections("Sección4").Controls("lblCliente").caption = vpedido.cliente.razon

    informe_tareas.Sections("Sección2").Controls("Etiqueta4").caption = "Tarea: " & ptp.Tarea.Id & " - " & ptp.Tarea.Tarea & " (Sector: " & ptp.Tarea.Sector.Sector & ")"


    Dim empleadosHabilitados As String
    Dim emp As clsEmpleado
    For Each emp In DAOEmpleados.GetAllByTareaId(ptp.Tarea.Id)
        empleadosHabilitados = empleadosHabilitados & truncar(emp.LegajoAndNombreCompleto, 20) & vbNewLine
    Next
    informe_tareas.Sections("Sección1").Controls("Etiqueta8").caption = empleadosHabilitados



    Dim rs As Recordset
    Set rs = conectar.RSFactory("SELECT 1")
    Set informe_tareas.DataSource = rs

    informe_tareas.PrintReport False

    Exit Sub
E:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdRutasPorSector_Click()
    On Error GoTo E
    If MsgBox("Va a imprimir toda la OT completa ordenada por sector, sin importar lo que este tildado." & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    porSector = True

    Dim rec As ReportRecord
    Dim sectores_id As New Dictionary


    Me.CommonDialog1.ShowPrinter


    '1ro hay que capturar los sectores involucrados
    For Each rec In Me.ReportControl.Records
        ExplorarSectorId rec, sectores_id
    Next rec

    'tildo las tareas por sector y voy imprimiendo
    Dim sector_id As Variant
    For Each sector_id In sectores_id.Items
        UncheckAll

        For Each rec In Me.ReportControl.Records
            TildarPorSectorId rec, CLng(sector_id)
            '        If Not HayTareaChecked(rec) Then 'si no tiene ninguna tarea chequeada, al pedo imprimir portada de la pieza para ese sector
            '            rec.Item(1).Checked = False
            '        End If
        Next rec
        cmdRutas2_Click
    Next sector_id


    porSector = False
    Exit Sub
E:
    If 32755 <> Err.Number Then
        MsgBox Err.Description, vbOKOnly + vbCritical
    End If
End Sub

'Private Function HayTareaChecked(rec As ReportRecord) As Boolean
'    Dim rec2 As ReportRecord
'
'    If rec.Tag < 0 Then
'        HayTareaChecked = HayTareaChecked Or rec.Item(1).Checked
'    End If
'
'    For Each rec2 In rec.Childs
'        HayTareaChecked = HayTareaChecked Or HayTareaChecked(rec2)
'    Next rec2
'
'End Function

Private Sub UncheckAll()
    Dim repRow As ReportRow
    For Each repRow In Me.ReportControl.rows
        repRow.record.item(1).Checked = False
    Next repRow
    Me.ReportControl.Redraw
End Sub

Private Sub TildarPorSectorId(rec As ReportRecord, sector_id As Long)
    Dim rec2 As ReportRecord
    Dim ptp As PlaneamientoTiempoProceso


    If rec.Tag < 0 Then    'es planeamiento tiempo proceso
        Set ptp = DAOTiemposProceso.FindById(rec.Tag * -1)
        If IsSomething(ptp) Then
            rec.item(1).Checked = (ptp.Tarea.Sector.Id = sector_id)
        End If
    Else    'es pieza
        rec.item(1).Checked = True
    End If

    For Each rec2 In rec.Childs
        TildarPorSectorId rec2, sector_id
    Next rec2
End Sub

Private Sub ExplorarSectorId(rec As ReportRecord, ByRef sectores_id As Dictionary)
    Dim rec2 As ReportRecord
    Dim ptp As PlaneamientoTiempoProceso

    For Each rec2 In rec.Childs
        If rec2.Tag < 0 Then    'es planeamiento tiempo proceso
            Set ptp = DAOTiemposProceso.FindById(rec2.Tag * -1)
            If IsSomething(ptp) Then
                If Not sectores_id.Exists(CStr(ptp.Tarea.Sector.Id)) Then
                    sectores_id.Add CStr(ptp.Tarea.Sector.Id), ptp.Tarea.Sector.Id
                End If
            End If
        End If
        ExplorarSectorId rec2, sectores_id
    Next rec2

End Sub


Private Sub Command1_Click()
    On Error GoTo err4
    Dim cod As Integer
    cod = CInt(idpedido)

    Me.CommonDialog1.ShowPrinter
    ImprimirPortadas


    Exit Sub

err4:
    If Err.Source = "CommonDialog" And Err.Number = 32755 Then Exit Sub
    MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub Command11_Click()
    On Error GoTo er1
    frmPrincipal.cd.ShowPrinter
    Dim rec As ReportRecord
    Dim tmpDetalle As DetalleOrdenTrabajo
    Dim l As Long
    Dim codBar As String
    Dim linea1, linea2, linea3, linea4, linea5

    For Each rec In Me.ReportControl.Records
        Set tmpDetalle = vpedido.Detalles(CStr(rec.Tag))

        If rec.item(1).Checked Then
            For l = 1 To tmpDetalle.CantidadPedida
                Set rs = conectar.RSFactory("Select p.descripcion,dp.idPedido,c.razon,dp.item,s.detalle from detalles_pedidos dp inner join pedidos p on dp.idPedido=p.id inner join stock s on dp.idPieza=s.id inner join clientes c on p.idCliente=c.id  where dp.id=" & tmpDetalle.Id)
                If Not rs.EOF And Not rs.BOF Then
                    codBar = "*" & Format(tmpDetalle.Id, "00000000") & "*"
                    linea1 = "O/T:" & vpedido.IdFormateado
                    linea2 = "Cliente:" & rs!razon
                    linea3 = rs!descripcion
                    linea4 = rs!item & " " & rs!detalle
                    linea5 = "Proveedor: SIGNO PLAST S.A."
                    baseP.etiquetas_informe codBar, linea1, linea2, linea3, linea4, linea5
                End If
            Next l
        End If
    Next
er1:
    If Err.Source = "CommonDialog" And Err.Number = 32755 Then Exit Sub
    MsgBox Err.Description, vbCritical, "Error"
End Sub



Private Sub Command3_Click()
    Command1_Click
    Command4_Click
    Command5_Click

End Sub




Private Sub Command2_Click()
    If Permisos.sistemaVerPrecios Then
        Dim cod As Integer
        cod = CInt(idpedido)
        DAOOrdenTrabajo.ImprimirFaltantesFacturacion vpedido




    Else
        sinAcceso
    End If
End Sub

Private Sub Command4_Click()
    If Permisos.sistemaVerPrecios Then
        Dim cod As Integer
        cod = CInt(idpedido)
        baseP.informePedidoContaduria cod, True, vpedido

    Else
        sinAcceso
    End If

End Sub

Private Sub Command5_Click()

    DAOOrdenTrabajo.informePedidoPlaneamiento CInt(idpedido), False, DAOTiemposProceso.SectorColl2RS(sectoresInvolucrados)
    DAOOrdenTrabajo.informePiezaMateriales CDbl(idpedido), 1, True
    DAOOrdenTrabajo.imprimirEtiquetas CLng(idpedido)

    If MsgBox("¿Imprimir las portadas para cada nave?", vbYesNo, "Confirmación") = vbYes Then
        DAOOrdenTrabajo.informePedido vpedido, False, "NAVE 1"   'imprimo la portada
        DAOOrdenTrabajo.grillaTiempos vpedido, False, "NAVE 1"
        DAOOrdenTrabajo.informePedido vpedido, False, "NAVE 2"   'imprimo la portada
        DAOOrdenTrabajo.grillaTiempos vpedido, False, "NAVE 2"
    End If

End Sub


Private Sub ImprimirPortadas()
    Dim sec As clsSector

    For Each sec In sectoresInvolucrados

        If Me.Check2.value Then DAOOrdenTrabajo.grillaTiempos vpedido, False, sec.Sector
        If Me.Check1.value Then DAOOrdenTrabajo.informePedido vpedido, False, sec.Sector

    Next sec

End Sub


Private Sub Command7_Click()
    Unload Me
End Sub
Private Sub Command9_Click()
    frmPlaneamientoDefinirCrono.Show 1
End Sub


Private Sub Form_Activate()
    Dim c As Long

    cod = CLng(idpedido)
    Dim tmpSector As clsSector
    Set sectores = DAOSectores.GetAll()
    Me.lstSectores.ListItems.Clear
    For Each tmpSector In sectores
        Me.lstSectores.ListItems.Add , , tmpSector.Sector
        Me.lstSectores.ListItems(Me.lstSectores.ListItems.count).Tag = tmpSector.Id
    Next tmpSector
    Dim r As Recordset
    Dim rs As Recordset

    Set sectoresInvolucrados = DAOTiemposProceso.GetSectoresByIdPedido(vpedido.Id)
    c = 0

    Dim x As Long
    Dim i As Long

    Me.cboSector.Clear
    For x = 1 To sectoresInvolucrados.count
        For i = 1 To Me.lstSectores.ListItems.count
            If sectoresInvolucrados(x).Id = Me.lstSectores.ListItems(i).Tag Or (Me.lstSectores.ListItems(i).Tag = 19 Or Me.lstSectores.ListItems(i).Tag = 2) Then
                Me.lstSectores.ListItems(i).Checked = True
                lstSectores_ItemCheck Me.lstSectores.ListItems(i)
            End If
        Next i

        Me.cboSector.AddItem sectoresInvolucrados(x).Sector
        Me.cboSector.ItemData(Me.cboSector.NewIndex) = sectoresInvolucrados(x).Id
    Next x
    '------------------------------------------


    Dim deta As DetalleOrdenTrabajo
    Dim P As Pieza
    Dim tmpdeta2 As DetalleOTConjuntoDTO
    Dim tmpdeta3 As DetalleOTConjuntoDTO
    Dim tmpdeta4 As DetalleOTConjuntoDTO


    Me.ReportControl.Records.DeleteAll

    Dim record As ReportRecord
    Dim Record2 As ReportRecord
    Dim Record3 As ReportRecord
    Dim Record4 As ReportRecord
    Dim idpos As String
    Dim pos As Integer
    Dim pos2 As Integer
    Dim pos3 As Integer
    Dim pos4 As Integer
    Dim item As ReportRecordItem

    For Each deta In vpedido.Detalles
        Set record = Me.ReportControl.Records.Add
        record.Tag = deta.Id
        record.AddItem deta.item
        Set item = record.AddItem(deta.Pieza.nombre)
        item.HasCheckbox = True
        record.PreviewText = deta.Nota
        record.AddItem deta.CantidadPedida
        record.AddItem deta.FechaEntrega

        If deta.EtiquetasImpresas > 0 Then
            record.AddItem deta.EtiquetasImpresas
        Else
            record.AddItem deta.CantidadPedida
        End If

        AddTareas record, deta.Id

        If deta.Pieza.EsConjunto Then

            For Each tmpdeta2 In DAODetalleOrdenTrabajo.FindAllConjunto(deta.Id, deta.Pieza.Id)
                Set Record2 = record.Childs.Add()
                Record2.Tag = tmpdeta2.Id
                Record2.AddItem vbNullString
                Set item = Record2.AddItem(tmpdeta2.IdentificadorPosicion & " - " & tmpdeta2.Pieza.nombre)
                item.HasCheckbox = True
                Record2.AddItem tmpdeta2.Cantidad & " (" & tmpdeta2.CantidadTotalStatic & " total)"
                Record2.AddItem vbNullString

                AddTareas Record2, deta.Id, tmpdeta2.Id
                If tmpdeta2.Pieza.EsConjunto Then

                    For Each tmpdeta3 In DAODetalleOrdenTrabajo.FindAllConjunto(deta.Id, tmpdeta2.Pieza.Id)
                        Set Record3 = Record2.Childs.Add
                        Record3.Tag = tmpdeta3.Id
                        Record3.AddItem vbNullString
                        Set item = Record3.AddItem(tmpdeta3.IdentificadorPosicion & " - " & tmpdeta3.Pieza.nombre)
                        item.HasCheckbox = True
                        Record3.AddItem tmpdeta3.Cantidad & " (" & tmpdeta3.CantidadTotalStatic & " total)"
                        Record3.AddItem vbNullString
                        AddTareas Record3, deta.Id, tmpdeta3.Id
                        pos3 = 0
                        If tmpdeta3.Pieza.EsConjunto Then
                            For Each tmpdeta4 In DAODetalleOrdenTrabajo.FindAllConjunto(deta.Id, tmpdeta3.Pieza.Id)
                                Set Record4 = Record3.Childs.Add
                                Record4.Tag = tmpdeta4.Id
                                Record4.AddItem vbNullString
                                Set item = Record4.AddItem(tmpdeta4.IdentificadorPosicion & " - " & tmpdeta4.Pieza.nombre)
                                item.HasCheckbox = True
                                Record4.AddItem tmpdeta4.Cantidad & " (" & tmpdeta4.CantidadTotalStatic & " total)"
                                Record4.AddItem vbNullString

                                AddTareas Record4, deta.Id, tmpdeta4.Id

                            Next tmpdeta4
                        End If

                    Next tmpdeta3
                End If
            Next tmpdeta2
        End If
    Next

    Me.ReportControl.Populate

End Sub

Private Sub AddTareas(ByRef rec As ReportRecord, ByRef idDetallePedido As Long, Optional ByRef idDetallePedidoConjunto As Long = 0)
    Dim rechijo As ReportRecord
    Dim item As ReportRecordItem

    Dim ptp As PlaneamientoTiempoProceso

    'For Each ptp In DAOTiemposProceso.FindAllByDetallePedidoIdAndPiezaId(idDetallePedido, P.Id)
    For Each ptp In DAOTiemposProceso.FindAllByDetallePedidoId(idDetallePedido, idDetallePedidoConjunto)
        Set rechijo = rec.Childs.Add
        rechijo.Tag = (ptp.Id * -1)    'negativo para distinguir de las piezas
        rechijo.AddItem vbNullString
        Set item = rechijo.AddItem("Tarea: " & ptp.Tarea.Id & " - " & ptp.Tarea.Tarea)
        item.HasCheckbox = True
        rechijo.AddItem vbNullString
        rechijo.AddItem vbNullString
    Next ptp

End Sub



Private Sub Form_Load()
    FormHelper.Customize Me
    Me.lstSectores.ColumnHeaders(1).Width = Me.lstSectores.Width - 300

    '--------------------
    Dim Column As ReportColumn
    Set Column = Me.ReportControl.Columns.Add(0, "Item", 10, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Editable = False

    Set Column = Me.ReportControl.Columns.Add(1, "Detalle", 65, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.TreeColumn = True
    Column.AllowDrag = False
    Column.AllowRemove = False


    Set Column = Me.ReportControl.Columns.Add(2, "Cantidad", 12, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Editable = False

    Set Column = Me.ReportControl.Columns.Add(3, "F. Entrega", 17, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Editable = False
    Set Column = Me.ReportControl.Columns.Add(4, "Etiquetas", 12, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Editable = True



    Me.ReportControl.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.ReportControl.PaintManager.VerticalGridStyle = xtpGridSmallDots

    procDef = False

    'Me.caption = caption & " (" & Name & ")"

End Sub

Private Sub idpedido_Click()

End Sub

Private Sub lstParaImprimir_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'lstStock.Sorted = True
'LstOrdenar Me.lstParaImprimir, CInt(ColumnHeader.Index)
End Sub




Private Sub lstSectores_ItemCheck(ByVal item As MSComctlLib.ListItem)
    If item.Checked Then
        If Not funciones.BuscarEnColeccion(sectoresInvolucrados, CStr(item.Tag)) Then
            sectoresInvolucrados.Add sectores(CStr(item.Tag)), CStr(item.Tag)
        End If
    Else
        If funciones.BuscarEnColeccion(sectoresInvolucrados, CStr(item.Tag)) Then
            sectoresInvolucrados.remove CStr(item.Tag)
        End If
    End If
End Sub

Private Sub mnuChequearTareas_Click()
    Dim rowHija As ReportRow

    If Not parentRow2Check Is Nothing Then
        For Each rowHija In parentRow2Check.Childs
            rowHija.record.item(1).Checked = True
        Next rowHija

        Me.ReportControl.Redraw
    End If
End Sub

Private Sub mnuDeschequearTareas_Click()
    Dim rowHija As ReportRow

    If Not parentRow2Check Is Nothing Then
        For Each rowHija In parentRow2Check.Childs
            rowHija.record.item(1).Checked = False
        Next rowHija

        Me.ReportControl.Redraw
    End If
End Sub

Private Sub PushButton1_Click()
    DAOOrdenTrabajo.informePiezaMateriales vpedido.Id, 1, True
End Sub

Private Sub PushButton2_Click()
    frmMaterializacion.Id = vpedido.Id
    frmMaterializacion.Ot = True
    frmMaterializacion.Show
End Sub

Private Sub PushButton3_Click()
    DAOOrdenTrabajo.InformePiezasFabricadas vpedido
End Sub





Private Sub PushButton4_Click()
    On Error GoTo E




    Dim c As Long
    Dim rec As ReportRecord
    Dim rec1 As ReportRecord
    Dim rec2 As ReportRecord
    Dim rec3 As ReportRecord
    Dim rec4 As ReportRecord



    Dim tmpDetalle As DetalleOrdenTrabajo
    Dim tmpDetalleConj1 As DetalleOTConjuntoDTO
    Dim tmpDetalleConj2 As DetalleOTConjuntoDTO
    Dim tmpDetalleConj3 As DetalleOTConjuntoDTO
    Dim tmpDetalleConj4 As DetalleOTConjuntoDTO
    Dim tmpCol As Collection
    Dim Cant As Double
    Dim cant1 As Double
    Dim cant2 As Double
    Dim cant3 As Double
    Dim cant4 As Double
    Dim it As ReportRecordItem
    c = 0

    If Not LabelHelper.PrintEtiquetaPedido(vpedido) Then GoTo E

    For Each rec In Me.ReportControl.Records
        Set tmpDetalle = vpedido.Detalles(CStr(rec.Tag))

        If rec.item(1).Checked Then

            Cant = rec.item(4).value
            If Not LabelHelper.PrintEtiquetaDetallePedido(vpedido, tmpDetalle, Cant) Then GoTo E
            c = c + 1
        End If

        For Each rec1 In rec.Childs

            Set tmpDetalleConj1 = Nothing

            If rec1.Tag > 0 Then
                Set tmpCol = DAODetalleOrdenTrabajo.FindAllConjunto(, , "dpc.Id = " & rec1.Tag)
                If tmpCol.count > 0 Then
                    Set tmpDetalleConj1 = tmpCol(1)
                End If
            End If

            If rec1.item(1).Checked Then
                If rec1.Tag > 0 Then
                    cant1 = rec1.item(5).value
                    If Not LabelHelper.PrintEtiquetaDetallePedido(vpedido, tmpDetalle, cant1, tmpDetalleConj1) Then GoTo E
                    c = c + 1
                End If
            End If

            For Each rec2 In rec1.Childs

                Set tmpDetalleConj2 = Nothing

                If rec2.Tag > 0 Then

                    Set tmpCol = DAODetalleOrdenTrabajo.FindAllConjunto(, , "dpc.Id = " & rec2.Tag)
                    If tmpCol.count > 0 Then
                        Set tmpDetalleConj2 = tmpCol(1)
                    End If
                End If

                If rec2.item(1).Checked Then
                    If rec2.Tag > 0 Then    'no es tarea
                        cant2 = rec2.item(4).value
                        If Not LabelHelper.PrintEtiquetaDetallePedido(vpedido, tmpDetalle, cant2, tmpDetalleConj2) Then GoTo E
                        c = c + 1
                    End If
                End If

                For Each rec3 In rec2.Childs

                    Set tmpDetalleConj3 = Nothing

                    If rec3.Tag > 0 Then
                        Set tmpCol = DAODetalleOrdenTrabajo.FindAllConjunto(, , "dpc.Id = " & rec3.Tag)
                        If tmpCol.count > 0 Then
                            Set tmpDetalleConj3 = tmpCol(1)
                        End If
                    End If

                    If rec3.item(1).Checked Then
                        If rec3.Tag > 0 Then
                            cant3 = rec3.item(4).value
                            If Not LabelHelper.PrintEtiquetaDetallePedido(vpedido, tmpDetalle, cant3, tmpDetalleConj3) Then GoTo E
                            c = c + 1
                        End If
                    End If

                    For Each rec4 In rec3.Childs

                        Set tmpDetalleConj4 = Nothing

                        If rec4.Tag > 0 Then
                            Set tmpCol = DAODetalleOrdenTrabajo.FindAllConjunto(, , "dpc.Id = " & rec4.Tag)
                            If tmpCol.count > 0 Then
                                Set tmpDetalleConj4 = tmpCol(1)
                            End If
                        End If

                        If rec4.item(1).Checked Then
                            If rec4.Tag > 0 Then
                                cant4 = rec4.item(4).value
                                If Not LabelHelper.PrintEtiquetaDetallePedido(vpedido, tmpDetalle, cant4, tmpDetalleConj4) Then GoTo E
                                c = c + 1
                            End If
                        End If
                    Next


                Next
            Next
        Next
    Next


    If c >= 3 Then If Not LabelHelper.PrintEtiquetaPedido(vpedido) Then GoTo E


    Exit Sub
E:
    If Err.Source = "CommonDialog" And Err.Number = 32755 Then Exit Sub
    MsgBox Err.Description, vbCritical, "Error"
    'Resume
End Sub


Private Sub ReportControl_ItemCheck(ByVal row As XtremeReportControl.IReportRow, ByVal item As XtremeReportControl.IReportRecordItem)
    Dim repRow As ReportRow
    'For Each repRow In Row.Childs
    ' repRow.record.item(1).Checked = item.Checked
    '   ReportControl_ItemCheck repRow, repRow.record.item(1)
    'Next repRow
End Sub


Private Sub ReportControl_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Set parentRow2Check = Nothing
    On Error GoTo err1
    If Button = 2 Then
        Dim Info As ReportHitTestInfo
        Set Info = Me.ReportControl.HitTest(x, y)
        If Not Info Is Nothing Then
            If Info.row.record.Tag < 0 Then    'es tarea
                Set parentRow2Check = Info.row.ParentRow
                Me.PopupMenu Me.mnuMain
            End If
        End If
    End If
    Exit Sub
err1:
End Sub

