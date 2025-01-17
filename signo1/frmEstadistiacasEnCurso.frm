VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmEstadistiacasEnCurso 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estadísticas -> Carga de producción actual"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   1935
   ClientWidth     =   14865
   Icon            =   "frmEstadistiacasEnCurso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   14865
   Begin GridEX20.GridEX GridEX1 
      Height          =   7830
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   13811
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   3
      Column(1)       =   "frmEstadistiacasEnCurso.frx":000C
      Column(2)       =   "frmEstadistiacasEnCurso.frx":0178
      Column(3)       =   "frmEstadistiacasEnCurso.frx":02E8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmEstadistiacasEnCurso.frx":03DC
      FormatStyle(2)  =   "frmEstadistiacasEnCurso.frx":0514
      FormatStyle(3)  =   "frmEstadistiacasEnCurso.frx":05C4
      FormatStyle(4)  =   "frmEstadistiacasEnCurso.frx":0678
      FormatStyle(5)  =   "frmEstadistiacasEnCurso.frx":0750
      FormatStyle(6)  =   "frmEstadistiacasEnCurso.frx":0808
      ImageCount      =   0
      PrinterProperties=   "frmEstadistiacasEnCurso.frx":08E8
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   1095
   End
   Begin MSChart20Lib.MSChart grafica_torta 
      Height          =   3690
      Left            =   4980
      OleObjectBlob   =   "frmEstadistiacasEnCurso.frx":0AC0
      TabIndex        =   3
      Top             =   4755
      Visible         =   0   'False
      Width           =   9675
   End
   Begin MSChart20Lib.MSChart grafica 
      Height          =   4815
      Left            =   4980
      OleObjectBlob   =   "frmEstadistiacasEnCurso.frx":2902
      TabIndex        =   5
      Top             =   -30
      Visible         =   0   'False
      Width           =   9675
   End
   Begin VB.Label lbltotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   8040
      Width           =   45
   End
   Begin VB.Label estado 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   495
      Left            =   11520
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmEstadistiacasEnCurso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vpresu As Boolean
Dim ids As Integer
Dim est As Integer
Dim vConjGrabado As Boolean
Dim tmpSectorTiempo As DTOSectoresTiempo
Public VerTiempoRestante
Private ARRSECTORES()
Private arrTareas()

Public col As Collection
Dim dto As DTOSectoresTiempo
Private mlistadtopiezacantidad As Collection

Public IDsOTAvance As New Collection
Private avancesOT As New Collection


Public Function GetTiemposPorCategoria() As Dictionary
    Dim col1 As New Dictionary
    Dim dto1 As DTOTareaTiempo
    Dim tmp As Double
    For Each dto In col
        For Each dto1 In dto.ListaDtoTareaTiempo

            If col1.Exists(dto1.Tarea.CategoriaSueldo.id) Then

                col1.item(dto1.Tarea.CategoriaSueldo.id) = col1.item(dto1.Tarea.CategoriaSueldo.id) + dto1.Tiempo
            Else

                col1.Add dto1.Tarea.CategoriaSueldo.id, dto1.Tiempo

            End If

        Next dto1
    Next

    Set GetTiemposPorCategoria = col1

End Function


Public Property Set listadtopiezacantidad(nvalue As Collection)
    Set mlistadtopiezacantidad = nvalue
    LlenarGrid

    If mlistadtopiezacantidad.count > 0 Then
        grafica.Visible = True
        Me.grafico
    Else
        grafica.Visible = False
        MsgBox "No hay datos para mostrar. Se cerrara la ventana", vbOKOnly, "No hay datos"
        Unload Me
    End If
End Property

Public Sub LlenarGridDesdeOT()

    Me.GridEX1.ItemCount = 0
    Me.GridEX1.ItemCount = col.count

    If col.count > 0 Then
        If IDsOTAvance.count > 0 Then
            Set avancesOT = DAOTiemposProceso.GetAvancesHsPorOTs(IDsOTAvance)
        End If

        grafica.Visible = True
        Me.grafico
    Else
        grafica.Visible = False
        MsgBox "No hay datos para mostrar. Se cerrara la ventana", vbOKOnly, "No hay datos"
        Unload Me
    End If

End Sub

Private Sub LlenarGrid()
    Set col = DAOPieza.ListaDTOTiempoPorSector(mlistadtopiezacantidad)

    Me.GridEX1.ItemCount = 0
    Me.GridEX1.ItemCount = col.count
End Sub

Public Property Let conjGrabado(nc As Boolean)
    vConjGrabado = nc
End Property

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    
    Dim col2 As New Dictionary
    Dim cs As CategoriaSueldo
    Set col2 = GetTiemposPorCategoria
    Dim d As Double
    Dim i As Variant
    Debug.Print "Categoria,Cant. Horas, Valor Hora, total"
    For Each i In col2.Keys
        Set cs = DAOCategoriaSueldo.FindAll("cs.id=" & i)(1)
        Debug.Print cs.nombre & "," & col2.item(i) & "," & funciones.FormatearDecimales(cs.Valor * 60) & "," & col2.item(i) * funciones.FormatearDecimales(cs.Valor) * 60

        d = d + col2.item(i)

    Next
    Debug.Print "total" & d

    grafica.EditCopy
    grafica_torta.EditCopy
    
    '    On Error GoTo err22
    '    Me.CommonDialog1.ShowPrinter
    '    c = Me.CommonDialog1.Copies
    '    For x = 1 To c
    '      Imprimir (est)
    '    Next x
    '    Exit Sub
    'err22:

    Load frmPrintPreview
    
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    
    GridEX1.PrintPreview frmPrintPreview.GEXPreview1
    
    GridEX1.PrintPreview frmPrintPreview.GEXPreview1

    
    frmPrintPreview.Show 1

End Sub

Public Sub GraficoTorta()
    Me.grafica_torta.Visible = True
    Dim dto1 As DTOTareaTiempo
    Dim tmpSectorTiempo As DTOSectoresTiempo
    Dim tmpTareaTiempo As DTOTareaTiempo
    ReDim arrTareas(1 To dto.ListaDtoTareaTiempo.count, 1 To 5)
    Dim i As Long
    i = 0
    For Each dto1 In dto.ListaDtoTareaTiempo
        i = i + 1
        arrTareas(i, 1) = dto1.Tarea.Tarea
        arrTareas(i, 2) = dto1.Tiempo

        If avancesOT.count > 0 Then
            If funciones.BuscarEnColeccion(avancesOT, CStr(dto1.Tarea.SectorID)) Then
                Set tmpSectorTiempo = avancesOT.item(CStr(dto1.Tarea.SectorID))
                If funciones.BuscarEnColeccion(tmpSectorTiempo.ListaDtoTareaTiempo, CStr(dto1.Tarea.id)) Then
                    Set tmpTareaTiempo = tmpSectorTiempo.ListaDtoTareaTiempo.item(CStr(dto1.Tarea.id))
                    arrTareas(i, 3) = tmpTareaTiempo.Tiempo
                    arrTareas(i, 4) = tmpTareaTiempo.Tiempo
                    arrTareas(i, 5) = tmpTareaTiempo.Tiempo
                End If

            End If
        End If
    Next


    grafica_torta.ChartData = arrTareas


    If IsSomething(tmpSectorTiempo) Then
        If tmpSectorTiempo.Tiempo > 0 Then
            grafica_torta.ColumnCount = 2
        Else
            grafica_torta.ColumnCount = 1
        End If
    Else
        grafica_torta.ColumnCount = 1
    End If




    grafica_torta.ColumnLabelCount = 1
    grafica_torta.Column = 1
    grafica_torta.ColumnLabel = "asdsad"
    grafica_torta.Refresh
End Sub




Public Sub grafico()
'    On Error GoTo e

    ReDim ARRSECTORES(1 To col.count, 1 To 5)
    Dim i As Integer
    Dim dto As DTOSectoresTiempo
    Dim c As Double
    Dim tmpSectorTiempo As DTOSectoresTiempo
    i = 0
    For Each dto In col
        i = i + 1
        ARRSECTORES(i, 1) = dto.Sector.Sector
        ARRSECTORES(i, 2) = dto.Tiempo
        If avancesOT.count > 0 Then
            If funciones.BuscarEnColeccion(avancesOT, CStr(dto.Sector.id)) Then
                Set tmpSectorTiempo = avancesOT.item(CStr(dto.Sector.id))
                ARRSECTORES(i, 4) = tmpSectorTiempo.Tiempo



            End If

        Else
            '        DAOTiemposProceso.FindAllByDetallePedidoId 89891
            ARRSECTORES(i, 4) = 0
        End If
        ARRSECTORES(i, 3) = dto.TiempoPendiente
        If dto.TotalTareasFinalizadas > 0 Then
            ARRSECTORES(i, 5) = (1 - (dto.TotalTareasFinalizadas / dto.TotalTareas)) * dto.Tiempo
        Else
            ARRSECTORES(i, 5) = dto.Tiempo
        End If
        c = c + dto.Tiempo
    Next dto

    grafica.ChartData = ARRSECTORES

    If avancesOT.count = 0 Then
        grafica.ColumnCount = 4
    Else
        grafica.ColumnCount = 4
    End If

    grafica.ColumnLabelCount = 1
    grafica.Column = 1
    grafica.ColumnLabel = "asdsad"
    grafica.Refresh

    Me.lbltotal = "Carga Total: " & funciones.FormatearDecimales(c) & " horas"

End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, False, False
    Me.GridEX1.ItemCount = 0

    ''Me.caption = caption & " (" & Name & ")"



End Sub


Private Sub GridEX1_SelectionChange()
    Set dto = col.item(Me.GridEX1.rowIndex(Me.GridEX1.row))
    GraficoTorta
End Sub

Private Sub GridEX1_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set dto = col(rowIndex)
    Values(1) = dto.Sector.Sector
    Values(2) = funciones.FormatearDecimales(dto.Tiempo) & " hs."

    If avancesOT.count > 0 Then
        If funciones.BuscarEnColeccion(avancesOT, CStr(dto.Sector.id)) Then
            Set tmpSectorTiempo = avancesOT.item(CStr(dto.Sector.id))
        End If
    End If


    If IsSomething(tmpSectorTiempo) Then
        Values(3) = funciones.FormatearDecimales(CDbl(tmpSectorTiempo.Tiempo)) & " hs."
    Else
        Values(3) = 0
    End If
End Sub


