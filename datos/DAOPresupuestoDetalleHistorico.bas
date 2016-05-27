Attribute VB_Name = "DAOPresupuestoDetalleHistorico"
Option Explicit
Dim historicoMDO As PresupuestoDetalleHistoricoMDO
Dim HistoricoMAT As PresupuestoDetalleHistoricoMAT

Public Type DatosMaterialDTO
    DimensionMaterial As String
    DimensionPieza As String
    costo As Double
    Kg As Double
    m2 As Double
End Type

Private Function histMDO(MDO As DesarrolloManoObra) As PresupuestoDetalleHistoricoMDO
    Set historicoMDO = New PresupuestoDetalleHistoricoMDO
    historicoMDO.CantOperarios = MDO.Cantidad
    Set historicoMDO.Tarea = MDO.Tarea

    historicoMDO.Tiempo = MDO.Tiempo
    historicoMDO.Valor = MDO.Tarea.CategoriaSueldo.Valor
    Set histMDO = historicoMDO
End Function

Private Function histMAT(MAT As DesarrolloMaterial) As PresupuestoDetalleHistoricoMAT
    Set HistoricoMAT = New PresupuestoDetalleHistoricoMAT
    HistoricoMAT.Ancho = MAT.Ancho
    HistoricoMAT.AnchoPieza = MAT.AnchoTerm
    HistoricoMAT.Cantidad = MAT.Cantidad
    HistoricoMAT.Largo = MAT.Largo
    HistoricoMAT.LargoPieza = MAT.LargoTerm
    Set HistoricoMAT.Material = MAT.Material
    HistoricoMAT.Scrap = MAT.Scrap
    HistoricoMAT.Valor = MAT.Material.Valor    '1* (1 + (MAT.Scrap / 100)) * MAT.Cantidad
    Set HistoricoMAT.Moneda = MAT.Material.Moneda
    Set histMAT = HistoricoMAT
End Function



Private Function CrearHistorico(Pieza As Pieza, deta As clsPresupuestoDetalle) As clsPresupuestoDetalleHistorico
    Dim MDO As DesarrolloManoObra
    Dim MAT As DesarrolloMaterial
    Dim historico As New clsPresupuestoDetalleHistorico
    historico.NombrePieza = Pieza.nombre
    Set historico.DetallePresupuesto = deta
    Set historico.Pieza = Pieza
    historico.FEcha = Now
    For Each MDO In Pieza.desarrollosManoObra
        historico.historicoMDO.Add histMDO(MDO)
    Next
    For Each MAT In Pieza.DesarrollosMaterial
        historico.HistoricoMAT.Add histMAT(MAT)
    Next
    Dim P As Pieza
    For Each P In Pieza.PiezasHijas
        historico.HistoricoHijos.Add CrearHistorico(P, deta)
    Next P
    Set CrearHistorico = historico
End Function
Private Function Guardar(historico As Collection, historicoPadreId As String) As Boolean
    On Error GoTo err1
    Guardar = True

    Dim strsql As String
    Dim ultimo_id As Long
    Dim h As clsPresupuestoDetalleHistorico
    Dim h2 As clsPresupuestoDetalleHistorico
    Dim MAT As PresupuestoDetalleHistoricoMAT
    Dim mo As PresupuestoDetalleHistoricoMDO
    For Each h In historico

        strsql = "  INSERT INTO detalle_presupuesto_historico" _
                 & "(nombre_pieza, pieza_id, fecha, id_detalle_presupuesto, id_detalle_presupuesto_historico_padre) Values" _
                 & "(" & conectar.Escape(h.Pieza.nombre) & "," & h.Pieza.id & " , " & conectar.Escape(h.FEcha) & " , " & h.DetallePresupuesto.id & ", " & historicoPadreId & ")"

        If Not conectar.execute(strsql) Then GoTo err1
        conectar.UltimoId "detalle_presupuesto_historico", ultimo_id

        For Each MAT In h.HistoricoMAT
            strsql = "INSERT INTO detalle_presupuesto_historico_mat  (material_id,  largo,   ancho,  largo_pieza, " _
                     & " ancho_pieza, Scrap,    cantidad, valor, id_detalle_presupuesto_historico, id_moneda) Values " _
                     & "( " & MAT.Material.id & " ," & MAT.Largo & "," & MAT.Ancho & "," & MAT.LargoPieza & "," & MAT.AnchoPieza & " , " & MAT.Scrap & " , " & MAT.Cantidad & ", " & MAT.Valor & " , " & ultimo_id & "," & MAT.Moneda.id & "   )"

            If Not conectar.execute(strsql) Then GoTo err1
        Next MAT

        For Each mo In h.historicoMDO
            strsql = "INSERT INTO detalle_presupuesto_historico_mdo" _
                     & " (tarea_id, valor, tiempo, cantidad, id_detalle_presupuesto_historico)" _
                     & " VALUES (" _
                     & mo.Tarea.id & "," & conectar.Escape(mo.Valor) & "," & conectar.Escape(mo.Tiempo) & "," & mo.CantOperarios & "," & ultimo_id & ")"
            If Not conectar.execute(strsql) Then GoTo err1
        Next mo


        If Not Guardar(h.HistoricoHijos, CStr(ultimo_id)) Then GoTo err1
    Next h

    Exit Function
err1:
    Guardar = False
End Function

Public Function Create(T As clsPresupuesto, Optional Save As Boolean = True) As Boolean
    On Error GoTo err1
    Create = False
    Dim deta As clsPresupuestoDetalle
    Dim col As New Collection
    For Each deta In DAOPresupuestosDetalle.GetAllByPresupuesto(T)
        col.Add CrearHistorico(DAOPieza.FindById(deta.Pieza.id, FL_4, True, True), deta)
    Next
    If Not col Is Nothing Then If Not Guardar(col, "NULL") Then GoTo err1

    Create = True
    Exit Function
err1:
    Create = False
End Function


