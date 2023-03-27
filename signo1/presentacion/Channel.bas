Attribute VB_Name = "Channel"
Option Explicit
Option Base 1
Const MAXIMO = 6

Private lista(MAXIMO, 1) As Variant
Private tmp As ISuscriber




Private frm As New Collection
Public Enum TipoSuscripcion
    Presupuestos_ = 1
    RequerimientosCompra_ = 2
    Proveedores_ = 3
    ordenesTrabajo = 4
    NuevoPresupuesto_ = 5
    Clientes_ = 6
    NuevaOT_ = 7
    Tareas_ = 8
    EdicionPieza_ = 9
    EdicionDetallePeticionOferta = 10
    Remitos_ = 11
    PasajeChequePropioCartera = 12
    FacturaProveedor_ = 13
    OrdenesPago_ = 14
    RemitosDetalle_ = 15
    FacturaCliente_ = 16
    Materiales_ = 17
    RubrosGrupos_ = 18
    Documentos_ = 20
    TS_InformeAccidente = 19
    TS_Siniestro = 20
    FacturarRemitosDetalle_ = 21
    EnvioMail_ = 22
    Recibos_ = 23

    LiquidacionCaja_ = 24
End Enum
Private suscriptores As New Dictionary
Public Function AgregarSuscriptor(value As ISuscriber, Tipo As TipoSuscripcion, Optional privateBroadcast As Boolean = False) As Boolean
    Dim tmpCol As Collection
    AgregarSuscriptor = True

    If CStr(value.Id) = vbNullString Then GoTo err1

    If suscriptores.Exists(CStr(value.Id)) Then    'ya esta el suscriptor, le cambio la suscripcion que me pide, o la agrego
        Set tmpCol = suscriptores.item(CStr(value.Id))
        If Not funciones.BuscarEnColeccion(tmpCol, CStr(Tipo)) Then tmpCol.Add Tipo, CStr(Tipo)
    Else    'lo agrego
        Set tmpCol = New Collection
        tmpCol.Add value, "-"     ' el suscriptor va a tener la calve - en la coleccion
        tmpCol.Add Tipo, CStr(Tipo)
        tmpCol.Add privateBroadcast, "privateBroadcast"    'esto aca esta mal, porque tendria que ser por tipo no en general

        suscriptores.Add CStr(value.Id), tmpCol
    End If
    Exit Function
err1:
    AgregarSuscriptor = False
End Function


Public Function RemoverSuscriptorParcial(value As ISuscriber, Tipo As TipoSuscripcion)
    Dim tmp As Collection
    If suscriptores.Exists(CStr(value.Id)) Then
        Set tmp = suscriptores.item(CStr(value.Id))
        If funciones.BuscarEnColeccion(tmp, CStr(Tipo)) Then
            tmp.remove CStr(Tipo)
        End If
    End If
End Function
Public Function RemoverSuscripcionTotal(value As ISuscriber)
    If suscriptores.Exists(CStr(value.Id)) Then suscriptores.remove (CStr(value.Id))
End Function


Public Function Notificar(nvalue As clsEventoObserver, Optional Tipo As TipoSuscripcion)

    Dim tmp As ISuscriber
    Dim suscripts As Variant
    Dim i As Long
    Dim kk As Object
    On Error GoTo E

    'chequeo si hay alguno privado
    '    Dim hayPrivado As Boolean: hayPrivado = False
    '    For Each suscripts In suscriptores.Items
    '        If suscripts.Item("privateBroadcast") Then
    '            hayPrivado = True
    '            Set tmp = suscripts.Item("-")
    '            For i = 2 To suscripts.count - 1
    '                If suscripts(i) = tipo Then    'verifico el tipo de suscripcion
    '                    nValue.tipo = tipo 'en realidad, tipo no tendria que llegar como param, tendria ya uqe llegar asignado a nvalue
    '                    tmp.Notificarse nValue
    '                End If
    '            Next i
    '        End If
    '    Next
    '
    'If Not hayPrivado Then
    For Each suscripts In suscriptores.Items
        Set tmp = suscripts.item("-")         'tengo el suscriptor

        For i = 2 To suscripts.count - 1     ' a lo ultimo esta el private
            If suscripts(i) = Tipo Then        'verifico el tipo de suscripcion
                nvalue.Tipo = Tipo     'en realidad, tipo no tendria que llegar como param, tendria ya uqe llegar asignado a nvalue
                tmp.Notificarse nvalue
            End If
        Next i
    Next
    'End If
    Exit Function
E:
    MsgBox "Error en Channel.Notificar()" & vbNewLine & Err.Description, vbCritical
End Function







