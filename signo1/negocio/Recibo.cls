VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Recibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Id As Long
Public Cliente As clsCliente
Public fechaCreacion As Date
Public fechaModificacion As Date
Public FechaAprobacion As Date
Public usuarioCreador As clsUsuario
Public usuarioAprobador As clsUsuario
Public estado As EstadoRecibo
Public redondeo As Double
'Public PagoACuenta As Boolean NO VA MAS
Public moneda As clsMoneda
Public TipoCambio As Currency

Public FEcha As Date

Public facturas As New Collection

'valores
Public Cheques As New Collection
Public OperacionesCaja As New Collection
Public operacionesBanco As New Collection

Public retenciones As New Collection

Public aCuenta As Double
Public aCuentaUsado As Double



Private m_validationMessages As String

Public pagosDeFacturas As New Dictionary    'tiene id factura y el monto pagado de la misma


Public TotalEstatico As New TotalEstaticoRecibo

Public Property Get ACuentaDisponible()
    If Me.estado = Reciboanulado Then
    ACuentaDisponible = 0
    Else
        ACuentaDisponible = aCuenta - aCuentaUsado
    End If

End Property


Public Property Get TotalFacturas() As Double

    Dim F As Factura
    Dim tot As Double
    Dim negativo As Integer
    Dim importe As Double

    For Each F In Me.facturas

        If F.TipoDocumento = tipoDocumentoContable.NotaCredito Then
            negativo = -1
        Else
            negativo = 1
        End If

        If Me.pagosDeFacturas.Exists(CStr(F.Id)) Then
            importe = Me.pagosDeFacturas.item(CStr(F.Id))
        Else
            importe = F.total
        End If

        'Si factura y recibo tienen la misma moneda,
        'no corresponde aplicar ninguna conversion.
        If F.moneda.Id = Me.moneda.Id Then

            tot = tot + (importe * negativo)

        Else

            'Se conserva provisionalmente la conversion
            'anterior para operaciones en distintas monedas.
            tot = tot + _
                  (importe * negativo * F.CambioAPatron)

        End If

    Next F

    TotalFacturas = tot

End Property

Public Property Get TotalRetenciones() As Double
    Dim r As retencionRecibo
    Dim tot As Double
    For Each r In Me.retenciones
        tot = tot + MonedaConverter.Convertir(r.valor, 0, Me.moneda.Id)
    Next r
    TotalRetenciones = tot
End Property

Public Property Get TotalCheques() As Double
    Dim che As cheque
    Dim tot As Double

    For Each che In Me.Cheques
        If IsSomething(che.moneda) And IsSomething(Me.moneda) Then
            tot = tot + MonedaConverter.Convertir(che.Monto, che.moneda.Id, Me.moneda.Id)
        End If
    Next che
    TotalCheques = tot
End Property

Public Property Get TotalOperacionesCaja() As Double
    Dim op As operacion
    Dim tot As Double

    For Each op In Me.OperacionesCaja
        tot = tot + MonedaConverter.Convertir(op.Monto, op.moneda.Id, Me.moneda.Id)
    Next op
    TotalOperacionesCaja = tot
End Property
Public Property Get TotalOperacionesBanco() As Double
    Dim op As operacion
    Dim tot As Double

    For Each op In Me.operacionesBanco
        tot = tot + MonedaConverter.Convertir(op.Monto, op.moneda.Id, Me.moneda.Id)
    Next op
    TotalOperacionesBanco = tot
End Property
Public Property Get total() As Double
    total = funciones.RedondearDecimales(Me.TotalFacturas - Me.TotalRetenciones, 2)
End Property


Public Property Get TotalRecibido() As Double

    Dim importe As Double

    importe = Me.TotalCheques + _
              Me.TotalOperacionesBanco + _
              Me.TotalOperacionesCaja

    importe = importe + _
              MonedaConverter.Convertir( _
                  Me.redondeo, _
                  DAOMoneda.MONEDA_PESO_ID, _
                  Me.moneda.Id)

    TotalRecibido = _
        funciones.RedondearDecimales(importe, 2)

End Property


Private Sub AddValidationMessage(msg As String)
    If LenB(m_validationMessages) > 0 Then m_validationMessages = m_validationMessages & vbNewLine
    m_validationMessages = m_validationMessages & msg
End Sub
Public Property Get IsValid() As Boolean
    m_validationMessages = vbNullString
    Dim F As Factura
    Dim facturas_id As New Dictionary
    For Each F In Me.facturas
        If facturas_id.Exists(CStr(F.GetShortDescription(False, True))) Then
            AddValidationMessage "Alguna de las facturas aparece mas de 1 vez en el recibo."
            Exit For
        End If
        facturas_id.Add CStr(F.GetShortDescription(False, True)), F.GetShortDescription(False, True)
    Next F

    If Me.TotalRecibido < Me.total Then    ' And Not Me.PagoACuenta Then
        AddValidationMessage "Los valores recibidos no igualan o superan al total del recibo."
    End If

    If Not Me.moneda Is Nothing Then
        If Me.moneda.Id = 1 Or Me.moneda.Id = 3 Then
            If Me.TipoCambio <= 0 Then
                AddValidationMessage _
                    "Debe ingresar un tipo de cambio mayor a cero."
            End If
        End If
    End If

    IsValid = (LenB(m_validationMessages) = 0)
    
End Property

Public Property Get ValidationMessages() As String
    ValidationMessages = m_validationMessages
End Property


