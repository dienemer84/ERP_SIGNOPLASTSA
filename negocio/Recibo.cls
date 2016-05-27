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

Public id As Long
Public Cliente As clsCliente
Public FechaCreacion As Date
Public FechaModificacion As Date
Public FechaAprobacion As Date
Public UsuarioCreador As clsUsuario
Public UsuarioAprobador As clsUsuario
Public estado As EstadoRecibo
Public Redondeo As Double
'Public PagoACuenta As Boolean NO VA MAS
Public Moneda As clsMoneda

Public FEcha As Date

Public facturas As New Collection

'valores
Public Cheques As New Collection
Public OperacionesCaja As New Collection
Public OperacionesBanco As New Collection

Public retenciones As New Collection

Public ACuenta As Double
Public ACuentaUsado As Double



Private m_validationMessages As String

Public PagosDeFacturas As New Dictionary    'tiene id factura y el monto pagado de la misma


Public TotalEstatico As New TotalEstaticoRecibo

Public Property Get ACuentaDisponible()
    ACuentaDisponible = ACuenta - ACuentaUsado
End Property

Public Property Get TotalFacturas() As Double
    Dim F As Factura
    Dim tot As Double
    Dim negativo As Integer

    For Each F In Me.facturas
        If F.TipoDocumento = tipoDocumentoContable.notaCredito Then
            negativo = -1
        Else
            negativo = 1
        End If

        If Me.PagosDeFacturas.Exists(CStr(F.id)) Then

            'tot = tot + MonedaConverter.Convertir(Me.PagosDeFacturas.Item(CStr(F.Id)) * negativo, F.Moneda.Id, Me.Moneda.Id)
            tot = tot + (Me.PagosDeFacturas.item(CStr(F.id)) * negativo * F.CambioAPatron)
        Else
            tot = tot + F.Total * negativo * F.CambioAPatron
        End If


    Next F
    TotalFacturas = tot
End Property

Public Property Get TotalRetenciones() As Double
    Dim r As retencionRecibo
    Dim tot As Double
    For Each r In Me.retenciones
        tot = tot + MonedaConverter.Convertir(r.Valor, 0, Me.Moneda.id)
    Next r
    TotalRetenciones = tot
End Property

Public Property Get TotalCheques() As Double
    Dim che As cheque
    Dim tot As Double

    For Each che In Me.Cheques
        If IsSomething(che.Moneda) And IsSomething(Me.Moneda) Then
            tot = tot + MonedaConverter.Convertir(che.Monto, che.Moneda.id, Me.Moneda.id)
        End If
    Next che
    TotalCheques = tot
End Property

Public Property Get TotalOperacionesCaja() As Double
    Dim op As operacion
    Dim tot As Double

    For Each op In Me.OperacionesCaja
        tot = tot + MonedaConverter.Convertir(op.Monto, op.Moneda.id, Me.Moneda.id)
    Next op
    TotalOperacionesCaja = tot
End Property
Public Property Get TotalOperacionesBanco() As Double
    Dim op As operacion
    Dim tot As Double

    For Each op In Me.OperacionesBanco
        tot = tot + MonedaConverter.Convertir(op.Monto, op.Moneda.id, Me.Moneda.id)
    Next op
    TotalOperacionesBanco = tot
End Property
Public Property Get Total() As Double
    Total = funciones.RedondearDecimales(Me.TotalFacturas - Me.TotalRetenciones, 2)
End Property
Public Property Get TotalRecibido() As Double
    TotalRecibido = funciones.RedondearDecimales(MonedaConverter.Convertir(Me.TotalCheques + Me.TotalOperacionesBanco + Me.TotalOperacionesCaja, Me.Moneda.id, DAOMoneda.MONEDA_PESO_ID) + MonedaConverter.Convertir(Me.Redondeo, DAOMoneda.MONEDA_PESO_ID, Me.Moneda.id), 2)
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

    If Me.TotalRecibido < Me.Total Then    ' And Not Me.PagoACuenta Then
        AddValidationMessage "Los valores recibidos no igualan o superan al total del recibo."
    End If

    IsValid = (LenB(m_validationMessages) = 0)
End Property

Public Property Get ValidationMessages() As String
    ValidationMessages = m_validationMessages
End Property


