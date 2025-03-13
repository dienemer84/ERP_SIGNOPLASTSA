VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrdenPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Id As Long
Public moneda As clsMoneda
Public FEcha As Date
Public FacturasProveedor As New Collection
Public DeudaCompensatorios As New Collection
Public CuentaContable As clsCuentaContable
Public CuentaContableDescripcion As String
Public ChequesTerceros As New Collection
Public ChequesPropios As New Collection
Public operacionesCaja As New Collection
Public operacionesBanco As New Collection
Public Compensatorios As New Collection
Public estado As EstadoOrdenPago
Public CertificadoRetencion As New CertificadoRetencion
Public alicuota As Double
Private m_validationMessages As String
Public staticTipoCambioHoy As Double
Public DiferenciaCambio As Double
Public OtrosDescuentos As Double
Public RetencionesAlicuota As New Collection

Public IdMonedaPago As Long
Public TipoCambio As Double
Public StaticTotalFacturas As Double
Public StaticTotalFacturasNG As Double
Public StaticTotalRetenido As Double
Public StaticTotalOrigenes As Double
Public DiferenciaCambioEnNG As Double
Public DiferenciaCambioEnTOTAL As Double
Public staticTotalDeudaCompensatorios As Double

Public PagosACuenta As New Collection

Public Property Get TotalNGCompensatorios() As Double
    Dim T As Double
    Dim C As New Compensatorio

    For Each C In Compensatorios
        T = T + C.NetoGravadoCompensado
    Next C
    TotalNGCompensatorios = T
End Property

Public Property Get TotalCompensatorios() As Double
    Dim C As New Compensatorio
    Dim T As Double
    Dim i As Integer
    For Each C In Compensatorios
        If C.Tipo = TC_Credito Then i = -1 Else i = 1
        T = T + (C.Monto * i)
    Next C
    TotalCompensatorios = T
End Property

Public Property Get TotalDeudaCompensatorios() As Double
    Dim C As New Compensatorio
    Dim T As Double
    Dim i As Integer
    For Each C In DeudaCompensatorios
        If C.Tipo = TC_Credito Then i = -1 Else i = 1
        T = T + (C.Monto * i)
    Next C
    TotalDeudaCompensatorios = T
End Property

Public Property Get TotalFacturas() As Double
    Dim F As clsFacturaProveedor
    Dim SUM As Double
    For Each F In Me.FacturasProveedor
        SUM = SUM + MonedaConverter.Convertir(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.total * -1, F.total), F.moneda.Id, Me.moneda.Id)
    Next F
    TotalFacturas = funciones.RedondearDecimales(SUM)
End Property

Public Property Get TotalFacturasNG() As Double
    Dim F As clsFacturaProveedor
    Dim SUM As Double
    For Each F In Me.FacturasProveedor
        SUM = SUM + MonedaConverter.Convertir(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.NetoGravado * -1, F.NetoGravado), F.moneda.Id, Me.moneda.Id)
    Next F
    TotalFacturasNG = funciones.RedondearDecimales(SUM)
End Property

Public Property Get TotalOrigenes() As Double

    If Me.moneda Is Nothing Then
        TotalOrigenes = 0
        Exit Property
    End If

    Dim total As Double
    Dim cheq As cheque
    For Each cheq In Me.ChequesTerceros
        If IsSomething(cheq.moneda) Then
            total = total + MonedaConverter.Convertir(cheq.Monto, cheq.moneda.Id, Me.moneda.Id)
        End If
    Next cheq

    For Each cheq In Me.ChequesPropios
        If IsSomething(cheq.moneda) Then
            '  Debug.Print cheq.numero, cheq.Monto


            total = total + MonedaConverter.Convertir(cheq.Monto, cheq.moneda.Id, Me.moneda.Id)
        End If
    Next cheq

    Dim op As operacion
    For Each op In Me.operacionesBanco
        If IsSomething(op.moneda) Then
            total = total + MonedaConverter.Convertir(op.Monto, op.moneda.Id, Me.moneda.Id)
        End If
    Next op

    For Each op In Me.operacionesCaja
        If IsSomething(op.moneda) Then
            total = total + MonedaConverter.Convertir(op.Monto, op.moneda.Id, Me.moneda.Id)
        End If
    Next op
    '13/08/12 corregi con sabrina esta funcion para que cierre bien la Op. A pedido de karin. chequear.

    ' TotalOrigenes = funciones.RedondearDecimales(Total + Me.OtrosDescuentos + Me.DiferenciaCambio)

    '15/10/12 corregi con sabrina esta funcion para que cierre bien la Op.  reemplazo funcion anterior
    TotalOrigenes = funciones.RedondearDecimales(total + Me.OtrosDescuentos + 0)
End Property


'Public Property Get TotalOrdenPago() As Double
''totFact - totRet - OrdenPago.DiferenciaCambio - OrdenPago.OtrosDescuentos
'
'TotalOrdenPago = TotalFacturas - Me.StaticTotalRetenido - DiferenciaCambio - OtrosDescuentos + TotalCompensatorios
'
'
'End Property

Public Property Get TotalOrigenesDiscriminado() As Collection

    If Me.moneda Is Nothing Then
        Set TotalOrigenesDiscriminado = Nothing
        Exit Property
    End If

    Dim colRetorno As New Collection
    Dim itemRetorno As Collection

    Dim total As Double
    Dim cheq As cheque
    For Each cheq In Me.ChequesTerceros
        If IsSomething(cheq.moneda) Then
            Set itemRetorno = New Collection
            itemRetorno.Add cheq.FechaVencimiento
            itemRetorno.Add MonedaConverter.Convertir(cheq.Monto, cheq.moneda.Id, Me.moneda.Id)
            colRetorno.Add itemRetorno
        End If
    Next cheq

    For Each cheq In Me.ChequesPropios
        If IsSomething(cheq.moneda) Then
            Set itemRetorno = New Collection
            itemRetorno.Add cheq.FechaVencimiento
            itemRetorno.Add MonedaConverter.Convertir(cheq.Monto, cheq.moneda.Id, Me.moneda.Id)
            colRetorno.Add itemRetorno
        End If
    Next cheq

    Dim op As operacion
    For Each op In Me.operacionesBanco
        If IsSomething(op.moneda) Then
            Set itemRetorno = New Collection
            itemRetorno.Add op.FechaOperacion
            itemRetorno.Add MonedaConverter.Convertir(op.Monto, op.moneda.Id, Me.moneda.Id)
            colRetorno.Add itemRetorno
        End If
    Next op

    For Each op In Me.operacionesCaja
        If IsSomething(op.moneda) Then
            Set itemRetorno = New Collection
            itemRetorno.Add op.FechaOperacion
            itemRetorno.Add MonedaConverter.Convertir(op.Monto, op.moneda.Id, Me.moneda.Id)
            colRetorno.Add itemRetorno
        End If
    Next op

    Set TotalOrigenesDiscriminado = colRetorno
End Property

Private Sub Class_Initialize()
    Me.FEcha = Now
End Sub

Private Sub AddValidationMessage(msg As String)
    If LenB(m_validationMessages) > 0 Then m_validationMessages = m_validationMessages & vbNewLine
    m_validationMessages = m_validationMessages & msg
End Sub


Public Property Get IsValid() As Boolean
    m_validationMessages = vbNullString

    Dim C As cheque
    Dim cheques_id As New Dictionary
    'busco cheques duplicadas
    For Each C In Me.ChequesTerceros
        If cheques_id.Exists(CStr(C.Id)) Then
            AddValidationMessage "Alguno de los cheques de 3ros aparece mas de 1 vez."
            Exit For
        End If
        cheques_id.Add CStr(C.Id), C.Id
    Next C

    Set cheques_id = New Dictionary
    For Each C In Me.ChequesPropios
        If cheques_id.Exists(CStr(C.Id)) Then
            AddValidationMessage "Alguno de los cheques propios aparece mas de 1 vez."
            Exit For
        End If
        cheques_id.Add CStr(C.Id), C.Id
    Next C


    If Me.TotalOrigenes = 0 Then
        AddValidationMessage "El total de la orden debe ser mayor a cero."
    End If

    If Me.CuentaContable Is Nothing And Me.FacturasProveedor.count = 0 Then
        AddValidationMessage "La orden de pago debe contener facturas o cuenta contable."
    End If

    If Me.TotalOrigenes <> funciones.RedondearDecimales(Me.StaticTotal) Then

        AddValidationMessage "El total supuesto a pagar (" & Me.moneda.NombreCorto & " " & (Me.StaticTotal) & ") debe ser igual a lo que se va a pagar con valores (" & Me.moneda.NombreCorto & " " & Me.TotalOrigenes & ")."
    End If


    'validar si los compensatorios creados tienen su comprobante correlativo
    IsValid = (LenB(m_validationMessages) = 0)

End Property


Property Get StaticTotal() As Double
    StaticTotal = Me.StaticTotalFacturas - Me.StaticTotalRetenido - Me.DiferenciaCambio - Me.OtrosDescuentos + Me.TotalCompensatorios + Me.DiferenciaCambioEnTOTAL + Me.staticTotalDeudaCompensatorios
End Property


Public Property Get ValidationMessages() As String
    ValidationMessages = m_validationMessages
End Property


Public Property Get EsParaFacturaProveedor() As Boolean
    EsParaFacturaProveedor = False
    If IsSomething(Me.FacturasProveedor) Then
        EsParaFacturaProveedor = Me.FacturasProveedor.count > 0
    Else
        If IsSomething(Me.CuentaContable) Then
            EsParaFacturaProveedor = (Me.CuentaContable.Id = 0)
        End If
    End If
End Property

Public Property Get TotalOrdenPago() As Double
'If Id = 7180 Then Stop


'Debug.Print "totales", Me.TotalOrigenes + Me.StaticTotalRetenido, "totales staticos", Me.StaticTotalOrigenes + Me.StaticTotalRetenido

    TotalOrdenPago = Me.TotalOrigenes + Me.StaticTotalRetenido    '+ Me.TotalCompensatorios

End Property
