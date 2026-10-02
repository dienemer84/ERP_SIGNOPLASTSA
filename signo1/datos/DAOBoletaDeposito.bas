Attribute VB_Name = "DAOBoletaDeposito"
Option Explicit

Public UltimoError As String


Public Function Save(ByVal boleta As BoletaDeposito) As Boolean

    On Error GoTo err1

    Dim chequeSeleccionado As cheque
    Dim chequeActual As cheque

    Dim chequesValidados As New Collection

    Dim montoTotal As Double
    Dim idBoleta As Long
    Dim q As String


    Save = False
    UltimoError = vbNullString


    '-------------------------------------------------------
    ' VALIDACIONES GENERALES
    '-------------------------------------------------------

    If boleta Is Nothing Then
        UltimoError = "No se recibió la boleta de depósito."
        Exit Function
    End If


    If boleta.CuentaDestino Is Nothing Then
        UltimoError = "No se indicó la cuenta bancaria destino."
        Exit Function
    End If


    If boleta.Cheques.count = 0 Then
        UltimoError = "La boleta no contiene cheques."
        Exit Function
    End If


    If boleta.numero <= 0 Then
        UltimoError = "El número de boleta no es válido."
        Exit Function
    End If


    If boleta.CuentaDestino.moneda Is Nothing Then
        UltimoError = "La cuenta bancaria seleccionada no tiene moneda definida."
        Exit Function
    End If

    '-------------------------------------------------------
    ' VALIDAR CONCILIACION BANCARIA
    '-------------------------------------------------------
    
    If Not PuedeGuardarDeposito( _
                boleta.CuentaDestino, _
                boleta.fechaDeposito) Then
    
        Exit Function
    
    End If

    '-------------------------------------------------------
    ' INICIAR TRANSACCION
    '-------------------------------------------------------

    conectar.BeginTransaction


    '-------------------------------------------------------
    ' VOLVER A LEER Y VALIDAR LOS CHEQUES
    ' A LA VEZ CALCULAMOS EL TOTAL REAL
    '-------------------------------------------------------

    montoTotal = 0


    For Each chequeSeleccionado In boleta.Cheques

        Set chequeActual = DAOCheques.FindById( _
                                chequeSeleccionado.Id)


        If chequeActual Is Nothing Then

            Err.Raise vbObjectError + 3201, _
                      "DAOBoletaDeposito.Save", _
                      "No se encontró uno de los cheques seleccionados."

        End If


        If Not chequeActual.EnCartera Then

            Err.Raise vbObjectError + 3202, _
                      "DAOBoletaDeposito.Save", _
                      "El cheque Nº " & chequeActual.numero & _
                      " ya no se encuentra en cartera."

        End If


        If chequeActual.Depositado Then

            Err.Raise vbObjectError + 3203, _
                      "DAOBoletaDeposito.Save", _
                      "El cheque Nº " & chequeActual.numero & _
                      " ya figura como depositado."

        End If


        If chequeActual.moneda Is Nothing Then

            Err.Raise vbObjectError + 3204, _
                      "DAOBoletaDeposito.Save", _
                      "El cheque Nº " & chequeActual.numero & _
                      " no tiene moneda definida."

        End If


        If chequeActual.moneda.Id <> _
           boleta.CuentaDestino.moneda.Id Then

            Err.Raise vbObjectError + 3205, _
                      "DAOBoletaDeposito.Save", _
                      "La moneda del cheque Nº " & _
                      chequeActual.numero & _
                      " no coincide con la moneda de la cuenta bancaria."

        End If


        montoTotal = montoTotal + chequeActual.Monto


        chequesValidados.Add _
            chequeActual, _
            CStr(chequeActual.Id)

    Next chequeSeleccionado


    '-------------------------------------------------------
    ' GUARDAR CABECERA DE LA BOLETA
    '-------------------------------------------------------

    q = "INSERT INTO boleta_deposito " & _
        "(monto, fecha_deposito, tipo_deposito, " & _
        "numero_boleta, id_cuenta) VALUES (" & _
        conectar.Escape(montoTotal) & ", " & _
        conectar.Escape(boleta.fechaDeposito) & ", " & _
        conectar.Escape(boleta.TipoDeposito) & ", " & _
        conectar.Escape(boleta.numero) & ", " & _
        conectar.Escape(boleta.CuentaDestino.Id) & ")"


    If Not conectar.execute(q) Then

        Err.Raise vbObjectError + 3206, _
                  "DAOBoletaDeposito.Save", _
                  "No se pudo guardar la cabecera de la boleta."

    End If


    idBoleta = conectar.UltimoId2()


    If idBoleta <= 0 Then

        Err.Raise vbObjectError + 3207, _
                  "DAOBoletaDeposito.Save", _
                  "No se pudo obtener el identificador de la boleta."

    End If


    boleta.Id = idBoleta
    boleta.Monto = montoTotal


    '-------------------------------------------------------
    ' DEPOSITAR LOS CHEQUES
    '-------------------------------------------------------

    For Each chequeActual In chequesValidados


        If Not DepositarChequeInterno( _
                    chequeActual, _
                    boleta.CuentaDestino, _
                    boleta.fechaDeposito, _
                    CStr(boleta.numero), _
                    idBoleta) Then


            Err.Raise vbObjectError + 3208, _
                      "DAOBoletaDeposito.Save", _
                      "No se pudo registrar el depósito del cheque Nº " & _
                      chequeActual.numero

        End If


    Next chequeActual


    '-------------------------------------------------------
    ' TODO OK
    '-------------------------------------------------------

    conectar.CommitTransaction


    Save = True
    Exit Function



err1:

    UltimoError = Err.Description


    If LenB(UltimoError) = 0 Then

        UltimoError = _
            "Se produjo un error al guardar la boleta de depósito."

    End If


    conectar.RollBackTransaction


    Save = False

End Function



Private Function DepositarChequeInterno( _
            ByVal cheque As cheque, _
            ByVal cuenta As CuentaBancaria, _
            ByVal fechaDeposito As Date, _
            ByVal Comprobante As String, _
            ByVal idBoleta As Long) As Boolean


    Dim op As operacion
    Dim q As String
    Dim valorIdBoleta As String


    DepositarChequeInterno = False


    '-------------------------------------------------------
    ' CREAR MOVIMIENTO BANCARIO
    '-------------------------------------------------------

    Set op = New operacion


    op.IdPertenencia = cheque.Id

    op.EntradaSalida = OPEntrada

    op.FechaOperacion = fechaDeposito

    op.Pertenencia = Banco

    Set op.moneda = cheque.moneda

    op.Monto = cheque.Monto

    Set op.CuentaBancaria = cuenta

    op.Comprobante = Comprobante


    If Not DAOOperacion.Save(op) Then
        Exit Function
    End If


    op.Id = conectar.UltimoId2()


    If op.Id <= 0 Then
        Exit Function
    End If


    '-------------------------------------------------------
    ' MARCAR CHEQUE COMO DEPOSITADO
    '-------------------------------------------------------

    cheque.Depositado = True
    cheque.EnCartera = False


    If Not DAOCheques.Guardar(cheque) Then
        Exit Function
    End If


    '-------------------------------------------------------
    ' RELACIONAR:
    '
    ' BOLETA
    '   |
    '   +-- CHEQUE
    '   |
    '   +-- OPERACION
    '-------------------------------------------------------

    If idBoleta > 0 Then

        valorIdBoleta = CStr(idBoleta)

    Else

        valorIdBoleta = "NULL"

    End If


    q = "INSERT INTO cheques_depositos " & _
        "(id_boleta, id_cheque, id_operacion) VALUES (" & _
        valorIdBoleta & ", " & _
        cheque.Id & ", " & _
        op.Id & ")"


    If Not conectar.execute(q) Then
        Exit Function
    End If


    DepositarChequeInterno = True

End Function



'==================================================================
' COMPATIBILIDAD CON CODIGO VIEJO
'
' No se elimina por si existiera algún llamado antiguo.
' Al no existir una BoletaDeposito completa, id_boleta queda NULL.
'==================================================================

Public Function Depositar( _
            cheque As cheque, _
            cuenta As CuentaBancaria, _
            FEcha As Date) As Boolean


    On Error GoTo err1


    UltimoError = vbNullString
    Depositar = False


    If cheque Is Nothing Then

        UltimoError = "No se recibió el cheque."
        Exit Function

    End If


    If cuenta Is Nothing Then

        UltimoError = "No se recibió la cuenta bancaria."
        Exit Function

    End If

    If Not PuedeGuardarDeposito( _
                cuenta, _
                FEcha) Then
    
        Exit Function
    
    End If

    conectar.BeginTransaction


    If Not DepositarChequeInterno( _
                cheque, _
                cuenta, _
                FEcha, _
                "-", _
                0) Then


        Err.Raise vbObjectError + 3210, _
                  "DAOBoletaDeposito.Depositar", _
                  "No se pudo efectuar el depósito."

    End If


    conectar.CommitTransaction


    Depositar = True
    Exit Function



err1:

    UltimoError = Err.Description


    conectar.RollBackTransaction


    Depositar = False

End Function


Public Function FindAll( _
        Optional ByVal FechaDesde As Variant, _
        Optional ByVal FechaHasta As Variant, _
        Optional ByVal idBoleta As Long = 0, _
        Optional ByVal numeroBoleta As Long = 0, _
        Optional ByVal idCuenta As Long = 0) As Collection

    On Error GoTo err1

    Dim q As String
    Dim rs As ADODB.Recordset
    Dim col As New Collection

    Dim tmp As BoletaDeposito
    Dim idCuentaTmp As Long

    UltimoError = vbNullString


    q = "SELECT " & _
        "b.id, " & _
        "b.monto, " & _
        "b.fecha_deposito, " & _
        "b.tipo_deposito, " & _
        "b.numero_boleta, " & _
        "b.id_cuenta, " & _
        "COUNT(cd.id) AS cantidad_cheques " & _
        "FROM boleta_deposito b " & _
        "LEFT JOIN cheques_depositos cd " & _
        "ON cd.id_boleta = b.id " & _
        "WHERE b.numero_boleta IS NOT NULL "


    '---------------------------------------------------
    ' FILTRO POR ID
    '---------------------------------------------------
    
    If idBoleta > 0 Then
    
        q = q & " AND b.id = " & idBoleta
    
    End If

    '---------------------------------------------------
    ' FECHA DESDE
    '---------------------------------------------------

    If Not IsEmpty(FechaDesde) And idBoleta = 0 Then

        If Not IsNull(FechaDesde) Then

            q = q & _
                " AND b.fecha_deposito >= " & _
                conectar.Escape(CDate(FechaDesde))

        End If

    End If


    '---------------------------------------------------
    ' FECHA HASTA
    '---------------------------------------------------

    If Not IsEmpty(FechaHasta) And idBoleta = 0 Then

        If Not IsNull(FechaHasta) Then

            q = q & _
                " AND b.fecha_deposito <= " & _
                conectar.Escape(CDate(FechaHasta))

        End If

    End If


    '---------------------------------------------------
    ' NUMERO DE BOLETA
    '---------------------------------------------------

    If numeroBoleta > 0 Then

        q = q & _
            " AND b.numero_boleta = " & numeroBoleta

    End If


    '---------------------------------------------------
    ' CUENTA
    '---------------------------------------------------

    If idCuenta > 0 Then

        q = q & _
            " AND b.id_cuenta = " & idCuenta

    End If


    '---------------------------------------------------
    ' AGRUPAR
    '---------------------------------------------------

    q = q & _
        " GROUP BY " & _
        "b.id, " & _
        "b.monto, " & _
        "b.fecha_deposito, " & _
        "b.tipo_deposito, " & _
        "b.numero_boleta, " & _
        "b.id_cuenta " & _
        "ORDER BY b.fecha_deposito DESC, b.id DESC"


    Set rs = conectar.RSFactory(q)


    '---------------------------------------------------
    ' MAPEAR
    '---------------------------------------------------

    While Not rs.EOF

        Set tmp = New BoletaDeposito


        tmp.Id = CLng(rs!Id)


        If Not IsNull(rs!numero_boleta) Then
            tmp.numero = CLng(rs!numero_boleta)
        End If


        If Not IsNull(rs!Monto) Then
            tmp.Monto = CDbl(rs!Monto)
        End If


        If Not IsNull(rs!fecha_deposito) Then
            tmp.fechaDeposito = CDate(rs!fecha_deposito)
        End If


        If Not IsNull(rs!tipo_deposito) Then
            tmp.TipoDeposito = CLng(rs!tipo_deposito)
        End If


        If Not IsNull(rs!cantidad_cheques) Then
            tmp.CantidadCheques = CLng(rs!cantidad_cheques)
        End If


        '-----------------------------------------------
        ' CUENTA BANCARIA
        '-----------------------------------------------

        If Not IsNull(rs!id_cuenta) Then

            idCuentaTmp = CLng(rs!id_cuenta)

            If idCuentaTmp > 0 Then

                Set tmp.CuentaDestino = _
                    DAOCuentaBancaria.FindById(idCuentaTmp)

            End If

        End If


        col.Add tmp, CStr(tmp.Id)


        rs.MoveNext

    Wend


    Set FindAll = col

    Exit Function


err1:

    UltimoError = Err.Description

    Set FindAll = Nothing

End Function


Public Function FindChequesByBoleta( _
            ByVal idBoleta As Long) As Collection

    On Error GoTo err1

    Dim q As String
    Dim rs As ADODB.Recordset

    Dim col As New Collection
    Dim ch As cheque


    UltimoError = vbNullString


    q = "SELECT id_cheque " & _
        "FROM cheques_depositos " & _
        "WHERE id_boleta = " & idBoleta & " " & _
        "ORDER BY id"


    Set rs = conectar.RSFactory(q)


    While Not rs.EOF

        Set ch = DAOCheques.FindById( _
                    CLng(rs!id_cheque))


        If Not ch Is Nothing Then

            col.Add ch, CStr(ch.Id)

        End If


        rs.MoveNext

    Wend


    Set FindChequesByBoleta = col

    Exit Function


err1:

    UltimoError = Err.Description

    Set FindChequesByBoleta = Nothing

End Function


Private Function PuedeGuardarDeposito( _
    ByVal cuenta As CuentaBancaria, _
    ByVal fechaDeposito As Date) As Boolean

    On Error GoTo err1

    Dim IdConciliacion As Long

    PuedeGuardarDeposito = False

    If cuenta Is Nothing Then Exit Function
    If cuenta.Id <= 0 Then Exit Function
    If CDbl(fechaDeposito) <= 0 Then Exit Function

    IdConciliacion = _
        DAOConciliacionBancaria.ObtenerIdConciliacionCerrada( _
            cuenta.Id, _
            fechaDeposito)

    If IdConciliacion > 0 Then

        UltimoError = _
            "No se puede registrar la boleta de depósito." & _
            vbCrLf & vbCrLf & _
            "Cuenta: " & cuenta.numero & vbCrLf & _
            "Fecha de depósito: " & _
            Format$(fechaDeposito, "dd/mm/yyyy") & _
            vbCrLf & vbCrLf & _
            "La cuenta se encuentra cerrada por la " & _
            "Conciliación Bancaria Nro " & _
            IdConciliacion & "."

        Exit Function

    End If

    PuedeGuardarDeposito = True
    Exit Function

err1:

    PuedeGuardarDeposito = False

    UltimoError = _
        "No se pudo verificar el período bancario." & _
        vbCrLf & _
        Err.Number & " - " & Err.Description

End Function


Public Function FindById( _
        ByVal idBoleta As Long) As BoletaDeposito

    On Error GoTo err1

    Dim rs As ADODB.Recordset
    Dim B As BoletaDeposito
    Dim q As String

    UltimoError = vbNullString

    If idBoleta <= 0 Then
        UltimoError = "El ID de la boleta no es válido."
        Exit Function
    End If

    q = "SELECT id, monto, fecha_deposito, " & _
        "tipo_deposito, numero_boleta, id_cuenta " & _
        "FROM boleta_deposito " & _
        "WHERE id = " & idBoleta & " LIMIT 1"

    Set rs = conectar.RSFactory(q)

    If rs.EOF Then
        UltimoError = "No se encontró la boleta."
        Exit Function
    End If

    If IsNull(rs!numero_boleta) Or _
       IsNull(rs!fecha_deposito) Or _
       IsNull(rs!tipo_deposito) Or _
       IsNull(rs!id_cuenta) Then

        UltimoError = "La boleta tiene datos incompletos."
        Exit Function

    End If

    Set B = New BoletaDeposito

    B.Id = CLng(rs!Id)
    B.numero = CLng(rs!numero_boleta)
    B.fechaDeposito = CDate(rs!fecha_deposito)
    B.TipoDeposito = CLng(rs!tipo_deposito)

    If Not IsNull(rs!Monto) Then
        B.Monto = CDbl(rs!Monto)
    End If

    Set B.CuentaDestino = _
        DAOCuentaBancaria.FindById(CLng(rs!id_cuenta))

    If B.CuentaDestino Is Nothing Then
        UltimoError = "No se encontró la cuenta bancaria."
        Exit Function
    End If

    Set FindById = B
    Exit Function

err1:

    UltimoError = Err.Description
    Set FindById = Nothing

End Function


Public Function PuedeEditarBoleta( _
    ByVal idBoleta As Long) As Boolean

    On Error GoTo err1

    Dim rs As ADODB.Recordset
    Dim q As String
    Dim idCuenta As Long
    Dim fechaOriginal As Date
    Dim IdConciliacion As Long

    PuedeEditarBoleta = False
    UltimoError = vbNullString

    If idBoleta <= 0 Then
        UltimoError = "El ID de la boleta no es válido."
        Exit Function
    End If

    'Obtener los datos actuales de la boleta.
    q = "SELECT id_cuenta, fecha_deposito " & _
        "FROM boleta_deposito " & _
        "WHERE id = " & idBoleta

    Set rs = conectar.RSFactory(q)

    If rs.EOF Then
        UltimoError = "La boleta ya no existe."
        Exit Function
    End If

    If IsNull(rs!id_cuenta) Or _
       IsNull(rs!fecha_deposito) Then

        UltimoError = "La boleta tiene datos incompletos."
        Exit Function

    End If

    idCuenta = CLng(rs!id_cuenta)
    fechaOriginal = CDate(rs!fecha_deposito)

    rs.Close
    Set rs = Nothing

    'Comprobar el período original.
    IdConciliacion = _
        DAOConciliacionBancaria.ObtenerIdConciliacionCerrada( _
            idCuenta, fechaOriginal)

    If IdConciliacion > 0 Then

        UltimoError = _
            "No se puede modificar esta boleta." & vbCrLf & _
            "El período original pertenece a la " & _
            "conciliación bancaria Nº " & IdConciliacion & "."

        Exit Function

    End If

    'Comprobar si alguno de sus movimientos
    'ya aparece en el detalle de una conciliación.
    q = "SELECT d.id_operacion " & _
        "FROM conciliaciones_bancarias_detalles d " & _
        "INNER JOIN cheques_depositos cd " & _
        "ON cd.id_operacion = d.id_operacion " & _
        "WHERE cd.id_boleta = " & idBoleta & " " & _
        "LIMIT 1"

    Set rs = conectar.RSFactory(q)

    If Not rs.EOF Then

        UltimoError = _
            "No se puede modificar esta boleta porque " & _
            "uno de sus movimientos ya figura en " & _
            "el detalle de una conciliación bancaria."

        Exit Function

    End If

    PuedeEditarBoleta = True
    Exit Function

err1:

    UltimoError = "Error al validar la boleta: " & _
                  Err.Description

    PuedeEditarBoleta = False

End Function


Private Function ContieneId( _
    ByVal ids As Collection, _
    ByVal Id As Long) As Boolean

    On Error GoTo noExiste

    Dim valor As Variant

    valor = ids.item(CStr(Id))
    ContieneId = True
    Exit Function

noExiste:
    ContieneId = False

End Function


Public Function Update( _
    ByVal boleta As BoletaDeposito, _
    ByVal original As BoletaDeposito) As Boolean

    On Error GoTo err1

    Dim rs As ADODB.Recordset
    Dim q As String

    Dim idsActuales As New Collection
    Dim operacionesActuales As New Collection
    Dim idsSeleccionados As New Collection
    Dim chequesNuevos As New Collection

    Dim ch As cheque
    Dim chActual As cheque
    Dim chOriginal As cheque
    Dim v As Variant

    Dim idCuentaOriginal As Long
    Dim fechaOriginal As Date
    Dim numeroOriginal As Long
    Dim IdOperacion As Long
    Dim IdConciliacion As Long

    Dim montoTotal As Double
    Dim FechaNueva As Date
    Dim transaccionIniciada As Boolean

    Update = False
    UltimoError = vbNullString

    '-------------------------------------------------
    ' VALIDACIONES GENERALES
    '-------------------------------------------------

    If boleta Is Nothing Or original Is Nothing Then
        UltimoError = "Faltan los datos de la boleta."
        Exit Function
    End If

    If boleta.Id <= 0 Or boleta.Id <> original.Id Then
        UltimoError = "El ID de la boleta no coincide."
        Exit Function
    End If

    If boleta.CuentaDestino Is Nothing Then
        UltimoError = "Debe seleccionar una cuenta bancaria."
        Exit Function
    End If

    If boleta.CuentaDestino.moneda Is Nothing Then
        UltimoError = "La cuenta no tiene moneda definida."
        Exit Function
    End If

    If original.CuentaDestino Is Nothing Then
        UltimoError = "No se conoce la cuenta original."
        Exit Function
    End If

    If boleta.numero <= 0 Or boleta.Cheques.count = 0 Then
        UltimoError = "La boleta debe tener número y cheques."
        Exit Function
    End If

    If CDbl(boleta.fechaDeposito) <= 0 Then
        UltimoError = "La fecha no es válida."
        Exit Function
    End If

    FechaNueva = DateValue(boleta.fechaDeposito)

    '-------------------------------------------------
    ' INICIAR TRANSACCIÓN
    '-------------------------------------------------

    conectar.BeginTransaction
    transaccionIniciada = True

    'Bloquear la cabecera.
    q = "SELECT numero_boleta, fecha_deposito, " & _
        "id_cuenta, tipo_deposito, monto " & _
        "FROM boleta_deposito " & _
        "WHERE id = " & boleta.Id & " FOR UPDATE"

    Set rs = conectar.RSFactory(q)

    If rs.EOF Then
        Err.Raise vbObjectError + 3401, , _
                  "La boleta ya no existe."
    End If

    If IsNull(rs!numero_boleta) Or _
       IsNull(rs!fecha_deposito) Or _
       IsNull(rs!id_cuenta) Or _
       IsNull(rs!tipo_deposito) Or _
       IsNull(rs!Monto) Then

        Err.Raise vbObjectError + 3402, , _
                  "La boleta tiene datos incompletos."
    End If

    numeroOriginal = CLng(rs!numero_boleta)
    fechaOriginal = CDate(rs!fecha_deposito)
    idCuentaOriginal = CLng(rs!id_cuenta)

    If CLng(rs!tipo_deposito) <> DepositoCheque Then
        Err.Raise vbObjectError + 3403, , _
                  "Solamente se pueden editar depósitos de cheques."
    End If

    'Detectar cambios realizados desde otra sesión.
    If original.numero <> numeroOriginal Or _
       DateValue(original.fechaDeposito) <> DateValue(fechaOriginal) Or _
       original.CuentaDestino.Id <> idCuentaOriginal Or _
       Abs(original.Monto - CDbl(rs!Monto)) > 0.01 Then

        Err.Raise vbObjectError + 3404, , _
                  "La boleta fue modificada. Vuelva a abrirla."
    End If

    rs.Close
    Set rs = Nothing

    '-------------------------------------------------
    ' VALIDAR CONCILIACIONES
    '-------------------------------------------------

    IdConciliacion = _
        DAOConciliacionBancaria.ObtenerIdConciliacionCerrada( _
            idCuentaOriginal, fechaOriginal)

    If IdConciliacion > 0 Then
        Err.Raise vbObjectError + 3405, , _
                  "El período bancario original está cerrado."
    End If

    IdConciliacion = _
        DAOConciliacionBancaria.ObtenerIdConciliacionCerrada( _
            boleta.CuentaDestino.Id, FechaNueva)

    If IdConciliacion > 0 Then
        Err.Raise vbObjectError + 3406, , _
                  "El nuevo período bancario está cerrado."
    End If

    q = "SELECT d.id_operacion " & _
        "FROM conciliaciones_bancarias_detalles d " & _
        "INNER JOIN cheques_depositos cd " & _
        "ON cd.id_operacion = d.id_operacion " & _
        "WHERE cd.id_boleta = " & boleta.Id & " LIMIT 1"

    Set rs = conectar.RSFactory(q)

    If Not rs.EOF Then
        Err.Raise vbObjectError + 3407, , _
                  "La boleta tiene operaciones conciliadas."
    End If

    rs.Close
    Set rs = Nothing

    '-------------------------------------------------
    ' BLOQUEAR Y VALIDAR DETALLE ORIGINAL
    '-------------------------------------------------

    q = "SELECT cd.id_cheque, cd.id_operacion, " & _
        "c.depositado, c.en_cartera, " & _
        "c.monto AS monto_cheque, " & _
        "c.id_moneda AS moneda_cheque, " & _
        "o.monto AS monto_operacion, " & _
        "o.cuentabanc_o_caja_id, " & _
        "o.fecha_operacion, o.comprobante, " & _
        "o.pertenencia, o.entrada_salida, " & _
        "o.moneda_id " & _
        "FROM cheques_depositos cd " & _
        "INNER JOIN Cheques c ON c.id = cd.id_cheque " & _
        "INNER JOIN operaciones o ON o.id = cd.id_operacion " & _
        "WHERE cd.id_boleta = " & boleta.Id & _
        " FOR UPDATE"

    Set rs = conectar.RSFactory(q)

    Do While Not rs.EOF
    
    
    If CLng(rs!Pertenencia) <> Banco Or _
       CLng(rs!entrada_salida) <> OPEntrada Then
    
        Err.Raise vbObjectError + 3427, , _
                  "La boleta contiene una operación " & _
                  "que no corresponde a un depósito bancario."
    
    End If
    
    
    If CLng(rs!moneda_id) <> _
       CLng(rs!moneda_cheque) Then
    
        Err.Raise vbObjectError + 3428, , _
                  "La moneda de una operación no coincide " & _
                  "con la moneda de su cheque."
    
    End If

        If Not CBool(rs!Depositado) Or _
           CBool(rs!en_cartera) Or _
           CLng(rs!cuentabanc_o_caja_id) <> idCuentaOriginal Or _
           DateValue(CDate(rs!fecha_operacion)) <> DateValue(fechaOriginal) Or _
           Abs(CDbl(rs!monto_cheque) - _
               CDbl(rs!monto_operacion)) > 0.01 Then

            Err.Raise vbObjectError + 3408, , _
                      "Se detectó un movimiento inconsistente."
        End If

        idsActuales.Add CLng(rs!id_cheque), _
                         CStr(rs!id_cheque)

        operacionesActuales.Add CLng(rs!id_operacion), _
                               CStr(rs!id_cheque)

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    'Comprobar que ningún vínculo haya desaparecido.
    q = "SELECT COUNT(*) AS cantidad " & _
        "FROM cheques_depositos " & _
        "WHERE id_boleta = " & boleta.Id

    Set rs = conectar.RSFactory(q)

    If CLng(rs!Cantidad) <> idsActuales.count Then
        Err.Raise vbObjectError + 3409, , _
                  "Existen vínculos incompletos en la boleta."
    End If

    rs.Close
    Set rs = Nothing

    'Comparar el detalle actual contra el original.
    If idsActuales.count <> original.Cheques.count Then
        Err.Raise vbObjectError + 3410, , _
                  "El detalle cambió. Vuelva a abrir la boleta."
    End If

    For Each chOriginal In original.Cheques

        If Not ContieneId(idsActuales, chOriginal.Id) Then
            Err.Raise vbObjectError + 3411, , _
                      "El detalle fue modificado por otro usuario."
        End If

    Next chOriginal

    '-------------------------------------------------
    ' VALIDAR LOS CHEQUES SELECCIONADOS
    '-------------------------------------------------

    montoTotal = 0

    For Each ch In boleta.Cheques

        If ContieneId(idsSeleccionados, ch.Id) Then
            Err.Raise vbObjectError + 3412, , _
                      "Hay un cheque duplicado."
        End If

        idsSeleccionados.Add True, CStr(ch.Id)

        'Bloquear también los cheques nuevos.
        If Not ContieneId(idsActuales, ch.Id) Then

            q = "SELECT id FROM Cheques " & _
                "WHERE id = " & ch.Id & " FOR UPDATE"

            Set rs = conectar.RSFactory(q)

            If rs.EOF Then
                Err.Raise vbObjectError + 3413, , _
                          "No se encontró uno de los cheques."
            End If

            rs.Close
            Set rs = Nothing

        End If

        Set chActual = DAOCheques.FindById(ch.Id)

        If chActual Is Nothing Then
            Err.Raise vbObjectError + 3414, , _
                      "No se pudo recuperar un cheque."
        End If

        If chActual.moneda Is Nothing Then
            Err.Raise vbObjectError + 3415, , _
                      "Uno de los cheques no tiene moneda."
        End If

        If chActual.moneda.Id <> _
           boleta.CuentaDestino.moneda.Id Then

            Err.Raise vbObjectError + 3416, , _
                      "Hay cheques cuya moneda no coincide."
        End If

        If Not ContieneId(idsActuales, ch.Id) Then

            If Not chActual.EnCartera Or _
               chActual.Depositado Then

                Err.Raise vbObjectError + 3417, , _
                          "Un cheque nuevo ya no está disponible."
            End If

            'Impedir que un cheque esté en dos depósitos.
            q = "SELECT id FROM cheques_depositos " & _
                "WHERE id_cheque = " & ch.Id & " LIMIT 1"

            Set rs = conectar.RSFactory(q)

            If Not rs.EOF Then
                Err.Raise vbObjectError + 3418, , _
                          "Un cheque ya pertenece a otro depósito."
            End If

            rs.Close
            Set rs = Nothing

            chequesNuevos.Add chActual, CStr(chActual.Id)

        End If

        montoTotal = montoTotal + chActual.Monto

    Next ch

    '-------------------------------------------------
    ' QUITAR CHEQUES Y ACTUALIZAR LOS EXISTENTES
    '-------------------------------------------------

    For Each v In idsActuales

        IdOperacion = CLng(operacionesActuales.item(CStr(v)))

        If Not ContieneId(idsSeleccionados, CLng(v)) Then

            'Comprobar que no tenga otra vinculación.
            q = "SELECT id FROM cheques_depositos " & _
                "WHERE id_cheque = " & CLng(v) & _
                " AND (id_boleta <> " & boleta.Id & _
                " OR id_boleta IS NULL) LIMIT 1"

            Set rs = conectar.RSFactory(q)

            If Not rs.EOF Then
                Err.Raise vbObjectError + 3419, , _
                          "Un cheque tiene otra vinculación."
            End If

            rs.Close
            Set rs = Nothing
    
        '-------------------------------------------------
        ' VERIFICAR QUE LA OPERACION SEA EXCLUSIVA
        ' DE ESTE CHEQUE / DEPOSITO
        '-------------------------------------------------
        
        q = "SELECT COUNT(*) AS cantidad " & _
            "FROM cheques_depositos " & _
            "WHERE id_operacion = " & IdOperacion
        
        Set rs = conectar.RSFactory(q)
        
        If CLng(rs!Cantidad) <> 1 Then
        
            Err.Raise vbObjectError + 3429, , _
                      "La operación bancaria del cheque " & _
                      "tiene más de una vinculación."
        
        End If
        
        rs.Close
        Set rs = Nothing

            'Devolverlo a cartera.
            q = "UPDATE Cheques SET " & _
                "depositado = 0, en_cartera = 1 " & _
                "WHERE id = " & CLng(v)

            If Not conectar.execute(q) Then
                Err.Raise vbObjectError + 3420, , _
                          "No se pudo liberar un cheque."
            End If

            'Eliminar vínculo y operación exclusivamente
            'de esta boleta.
            q = "DELETE FROM cheques_depositos " & _
                "WHERE id_boleta = " & boleta.Id & _
                " AND id_cheque = " & CLng(v) & _
                " AND id_operacion = " & IdOperacion

            If Not conectar.execute(q) Then
                Err.Raise vbObjectError + 3421, , _
                          "No se pudo eliminar una relación."
            End If

            q = "DELETE FROM operaciones " & _
                "WHERE id = " & IdOperacion

            If Not conectar.execute(q) Then
                Err.Raise vbObjectError + 3422, , _
                          "No se pudo eliminar una operación."
            End If

        Else

            'Conservar el ID del movimiento existente.
            q = "UPDATE operaciones SET " & _
                "fecha_operacion = " & _
                conectar.Escape(FechaNueva) & ", " & _
                "cuentabanc_o_caja_id = " & _
                boleta.CuentaDestino.Id & ", " & _
                "comprobante = " & _
                conectar.Escape(CStr(boleta.numero)) & _
                " WHERE id = " & IdOperacion

            If Not conectar.execute(q) Then
                Err.Raise vbObjectError + 3423, , _
                          "No se pudo actualizar un movimiento."
            End If

        End If

    Next v

    '-------------------------------------------------
    ' AGREGAR CHEQUES NUEVOS
    '-------------------------------------------------

    For Each chActual In chequesNuevos

        If Not DepositarChequeInterno( _
            chActual, _
            boleta.CuentaDestino, _
            FechaNueva, _
            CStr(boleta.numero), _
            boleta.Id) Then

            Err.Raise vbObjectError + 3424, , _
                      "No se pudo agregar un cheque nuevo."
        End If

    Next chActual

    '-------------------------------------------------
    ' ACTUALIZAR CABECERA
    '-------------------------------------------------

    q = "UPDATE boleta_deposito SET " & _
        "numero_boleta = " & boleta.numero & ", " & _
        "fecha_deposito = " & _
        conectar.Escape(FechaNueva) & ", " & _
        "id_cuenta = " & boleta.CuentaDestino.Id & ", " & _
        "monto = " & conectar.Escape(montoTotal) & _
        " WHERE id = " & boleta.Id

    If Not conectar.execute(q) Then
        Err.Raise vbObjectError + 3425, , _
                  "No se pudo actualizar la cabecera."
    End If

    '-------------------------------------------------
    ' COMPROBACIÓN FINAL
    '-------------------------------------------------

    q = "SELECT COUNT(*) AS cantidad, " & _
        "COALESCE(SUM(c.monto), 0) AS total " & _
        "FROM cheques_depositos cd " & _
        "INNER JOIN Cheques c ON c.id = cd.id_cheque " & _
        "WHERE cd.id_boleta = " & boleta.Id

    Set rs = conectar.RSFactory(q)

    If CLng(rs!Cantidad) <> boleta.Cheques.count Or _
       Abs(CDbl(rs!total) - montoTotal) > 0.01 Then

        Err.Raise vbObjectError + 3426, , _
                  "El detalle no coincide con el total."
    End If

    rs.Close
    Set rs = Nothing

    conectar.CommitTransaction
    transaccionIniciada = False

    boleta.Monto = montoTotal
    Update = True

    Exit Function

err1:

    UltimoError = Err.Description

    On Error Resume Next

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    If transaccionIniciada Then
        conectar.RollBackTransaction
    End If

    Update = False

End Function
