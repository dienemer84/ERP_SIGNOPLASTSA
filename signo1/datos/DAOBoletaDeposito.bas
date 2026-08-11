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
            ByVal comprobante As String, _
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

    op.comprobante = comprobante


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

