Attribute VB_Name = "DAOBoletaDeposito"
Option Explicit

Public Function Depositar(cheque As cheque, cuenta As CuentaBancaria, FEcha As Date) As Boolean
    On Error GoTo err1
    'cheques_depositos

    Dim op As New operacion    'operacion de ingreso de papota a la cuenta
    op.IdPertenencia = cheque.Id
    op.EntradaSalida = OPEntrada
    op.FechaCarga = Now
    op.FechaOperacion = FEcha
    op.Pertenencia = Banco
    Set op.moneda = cheque.moneda
    op.Monto = cheque.Monto
    If Not DAOOperacion.Save(op) Then GoTo err1
    op.Id = conectar.UltimoId2


    cheque.Depositado = True
    cheque.EnCartera = False
    If Not DAOCheques.Guardar(cheque) Then GoTo err1

    'si esta todo ok, guardo el deposito

    If Not conectar.execute("insert into cheques_depositos (id_cheque, id_operacion) values(" & op.Id & "," & cheque.Id & ")") Then GoTo err1
    Depositar = True
    Exit Function
err1:
    Depositar = False

End Function


Public Function Save() As Boolean

End Function
