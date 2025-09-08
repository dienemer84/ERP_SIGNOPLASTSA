Attribute VB_Name = "DAOProduccion"
Option Explicit
Public LastError As String

Public Function SaveMany(ByVal regs As Collection) As Boolean
    On Error GoTo err1
    Dim cn As ADODB.Connection, cmd As ADODB.Command
    Set cn = GetConnection()                 'tu fábrica de conexiones
    cn.BeginTrans

    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = "INSERT INTO pedidos_produccion_carga " & _
            "(id_pedido,id_pieza_pedido,cant_recibida,cant_fabricada,cant_scrap,fecha_inicio,fecha_fin,recibio,siguiente_proceso) " & _
            "VALUES (?,?,?,?,?,?,?,?,?)"

        'Definir parámetros una sola vez
        .Parameters.Append .CreateParameter("@id_pedido", adInteger, adParamInput)
        .Parameters.Append .CreateParameter("@id_pieza_pedido", adInteger, adParamInput)
        .Parameters.Append .CreateParameter("@cant_recibida", adBigInt, adParamInput)
        .Parameters.Append .CreateParameter("@cant_fabricada", adBigInt, adParamInput)
        .Parameters.Append .CreateParameter("@cant_scrap", adBigInt, adParamInput)
        .Parameters.Append .CreateParameter("@fecha_inicio", adDate, adParamInput)
        .Parameters.Append .CreateParameter("@fecha_fin", adDate, adParamInput)
        .Parameters.Append .CreateParameter("@recibio", adInteger, adParamInput)
        .Parameters.Append .CreateParameter("@siguiente_proceso", adVarChar, adParamInput, 50)
    End With

    Dim i As Long, reg As RegistroProduccion
    For i = 1 To regs.count
        Set reg = regs(i)
        cmd.Parameters(0).value = reg.id_pedido
        cmd.Parameters(1).value = reg.id_pieza_pedido
        cmd.Parameters(2).value = reg.cant_recibida
        cmd.Parameters(3).value = reg.cant_fabricada
        cmd.Parameters(4).value = reg.cant_scrap
        cmd.Parameters(5).value = IIf(IsNull(reg.fecha_inicio), Null, reg.fecha_inicio)
        cmd.Parameters(6).value = IIf(IsNull(reg.fecha_fin), Null, reg.fecha_fin)
        cmd.Parameters(7).value = reg.recibio
        cmd.Parameters(8).value = reg.siguiente_proceso
        cmd.execute , , adExecuteNoRecords
    Next i

    cn.CommitTrans
    SaveMany = True
    Exit Function

err1:
    On Error Resume Next
    LastError = Err.Description
    If Not cn Is Nothing Then If cn.State = adStateOpen Then cn.RollbackTrans
    SaveMany = False
End Function

