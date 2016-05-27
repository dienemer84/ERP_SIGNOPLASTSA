VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim strsql As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
'



Private Sub Class_Initialize()
    On Error GoTo err22
    Set cn = conectar.obternerConexion


    Exit Sub

err22:
    MsgBox Err.Description
End Sub
Public Function cambiarPrecios(id As Integer, incremental As Double, accion As Integer, Optional accion2 As Integer) As Boolean
    On Error GoTo erra
    Dim Valor As Double

    inc = 1 + (incremental / 100)    'valor incremental/decremental
    fech = Format(Date, "yyyy/mm/dd")    'fecha actual
    cambiarPrecios = True
    If accion = 1 Then    'por rubro
        Me.ejecutar_consulta "select m.id,m.valor_unitario,m.id_moneda from  materiales m where  m.id_Rubro=" & id
    ElseIf accion = 2 Then    'por grupo
        Me.ejecutar_consulta "select m.id,m.valor_unitario,m.id_moneda from  materiales m where m.id_grupo=" & id
    ElseIf accion = 3 Then    'por material
        Me.ejecutar_consulta "select m.id,m.valor_unitario,m.id_moneda from  materiales m where m.id=" & id
    End If

    While Not rs.EOF
        IdMoneda = rs!id_moneda
        If accion2 = 1 Then
            Valor = incremental
        Else
            Valor = rs!valor_unitario * inc
        End If

        IdMaterial = rs!id

        Valor = funciones.RedondearDecimales(Valor, 2)

        tra = True
        cn.BeginTrans
        cn.execute "insert into historico (id_material,valor,fecha_actualizacion,id_moneda) VALUES (" & IdMaterial & "," & Valor & " ,'" & fech & "'," & IdMoneda & ")"
        ahora = funciones.dateFormateada(Now)
        cn.execute "update materiales set fecha_valor='" & ahora & "', valor_unitario=" & Valor & " where id=" & IdMaterial

        cn.CommitTrans
        tra = False


        rs.MoveNext
    Wend
    Exit Function
erra:
    cambiarPrecios = False
    If tra = True Then cn.RollbackTrans

End Function
Public Function ejecutar_consulta(consulta As String)
    On Error GoTo err44
    Set rs = conectar.RSFactory(consulta)
    Exit Function
err44:
    MsgBox "Se produjo un error: " & Err.Description
End Function





Private Sub Class_Terminate()
    Set cn = Nothing
    Set rs = Nothing
End Sub





