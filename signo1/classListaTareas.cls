VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classListaTareas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim strsql As String
Dim cn As ADODB.Connection
Dim rs As ADODB.recordset
Private Sub Class_Initialize()
    On Error GoTo err22
    Set cn = conectar.obternerConexion    'New ADODB.Connection
    'Set rs = New ADODB.recordset
    'CN.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & serverBBDD & ";uid=root;pwd=3l3c720n;database=sp;OPTION=3;connection=adUseClient"
    'CN.Open
    Exit Sub
err22:
    MsgBox Err.Description
End Sub
Public Function ejecutar_consulta(consulta As String)
    On Error GoTo err14:
    Set rs = conectar.RSFactory(consulta)
    'If rs.State = 1 Then rs.Close
    'rs.Open consulta, CN, adOpenDynamic, adLockOptimistic, adCmdText
    Exit Function
err14:
    MsgBox "Se produjo un error: " & Err.Description

End Function

Public Function crear_historico(id_material, Valor, id_moneda, accion) As Integer
    fech = Format(Date, "yyyy/mm/dd")

    If accion = 0 Then    '0=nuevo
        strsql = "insert into historico (id_material,valor,fecha_actualizacion,id_moneda) VALUES (" & id_material & "," & Valor & " ,'" & fech & "'," & id_moneda & ")"
        Me.ejecutar_consulta strsql
    ElseIf accion = 1 Then    '1=modificacion
        strsql = "SELECT count(id) as canti from historico where fecha_actualizacion = '" & fech & "' and id_material=" & id_material
        Me.ejecutar_consulta strsql
        Cant = rs!canti


        If Cant = 0 Then
            strsql = "insert into historico (id_material,valor,fecha_actualizacion,id_moneda) VALUES (" & id_material & "," & Valor & " ,'" & fech & "'," & id_moneda & ")"
            Me.ejecutar_consulta strsql
        Else
            a = MsgBox("Ya existe un cambio el día de hoy para ese producto." & Chr(10) & "¿Desea actualizarlo?", vbYesNo, "Advertencia")
            If a = 6 Then

                strsql = "select id from historico where fecha_actualizacion='" & fech & "'"
                Me.ejecutar_consulta strsql
                id = rs!id
                strsql = "update historico set valor= " & Valor & ", fecha_actualizacion = '" & fech & "',id_moneda=" & id_moneda & " where id=" & id
                Me.ejecutar_consulta strsql
            End If
        End If
    End If
End Function
Public Function LlenarListahistoricos()

    frmHistorico.lstMateriales.ListItems.Clear
    strsql = "select g.grupo,m.id, m.codigo, m.descripcion,m.espesor,m.pesoxunidad,m.id_unidad,r.rubro from materiales m  inner join rubros r on m.id_rubro=r.id inner join grupos g on g.id=m.id_grupo "
    Me.ejecutar_consulta strsql
    Cant = vCantidad_materiales
    While Not rs.EOF

        If rs!id_Unidad = 0 Then
            unidad = "Kg"
        ElseIf rs!id_Unidad = 1 Then
            unidad = "M2"
        Else
            unidad = "Ml"
        End If

        Moneda = "$"


        Set X = frmHistorico.lstMateriales.ListItems.Add(, , rs!codigo)
        X.SubItems(1) = rs!id
        X.SubItems(2) = rs!rubro
        X.SubItems(3) = rs!grupo
        X.SubItems(4) = rs!Descripcion
        X.SubItems(5) = rs!Espesor & " mm"
        rs.MoveNext



    Wend
End Function

Public Function LlenarListahistoricos_valores(id As Integer)

    frmHistorico.lstHistoricosValores.ListItems.Clear
    strsql = "SELECT * from historico WHERE id_material =" & id
    Me.ejecutar_consulta strsql
    Cant = vCantidad_materiales

    While Not rs.EOF

        Moneda = "$"
        If rs!id_moneda = 1 Then Moneda = "u$S"

        Set X = frmHistorico.lstHistoricosValores.ListItems.Add(, , rs!id)
        X.SubItems(1) = rs!FEcha_actualizacion
        X.SubItems(2) = rs!Valor
        X.SubItems(3) = Moneda
        rs.MoveNext
    Wend
    rs.Close
End Function



Private Sub Class_Terminate()
    Set cn = Nothing
    Set rs = Nothing
End Sub
