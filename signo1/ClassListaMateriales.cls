VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassListaMateriales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim strsql As String
Dim CN As ADODB.Connection
Dim rs As ADODB.Recordset
Dim vCantidad_rubros As Integer
Dim vCantidad_grupos As Integer
Dim vCantidad_materiales As Integer
Public Property Get idUnidad() As Integer
idUnidad = rs!id_unidad
End Property
Private Sub Class_Initialize()
On Error GoTo err22
   Set CN = New ADODB.Connection
   Set rs = New ADODB.Recordset
   CN.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & serverBBDD & ";uid=root;pwd=3l3c720n;database=sp;OPTION=3;connection=adUseClient"
   CN.Open
    
If rs.State = 1 Then rs.Close
Me.ejecutar_consulta "select count(id) as canti from rubros"
vCantidad_rubros = rs!canti

Me.ejecutar_consulta "select count(id) as canti from grupos"
vCantidad_grupos = rs!canti
Me.ejecutar_consulta "select count(id) as canti from materiales"
vCantidad_materiales = rs!canti
Exit Sub
err22:
    MsgBox Err.Description

End Sub
Public Function ejecutar_consulta(Consulta As String)
On Error GoTo err14:
If rs.State = 1 Then rs.Close
rs.Open Consulta, CN, adOpenDynamic, adLockOptimistic, adCmdText
Exit Function
err14:
MsgBox "Se produjo un error: " & Err.Description

End Function




Public Function verifica_codigo(codigos As String) As Integer
If rs.State = 1 Then rs.Close
strsql = "select count(id) AS canti from materiales where codigo='" & codigos & "'"
Me.ejecutar_consulta strsql
With rs
If !canti > 0 Then
 verifica_codigo = 1 'si hay codigo, devuelvo 1
Else
 verifica_codigo = 0 'si no hay, devuelvo 0
End If
End With
End Function


Public Function QueIdMaterial(codigos As String) As Long
On Error Resume Next
strsql = "select id  from materiales where codigo='" & normaliza(codigos) & "'"
Me.ejecutar_consulta strsql
c = 0
While Not rs.EOF
c = c + 1
rs.MoveNext
Wend
If c = 1 Then
rs.MoveFirst
    QueIdMaterial = rs!id  'si exist codigo, dvuelvo id
Else
    QueIdMaterial = -1 'sino devuelvo -1
End If

End Function

Public Function crear_historico(id_material, valor, id_moneda, accion) As Integer
    Fech = Format(Date, "yyyy/mm/dd")
    
    If accion = 0 Then '0=nuevo
     strsql = "insert into historico (id_material,valor,fecha_actualizacion,id_moneda) VALUES (" & id_material & "," & valor & " ,'" & Fech & "'," & id_moneda & ")"
    Me.ejecutar_consulta strsql
    ElseIf accion = 1 Then '1=modificacion
    strsql = "SELECT count(id) as canti from historico where fecha_actualizacion = '" & Fech & "' and id_material=" & id_material
    Me.ejecutar_consulta strsql
    Cant = rs!canti
    
    If Cant = 0 Then
     strsql = "insert into historico (id_material,valor,fecha_actualizacion,id_moneda) VALUES (" & id_material & "," & valor & " ,'" & Fech & "'," & id_moneda & ")"
     Me.ejecutar_consulta strsql
    Else
    a = MsgBox("Ya existe un cambio el día de hoy para ese producto." & Chr(10) & "¿Desea actualizarlo?", vbYesNo, "Advertencia")
        If a = 6 Then
          strsql = "select id from historico where fecha_actualizacion='" & Fech & "'"
          Me.ejecutar_consulta strsql
          id = rs!id
          
          strsql = "update historico set valor= " & valor & ", fecha_actualizacion = '" & Fech & "',id_moneda=" & id_moneda & " where id=" & id
          Me.ejecutar_consulta strsql
        End If
    End If
    End If

End Function

Public Function LlenarListahistoricos()

frmComprasPreciosHistorico.lstMateriales.ListItems.Clear
strsql = "select g.grupo,m.id, m.codigo, m.descripcion,m.espesor,m.pesoxunidad,m.id_unidad,r.rubro from materiales m  inner join rubros r on m.id_rubro=r.id inner join grupos g on g.id=m.id_grupo "
Me.ejecutar_consulta strsql
Cant = vCantidad_materiales
While Not rs.EOF

If rs!id_unidad = 0 Then
   Unidad = "Kg"
ElseIf rs!id_unidad = 1 Then
    Unidad = "M2"
Else
    Unidad = "Ml"
End If
Moneda = "$"
    Set X = frmComprasPreciosHistorico.lstMateriales.ListItems.Add(, , rs!codigo)
     X.SubItems(1) = rs!id
     X.SubItems(2) = rs!rubro
     X.SubItems(3) = rs!grupo
     X.SubItems(4) = rs!descripcion
     X.SubItems(5) = rs!espesor & " mm"
rs.MoveNext
Wend
End Function

Public Function LlenarListahistoricos_valores(id As Integer)

frmComprasPreciosHistorico.lstHistoricosValores.ListItems.Clear
strsql = "SELECT * from historico WHERE id_material =" & id
Me.ejecutar_consulta strsql
Cant = vCantidad_materiales
While Not rs.EOF
Moneda = "$"
If rs!id_moneda = 1 Then Moneda = "u$S"
    Set X = frmComprasPreciosHistorico.lstHistoricosValores.ListItems.Add(, , rs!id)
     X.SubItems(1) = rs!fecha_actualizacion
     X.SubItems(2) = funciones.RedondearDecimales(rs!valor, 3)
     X.SubItems(3) = Moneda
rs.MoveNext
Wend
End Function
Public Function LlenarListaTareas()
frmListaMateriales.lstMateriales.ListItems.Clear
strsql = "select v.id_tarea,v.valor, v.fecha, v.descripcion,s.sector,t.tarea,t.id_sector,t.cantxproc from tareas t inner join sectores s on t.id_sector=s.id inner join  valores_MDO v on v.id_tarea=t.id"
Me.ejecutar_consulta strsql
While Not rs.EOF
   
    Set X = frmListaTareas.lstTareas.ListItems.Add(, , rs!id)
     X.SubItems(1) = rs!sector
     X.SubItems(2) = rs!cantxproc
     X.SubItems(3) = rs!tarea
     X.SubItems(4) = rs!descripcion
     X.SubItems(5) = rs!FEcha
     rs.MoveNext
Wend
End Function

Private Sub Class_Terminate()
Set CN = Nothing
Set rs = Nothing
End Sub
