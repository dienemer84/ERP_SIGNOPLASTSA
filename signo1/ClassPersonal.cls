VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.recordset
Dim strsql As String
Dim vLegajo As Long
Dim vDocumento As Long
Dim vApellido As String
Dim vNombre As String
Dim vNombres As String
Dim vDireccion As String
Dim vTel1 As String, vTel2 As String
Attribute vTel2.VB_VarUserMemId = 1073938441
Dim vLocalidad As String
Attribute vLocalidad.VB_VarUserMemId = 1073938443
Dim vUser As Long
Attribute vUser.VB_VarUserMemId = 1073938444
Dim vmail As String
Attribute vmail.VB_VarUserMemId = 1073938445
Dim vapto As Boolean
Attribute vapto.VB_VarUserMemId = 1073938446

Public GrupoSanguineo As String
Public FechaIngreso As Date
Public FechaNacimiento As Date

Public Property Let Apellido(nApellido As String)
    vApellido = normaliza(nApellido)
End Property
Public Property Let aptoSistema(napto As Boolean)
    vapto = napto
End Property
Public Property Get Apellido() As String
    Apellido = rs!Apellido
End Property

Public Property Let localidad(nLocalidad As String)
    vLocalidad = normaliza(nLocalidad)
End Property
Public Property Get localidad() As String
    localidad = rs!localidad
End Property

Public Property Let legajo(nLegajo As Long)
    vLegajo = nLegajo
End Property
Public Property Get legajo() As Long
    legajo = rs!legajo
End Property

Public Property Let Documento(nDocumento As Long)
    vDocumento = nDocumento
End Property
Public Property Get Documento() As Long
    Documento = rs!Documento
End Property

Public Property Let nombre(nNombre As String)
    vNombre = normaliza(nNombre)
End Property
Public Property Get nombre() As String
    nombre = rs!nombre
End Property


Public Property Let Email(nmail As String)
    vmail = LCase(nmail)
End Property
Public Property Get Email() As String
    Email = rs!Email
End Property


Public Property Let direccion(nDireccion As String)
    vDireccion = normaliza(nDireccion)
End Property
Public Property Get direccion() As String
    direccion = rs!direccion
End Property

Public Property Let tel1(nTel1 As String)
    vTel1 = normaliza(nTel1)
End Property
Public Property Get tel1() As String
    tel1 = rs!Telefono1
End Property

Public Property Let tel2(nTel2 As String)
    vTel2 = normaliza(nTel2)
End Property
'GETTERS
Public Property Get tel2() As String
    tel2 = rs!Telefono2
End Property
Public Property Get estado() As Integer
    estado = rs!estado
End Property
Public Property Let Nombres(nNombres As String)
    vNombres = normaliza(nNombres)
End Property
Public Property Get Nombres() As String
    Nombres = rs!Nombres
End Property

Private Sub Class_Initialize()
    On Error GoTo err22
    Set cn = conectar.obternerConexion
    Exit Sub
err22:
    MsgBox Err.Description
End Sub
Public Function ejecutar(consulta As String) As Boolean
    On Error GoTo errEj
    ejecutar = True
    'If rs.State = 1 Then rs.Close
    'rs.Open consulta, CN, adOpenDynamic, adLockOptimistic, adCmdText

    Set rs = conectar.RSFactory(consulta)
    Exit Function
errEj:
    MsgBox "Se produjo un error: " & Err.Description
    ejecutar = False
End Function
Public Function existeLegajo(leg As Long) As Boolean
    Me.ejecutar "Select count(id) as cantidad from personal where legajo=" & leg
    If rs!Cantidad > 0 Then
        existeLegajo = True
    Else
        existeLegajo = False
    End If
End Function
Public Function altaEmpleado() As Boolean
    Dim m As New classMD5
    Dim Usuario As String
    altaEmpleado = True
    On Error GoTo err3
    cn.BeginTrans
    tra = True

    strsql = "insert into personal (legajo,documento,apellido,nombre,nombres,direccion,localidad,telefono1,telefono2,estado, grupo_sanguineo, fecha_ingreso, fecha_nacimiento) values "
    strsql = strsql & "(" & vLegajo & "," & vDocumento & ",'" & vApellido & "','" & vNombre & "','" & vNombres & "','" & vDireccion & "','" & vLocalidad & "','" & vTel1 & "','" & vTel2 & "',1," & Escape(Me.GrupoSanguineo) & ", " & Escape(Me.FechaIngreso) & ", " & Escape(Me.FechaNacimiento) & " )"
    'inserto empleado
    cn.execute strsql
    Me.ejecutar "select last_insert_id() as ultid from personal"
    uid = rs!ultid
    Usuario = crearUsuario(vNombre, vApellido)
    'inserto usuario
    esta = 1
    If vapto = False Then esta = 0
    pass = m.DigestStrToHexStr(Usuario)
    cn.execute "insert into usuarios (idEmpleado,usuario,password,estado) values (" & uid & ",'" & Usuario & "','" & pass & "'," & esta & ")"
    Me.ejecutar "select last_insert_id() as ultid from usuarios"
    uid = rs!ultid
    'inserto permisos x default

    cn.execute "insert into sp_permisos.Config (idUsuario) values (" & uid & ")"

    cn.execute "insert into sp_permisos.Plan (idUsuario) values (" & uid & ")"
    cn.execute "insert into sp_permisos.Desarrollo (idUsuario) values (" & uid & ")"
    cn.execute "insert into sp_permisos.Ventas (idUsuario) values (" & uid & ")"
    cn.execute "insert into sp_permisos.Administracion (idUsuario) values (" & uid & ")"
    cn.execute "insert into sp_permisos.Compras (idUsuario) values (" & uid & ")"
    tra = False
    cn.CommitTrans
    Exit Function
err3:
    If tra Then cn.RollbackTrans
    altaEmpleado = False




End Function

Public Function lista() As recordset
    Me.ejecutar "Select * from personal"
    Set lista = rs
End Function
Public Function ListaSectoresEmpleados(legajo As Long) As recordset

    Dim r As recordset
    Set r = conectar.RSFactory("select id from personal where legajo=" & legajo)
    If Not r.EOF And Not r.BOF Then
        idEmpleado = r!Id
    End If
    Set r = Nothing
    If Me.ejecutar("select e.idsector as id, s.sector from sectores s,sectorizacion e where e.idSector=s.id and e.idEmpleado=" & idEmpleado) Then
        Set ListaSectoresEmpleados = rs
    Else
        MsgBox "Se produjo un error", vbCritical, "Error"
    End If
End Function

Public Function sectorizar(lst As ListView, legajo As Long)
    On Error GoTo err1
    Dim tra As Boolean
    cn.BeginTrans
    tra = True
    Dim r As recordset
    Set r = conectar.RSFactory("select id from personal where legajo=" & legajo)
    If Not r.EOF And Not r.BOF Then
        idEmpleado = r!Id
    End If
    Set r = Nothing
    cn.execute "delete from sectorizacion where idEmpleado=" & idEmpleado
    For X = 1 To lst.ListItems.count
        idSector = CLng(lst.ListItems(X).ListSubItems(1))
        cn.execute "insert into sectorizacion (idEmpleado,idSector) values (" & idEmpleado & "," & idSector & ")"
    Next X
    cn.CommitTrans
    tra = False
    Exit Function
err1:
    If tra Then cn.RollbackTrans
End Function
Public Function getUserName(IdU) As String
    Me.ejecutar "select concat(p.apellido,', ', p.nombre) as nombre,u.usuario from personal p,usuarios u where u.idEmpleado=p.id and u.id=14"
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend
    If c = 1 Then
        rs.MoveFirst
        gerUserName = rs!nombre
    End If
End Function
Public Function addMensajes(enviado, destinatarios, adjuntos, mensaje, Asunto, Usuario)
    On Error GoTo err123
    strsql = "insert into usuariosMensajes ( FechaEnviado, Destinatarios, Adjuntos, Asunto, Mensaje, idUsuario)     values  ('" & enviado & "','" & destinatarios & "','" & adjuntos & "','" & Asunto & "','" & mensaje & "'," & Usuario & ")"
    cn.execute strsql
    Exit Function
err123:
    MsgBox "Error: msgbox err.description"
End Function
Public Function marcarCalendario(cal As stCalendar, Mes As Integer)
    strsql = "select day(fecha) as d, year(fecha) as y,month(fecha) as m,tipo from usuariosAgenda where month(fecha)=2 and idUsuario=" & funciones.getUser
    Me.ejecutar strsql
    While Not rs.EOF
        If rs!Y = cal.cYear And rs!m = cal.cMonth Then
            cal.DayMarking rs!d, rs!Tipo, True
        End If
        rs.MoveNext
    Wend
    cal.CalendarRedraw
End Function
Public Function queEmpleado(legajo As Long) As String
'devuielve nombre, apellido
    Me.ejecutar "Select apellido, nombre from personal where legajo=" & legajo
    queEmpleado = "S/D"
    If Not rs.EOF And Not rs.BOF Then
        queEmpleado = Apellido & ", " & nombre
    End If
End Function

Private Sub Class_Terminate()
    Set cn = Nothing
    Set rs = Nothing
End Sub
