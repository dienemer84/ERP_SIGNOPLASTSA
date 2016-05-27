VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classListaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim strsql As String
Dim cn As ADODB.Connection
Dim rs As ADODB.recordset
Dim classA As New classAdministracion

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
Public Function ejecutar(consulta As String)
    On Error GoTo err14:

    Set rs = conectar.RSFactory(consulta)
    'If rs.State = 1 Then rs.Close
    'rs.Open consulta, CN, adOpenDynamic, adLockOptimistic, adCmdText
    Exit Function
err14:
    MsgBox "Se produjo un error: " & Err.Description

End Function

Public Function ejecutar_consulta(consulta As String) As Integer
    ejecutar_contula = 0
    On Error GoTo e2:
    cn.BeginTrans
    g = True
    cn.execute consulta
    cn.CommitTrans
    Exit Function
e2:
    If g Then
        MsgBox Err.Description
        ejecutar_consulta = -1
        cn.RollbackTrans
    Else

    End If

End Function
Function llenar_form_mod(id As Integer)
    strsql = "SELECT * FROM clientes where id=" & id
    Me.ejecutar strsql


    Razon = rs!Razon
    Domicilio = rs!Domicilio
    localidad = rs!localidad
    CP = rs!CP
    telefono = rs!telefono
    Fax = rs!Fax
    Email = rs!Email
    Cuit = rs!Cuit
    Ciudad = rs!Ciudad
    provincia = rs!provincia
    pais = rs!pais
    FP = rs!FP


    frmVentasClienteModificar.Text1(0) = Razon
    frmVentasClienteModificar.Text1(1) = Domicilio
    frmVentasClienteModificar.Text1(2) = localidad
    frmVentasClienteModificar.Text1(3) = CP
    frmVentasClienteModificar.Text1(4) = telefono
    frmVentasClienteModificar.Text1(5) = Fax
    frmVentasClienteModificar.Text1(6) = Email
    frmVentasClienteModificar.Text1(7) = Cuit
    frmVentasClienteModificar.Text1(8) = Ciudad
    frmVentasClienteModificar.Text1(9) = provincia
    frmVentasClienteModificar.Text1(10) = pais
    frmVentasClienteModificar.txtFP = FP
    'frmModificarCliente.ivas(CInt(frmModificarCliente.lbliva)).value = True
    frmVentasClienteModificar.cboIVA.ListIndex = funciones.PosIndexCbo(rs!Iva, frmVentasClienteModificar.cboIVA)
End Function

Function buscar_cliente(id As Integer)
    strsql = "select * from clientes where id=" & id
    Me.ejecutar strsql
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend
    If c = 1 Then
        rs.MoveFirst
        buscar_cliente = rs!Razon
    Else
        buscar_cliente = Empty
    End If
End Function




Private Sub Class_Terminate()
    Set cn = Nothing
    Set rs = Nothing
End Sub


