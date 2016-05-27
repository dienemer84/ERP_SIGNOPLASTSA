VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classNuevoMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim strsql As String
Dim rs As ADODB.recordset
Dim vCantidad_rubros As Integer
Dim vCantidad_grupos As Integer
Private Sub Class_Initialize()
    On Error GoTo err22
    Set cn = conectar.obternerConexion    '  New ADODB.Connection
    Set rs = New ADODB.recordset
    '    CN.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & serverBBDD & ";uid=root;pwd=3l3c720n;database=sp;OPTION=3;connection=adUseClient"
    '   CN.Open

    If rs.State = 1 Then rs.Close
    Me.ejecutar_consulta "select count(id) as canti from rubros"
    vCantidad_rubros = rs!canti
    Me.ejecutar_consulta "select count(id) as canti from grupos"
    vCantidad_grupos = rs!canti
    Exit Sub
err22:
    MsgBox Err.Description
End Sub

Public Property Get ultimo_agregado() As Integer
    Me.ejecutar_consulta "select last_insert_id() as nuevo from materiales"
    ultimo_agregado = rs!nuevo


End Property
Public Function existeCodigo(codigo As String) As Boolean
    On Error GoTo err3
    existeCodigo = False

    Me.ejecutar_consulta "select count(id) as cantid from materiales where codigo='" & UCase(Trim(codigo)) & "'"
    If rs!cantid > 0 Then existeCodigo = True

    Exit Function
err3:
    existeCodigo = False
End Function

Public Function llenar_combo_rubros()
    strsql = "select * from rubros order by rubro"
    Me.ejecutar_consulta strsql

    frmMaterialesNuevo.cboRubros.Clear

    While Not rs.EOF

        frmMaterialesNuevo.cboRubros.AddItem rs!rubro
        frmMaterialesNuevo.cboRubros.ItemData(frmMaterialesNuevo.cboRubros.NewIndex) = rs!Id
        rs.MoveNext
    Wend
    rs.Close

    If frmMaterialesNuevo.cboRubros.ListCount > 0 Then
        frmMaterialesNuevo.cboRubros.ListIndex = 0
    End If

End Function

Public Function ejecutar_consulta(consulta As String)
    On Error GoTo err10
    'If rs.State = 1 Then rs.Close
    'rs.Open consulta, CN, adOpenDynamic, adLockOptimistic, adCmdText
    Set rs = conectar.RSFactory(consulta)
    Exit Function
err10:
    MsgBox "Se produjo un error: " & Err.Description
End Function
Public Function llenar_combo_grupos(id_rubro As Integer)
    frmMaterialesNuevo.cboGrupos.Clear
    Me.ejecutar_consulta "select count(id) as canti_grupos from grupos where id_rubro=" & id_rubro
    vCantid = rs!canti_grupos
    Me.ejecutar_consulta "select * from grupos where id_rubro=" & id_rubro & " order  by grupo"

    If vCantidad_grupos > 0 Then

        For i = 1 To vCantid
            frmMaterialesNuevo.cboGrupos.AddItem rs!grupo
            frmMaterialesNuevo.cboGrupos.ItemData(frmMaterialesNuevo.cboGrupos.NewIndex) = rs!Id
            rs.MoveNext
        Next
    End If
    rs.Close
    If frmMaterialesNuevo.cboGrupos.ListCount > 0 Then
        frmMaterialesNuevo.cboGrupos.ListIndex = 0
    End If
End Function

Private Sub Class_Terminate()
    Set cn = Nothing
    Set rs = Nothing
End Sub
