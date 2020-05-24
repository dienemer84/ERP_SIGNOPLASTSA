VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classRubrosGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.recordset
Dim vConsulta As String
Dim vtabla As Integer
Dim vCantidad_rubros As Integer
Dim vCantidad_grupos As Integer
Private Sub Class_Initialize()
On Error GoTo err22
    Set cn = conectar.obternerConexion 'New ADODB.Connection
'    Set rs = New ADODB.recordset
'    CN.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & serverBBDD & ";uid=root;pwd=3l3c720n;database=sp;OPTION=3;connection=adUseClient"
'    CN.Open

Me.ejecutar "select count(id) as canti from rubros"
vCantidad_rubros = rs!canti

Me.ejecutar "select count(id) as canti from grupos"
vCantidad_grupos = rs!canti
Exit Sub
err22:
    MsgBox Err.Description
End Sub
Public Function ejecutar(strsql)
On Error GoTo err8
'    If rs.State = 1 Then rs.Close
'    rs.Open strsql, CN, adOpenDynamic, adLockOptimistic, adCmdText

Set rs = conectar.RSFactory(strsql)
Exit Function
err8:
MsgBox "Se produjo un error: " & Err.Description
End Function
Public Function Terminar()
    cn.Close
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    Set cn = Nothing
End Function
Public Function rubros_llenarlst()
frmRubrosGrupos.lstRubros.ListItems.Clear
Me.ejecutar "select * from rubros order by rubro"

If vCantidad_rubros > 0 Then
        Dim X As ListItem
        Dim I As Integer
            For I = 1 To vCantidad_rubros
                Set X = frmRubrosGrupos.lstRubros.ListItems.Add(, , Format(rs!id, "000") & " - " & rs!rubro)
                X.SubItems(1) = rs!iniciales
                X.Tag = rs!id
                X.ListSubItems(1).Tag = rs!rubro
                rs.MoveNext
            Next
End If
rs.Close
End Function
 
Public Property Get CantidadRubros() As Integer
CantidadRubros = vCantidad_rubros
End Property
Public Property Get CantidadGrupos() As Integer
CantidadGrupos = vCantidad_grupos
End Property

Public Function grupos_llenarlst(id_rubro)
If id_rubro > 0 Then
frmRubrosGrupos.lstGrupos.ListItems.Clear
Me.ejecutar "select count(id) as canti_grupos from grupos where id_rubro=" & id_rubro
vCantid = rs!canti_grupos
rs.Close

Me.ejecutar "select * from grupos where id_rubro=" & id_rubro & " order by grupo"
If vCantidad_grupos > 0 Then
        Dim X As ListItem
        Dim I As Integer
            For I = 1 To vCantid
                Set X = frmRubrosGrupos.lstGrupos.ListItems.Add(, , rs!grupo)
                X.Tag = rs!id
                rs.MoveNext
            Next
End If
rs.Close
Else
frmRubrosGrupos.lstGrupos.ListItems.Clear
End If
End Function
Public Function llenar_combo_rubros(cbo As ComboBox)
Set rs = New ADODB.recordset
cbo.Clear
'frmRubrosGrupos.cboRubros.Text = "Seleccione el rubro correspondiente"
Me.ejecutar "select * from rubros order by rubro"
While Not rs.EOF
cbo.AddItem rs!rubro
cbo.ItemData(cbo.NewIndex) = rs!id
rs.MoveNext
Wend
rs.Close
If cbo.ListCount > 0 Then
    cbo.ListIndex = 0
End If

 
End Function

Public Function add_rubro(txtRubro As String, inic As String)
On Error GoTo err4
cn.execute "insert into rubros (rubro,iniciales) VALUES ('" & txtRubro & "','" & inic & "')"
Exit Function
err4:
MsgBox "Se produjo un error al cargar los datos!", vbCritical, "Error"
End Function

Public Function add_grupo(txtGrupo As String, id_rubro As Integer)
On Error GoTo err4
cn.execute "insert into grupos (grupo,id_rubro) VALUES ('" & txtGrupo & "',  " & id_rubro & ")"
Exit Function
err4:
MsgBox "Se produjo un error al cargar los datos!", vbCritical, "Error"
End Function


Public Function modif_rubro(idRubro, texto, inicial) As Boolean
On Error GoTo er4
modif_rubro = True
cn.execute "update rubros set rubro='" & Trim(UCase(texto)) & "', iniciales='" & Trim(UCase(inicial)) & "' where id=" & idRubro

Exit Function
er4:
modif_rubro = False
End Function


Public Function modif_grupo(idGrupo, texto, idRubro) As Boolean
On Error GoTo er4
modif_grupo = True
cn.execute "update grupos set grupo='" & Trim(UCase(texto)) & "',id_rubro=" & idRubro & " where id=" & idGrupo

Exit Function
er4:
modif_grupo = False
End Function


Private Sub Class_Terminate()
Set cn = Nothing
Set rs = Nothing
End Sub


