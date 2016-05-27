VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classModificaMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim strsql As String
Dim CN As ADODB.Connection
Dim rs As ADODB.Recordset
Dim vCantidad_rubros As Integer
Dim vCantidad_grupos As Integer
Private Sub Class_Initialize()
 On Error GoTo err22
    Set CN = New ADODB.Connection
    Set rs = New ADODB.Recordset
    CN.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & serverBBDD & ";uid=root;pwd=3l3c720n;database=sp;OPTION=3;connection=adUseClient"
    CN.Open
    
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


Public Function llenar_combo_rubros()
strsql = "select * from rubros"
Me.ejecutar_consulta strsql
frmModificaMaterial.cboRubros.Clear

While Not rs.EOF

frmModificaMaterial.cboRubros.AddItem rs!rubro
frmModificaMaterial.cboRubros.ItemData(frmModificaMaterial.cboRubros.NewIndex) = rs!id
rs.MoveNext
Wend
rs.Close

If frmModificaMaterial.cboRubros.ListCount > 0 Then
    frmModificaMaterial.cboRubros.ListIndex = 0
End If

End Function

Public Function ejecutar_consulta(Consulta As String)
On Error GoTo err22
If rs.State = 1 Then rs.Close
rs.Open Consulta, CN, adOpenDynamic, adLockOptimistic, adCmdText
Exit Function
err22:
MsgBox "Se produjo un error: " & Err.Description

End Function




Private Sub Class_Terminate()
Set CN = Nothing
Set rs = Nothing
End Sub
