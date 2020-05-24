VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatosxxxxxxxxx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim CN As ADODB.Connection
Dim rs As ADODB.recordset
Dim vConsulta As String
Private Sub Class_Initialize()
On Error GoTo err22
    Set CN = New ADODB.Connection
    Set rs = New ADODB.recordset
    CN.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & serverBBDD & ";uid=root;pwd=3l3c720n;database=sp;OPTION=3;connection=adUseClient"
    CN.Open

Exit Sub
err22:
    MsgBox "Se produjo un error: " & Err.Description
    Err.Clear
End Sub
Private Sub Class_Terminate()
    CN.Close
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    Set CN = Nothing
End Sub
Public Function ejecutarConsulta(strsql) As Boolean
On Error GoTo err8
ejecutar = True
    If rs.State = 1 Then rs.Close
    rs.Open strsql, CN, adOpenDynamic, adLockOptimistic, adCmdText
Exit Function
err8:
ejecutar = False
MsgBox "Se produjo un error: " & Err.Description
Err.Clear
End Function
Public Function ejecutarComando(comando) As Boolean
On Error GoTo err45
ejecutarComando = True
CN.Execute comando
Exit Function
err45:
ejecutarComando = False
MsgBox "Se produjo un error: " & Err.Description
Err.Clear
End Function
Public Function CrearRS(Consulta As String) As recordset
On Error GoTo err1
Dim rs As New recordset
'If rs.State = 1 Then rs.Close
'rs.CursorLocation = adUseClient
'rs.Open Consulta, CN, adOpenDynamic, adLockOptimistic, adCmdText
Me.ejecutarConsulta Consulta
Set CrearRS = rs
Exit Function
err1:
Set CrearRS = Nothing
MsgBox "Se produjo un error: " & Err.Description
Err.Clear
End Function




