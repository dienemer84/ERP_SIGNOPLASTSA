VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classNuevaMDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Private Sub Class_Initialize()
    On Error GoTo err22
    Set cn = conectar.obternerConexion
    Exit Sub
err22:
    MsgBox Err.Description
End Sub
Public Function ejecutar_consulta(consulta As String)
    On Error GoTo err12
    Set rs = conectar.RSFactory(consulta)
    Exit Function
err12:
    MsgBox "Se produjo un error: " & Err.Description
End Function
Public Function VERMDO(codi As Integer, ByRef cantxproc, ByRef Sector, ByRef Tarea, ByRef mdoDescrip)
    Dim strsql As String
    strsql = "select v.id_tarea,v.valor, v.fecha, v.descripcion,s.sector,t.tarea,t.id_sector,t.cantxproc from tareas t inner join sectores s on t.id_sector=s.id inner join  valores_MDO v on v.id_tarea=t.id and t.id=" & codi
    Me.ejecutar_consulta strsql
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend
    If c = 1 Then
        rs.MoveFirst
        Sector = rs!Sector
        cantxproc = rs!cantxproc
        Sector = rs!Sector
        Tarea = rs!Tarea
        mdoDescrip = rs!descripcion
        Valor = rs!Valor
    End If
End Function
Private Sub Class_Terminate()
    Set cn = Nothing
    Set rs = Nothing
End Sub
