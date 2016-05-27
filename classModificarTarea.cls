VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classModificarTarea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.recordset
Dim vCantidad_rubros As Integer
Dim vCantidad_grupos As Integer
Private Sub Class_Initialize()
    On Error GoTo err22
    Set cn = conectar.obternerConexion  'New ADODB.Connection
    'Set rs = New ADODB.recordset
    'CN.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & serverBBDD & ";uid=root;pwd=3l3c720n;database=sp;OPTION=3;connection=adUseClient"
    ' CN.Open
    Exit Sub
err22:
    MsgBox Err.Description
End Sub


Public Function llenar_form_mod(Id As Integer)

    Set rs = conectar.RSFactory("SELECT * FROM tareas where id=" & Id)


    frmModificarTarea.Text1(0) = rs!Tarea
    frmModificarTarea.lblCant = rs!cantxproc
    frmModificarTarea.lblSector = rs!id_sector

    Set rs = conectar.RSFactory("SELECT * FROM valores_MDO where id_tarea=" & Id)

    frmModificarTarea.Text1(1) = rs!Descripcion
    frmModificarTarea.Text1(2) = rs!Valor



End Function

Public Function llenar_combo_sectores()

    frmModificarTarea.cboSectores.Clear
    Set rs = conectar.RSFactory("select * from sectores")

    While Not rs.EOF
        frmModificarTarea.cboSectores.AddItem rs!Sector
        frmModificarTarea.cboSectores.ItemData(frmModificarTarea.cboSectores.NewIndex) = rs!Id
        rs.MoveNext
    Wend

    If frmModificarTarea.cboSectores.ListCount > 0 Then
        frmModificarTarea.cboSectores.ListIndex = 0
    End If
End Function



Public Function ejecutar_consulta(consulta As String)
    On Error GoTo err14:
    Set rs = conectar.RSFactory(consulta)
    'If rs.State = 1 Then rs.Close
    'rs.Open consulta, CN, adOpenDynamic, adLockOptimistic, adCmdText
    Exit Function
err14:
    MsgBox "Se produjo un error: " & Err.Description

End Function

Private Sub Class_Terminate()
    Set cn = Nothing
    Set rs = Nothing
End Sub
