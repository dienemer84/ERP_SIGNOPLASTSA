VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classSectores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim vCantidad As Integer
Public Function LlenarListaSectores(marcado As Integer)
    Dim strsql As String
    frmSectores.lstSectores.ListItems.Clear
    strsql = "select * from sectores"
    Me.ejecutar strsql
    'rs.Open strsql, CN, adOpenDynamic, adLockOptimistic, adCmdText
    Cant = vCantidad
    If Cant > 0 Then
        While Not rs.EOF
            Set x = frmSectores.lstSectores.ListItems.Add(, , rs!Id)
            x.SubItems(1) = rs!Sector
            If x = marcado Then
                x.Selected = True
                x.EnsureVisible

            End If

            rs.MoveNext
        Wend
        rs.Close
    End If
End Function
Public Sub ejecutar(consulta As String)
    On Error GoTo err7
    '   If rs.State = 1 Then rs.Close
    '   rs.Open consulta, CN, adOpenDynamic, adLockOptimistic, adCmdText
    Set rs = conectar.RSFactory(consulta)
    Exit Sub
err7:
    MsgBox "Se produjo un error: " & Err.Description

End Sub

Private Sub Class_Terminate()
    Set cn = Nothing
    Set rs = Nothing
End Sub
