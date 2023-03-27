Attribute VB_Name = "DAOActualizar"
Option Explicit


Public Function FindAll(Optional ByVal filter As String = vbNullString) As Collection
    Dim q As String
    q = "SELECT * FROM ActualizacionSistemaLog WHERE 1 = 1 ORDER BY id DESC"

    Dim col As New Collection
    Dim rs As Recordset
    Set rs = conectar.RSFactory(q)
    Dim fieldsIndex As New Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Dim sin As Actualizacion


    While Not rs.EOF
        Set sin = Map(rs, fieldsIndex, "ActualizacionSistemaLog", "id")
        col.Add sin, CStr(sin.Id_)

        rs.MoveNext
    Wend

    Set FindAll = col

End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, Optional Id As String = vbNullString) As Actualizacion

    Dim s As Actualizacion

    Id = GetValue(rs, indice, tabla, "id")

    If Id > 0 Then
        Set s = New Actualizacion

        s.Id_ = Id

        s.Id_ = GetValue(rs, indice, tabla, "id")
        s.Fecha_ = GetValue(rs, indice, tabla, "fecha")
        s.Detalle_ = GetValue(rs, indice, tabla, "detalle")
        s.Modulo_ = GetValue(rs, indice, tabla, "sector")

    End If

    Set Map = s
End Function

Public Function CargarNuevoDetalle(Nota As clsNotas) As Boolean

    On Error GoTo E

    Dim strsql As String

    conectar.BeginTransaction

    strsql = "INSERT INTO ActualizacionSistemaLog (fecha, sector, detalle) VALUES ('" & funciones.datetimeFormateada(Nota.FechaD_) & "', '" & Nota.Modulo_ & "','" & Nota.TextoD_ & "')"
    If Not conectar.execute(strsql) Then GoTo E

    conectar.CommitTransaction

    CargarNuevoDetalle = True
    Exit Function
E:
    CargarNuevoDetalle = False
    conectar.RollBackTransaction

End Function


