Attribute VB_Name = "DAOLocalidades"
Option Explicit



Public Function FindAll( _
    Optional ByVal filtro As String = vbNullString _
) As Collection

    On Error GoTo ManejarError

    Dim idx As Dictionary
    Dim rs As Recordset
    Dim strsql As String
    Dim col As Collection
    Dim objLocalidad As localidad

    Set idx = New Dictionary
    Set col = New Collection

    strsql = _
        "SELECT * " & _
        "FROM Localidades l " & _
        "INNER JOIN Provincia p " & _
        "ON l.idProvincia = p.id " & _
        "INNER JOIN Pais pa " & _
        "ON p.idPais = pa.id " & _
        "WHERE 1 = 1 "

    If LenB(filtro) > 0 Then
        strsql = strsql & filtro
    End If

    strsql = strsql & " ORDER BY l.Nombre"

    Set rs = conectar.RSFactory(strsql)

    If rs Is Nothing Then
        Err.Raise _
            vbObjectError + 1200, _
            "DAOLocalidades.FindAll", _
            "No se pudo abrir el Recordset de localidades."
    End If

    conectar.BuildFieldsIndex rs, idx

    Do While Not rs.EOF

        Set objLocalidad = Map( _
                                rs, _
                                idx, _
                                "l", _
                                "p", _
                                "pa" _
                            )

        If Not objLocalidad Is Nothing Then
            col.Add objLocalidad
        End If

        rs.MoveNext

    Loop

    Set FindAll = col
    Exit Function

ManejarError:

    Dim numeroError As Long
    Dim descripcionError As String

    numeroError = Err.Number
    descripcionError = Err.Description

    Err.Raise _
        numeroError, _
        "DAOLocalidades.FindAll", _
        descripcionError

End Function

Public Function FindById(Id As Long) As localidad
    Dim col As New Collection
    Set col = FindAll("And l.id=" & Id)
    Set FindById = col(1)
End Function

Public Function FindAllByProvincia( _
    ByVal idProvincia As Long _
) As Collection

    Set FindAllByProvincia = FindAll( _
        " AND p.id = " & CStr(idProvincia) _
    )

End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, _
                    Optional tablaProv As String = vbNullString, _
                    Optional tablaPais As String = vbNullString _
                  ) As localidad

    Dim loc As localidad
    Dim Id As Long: Id = GetValue(rs, indice, tabla, "ID")

    If Id > 0 Then
        Set loc = New localidad
        loc.Id = Id
        loc.nombre = GetValue(rs, indice, tabla, "Nombre")
        loc.cp = GetValue(rs, indice, tabla, "CP")
        If LenB(tablaProv) > 0 Then Set loc.provincia = DAOProvincias.Map(rs, indice, tablaProv, tablaPais)

    End If

    Set Map = loc
End Function



Public Sub LlenarCombo( _
    ByRef cbo As XtremeSuiteControls.ComboBox, _
    ByVal idProvincia As Long _
)

    On Error GoTo ManejarError

    Dim L As localidad
    Dim col As Collection

    cbo.Clear

    Set col = FindAllByProvincia(idProvincia)

    If col Is Nothing Then
        Err.Raise _
            vbObjectError + 1201, _
            "DAOLocalidades.LlenarCombo", _
            "No se pudo obtener la colección de localidades."
    End If

    For Each L In col

        If Not L Is Nothing Then

            cbo.AddItem L.nombre
            cbo.ItemData(cbo.NewIndex) = L.Id

        End If

    Next L

    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If

    Exit Sub

ManejarError:

    Dim numeroError As Long
    Dim descripcionError As String

    numeroError = Err.Number
    descripcionError = Err.Description

    Err.Raise _
        numeroError, _
        "DAOLocalidades.LlenarCombo", _
        descripcionError

End Sub

Public Function Save(L As localidad) As Boolean
    Dim q As String
    On Error GoTo err1
    If L.Id > 0 Then

        q = "UPDATE sp.Localidades  SET  CP='" & L.cp & "', idProvincia=" & L.provincia.Id & ", Nombre = '" & UCase(L.nombre) & "'   WHERE   ID = '" & L.Id & "' "
    Else
        q = "INSERT INTO sp.Localidades (Nombre,idProvincia,CP)VALUES('" & UCase(L.nombre) & "'," & L.provincia.Id & ",'" & L.cp & "')"
    End If

    If Not conectar.execute(q) Then GoTo err1
    Exit Function
err1:
    Save = False
End Function
