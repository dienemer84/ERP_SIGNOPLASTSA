Attribute VB_Name = "DAOProvincias"
Option Explicit



Public Function FindAll( _
    Optional ByVal filtro As String = vbNullString _
) As Collection

    On Error GoTo ManejarError

    Dim idx As Dictionary
    Dim rs As Recordset
    Dim strsql As String
    Dim col As Collection
    Dim objProvincia As provincia

    Set idx = New Dictionary
    Set col = New Collection

    strsql = _
        "SELECT * " & _
        "FROM Provincia p " & _
        "INNER JOIN Pais pa ON p.idPais = pa.id " & _
        "WHERE 1 = 1 "

    If LenB(filtro) > 0 Then
        strsql = strsql & filtro
    End If

    Set rs = conectar.RSFactory(strsql)

    If rs Is Nothing Then
        Err.Raise _
            vbObjectError + 1100, _
            "DAOProvincias.FindAll", _
            "No se pudo abrir el Recordset de provincias."
    End If

    conectar.BuildFieldsIndex rs, idx

    Do While Not rs.EOF

        Set objProvincia = Map(rs, idx, "p", "pa")

        If Not objProvincia Is Nothing Then
            col.Add objProvincia
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
        "DAOProvincias.FindAll", _
        descripcionError

End Function

'''Public Function FindAllByPais(idpais As Long) As Collection
'''    Set FindAllByPais = FindAll("and pa.id=" & idpais)
'''End Function

Public Function FindAllByPais(idPais As Long) As Collection
    Set FindAllByPais = FindAll( _
        " AND pa.id=" & idPais & _
        " ORDER BY CASE WHEN UCASE(p.Nombre) = 'EXTERIOR' THEN 1 ELSE 0 END, p.Nombre")
End Function


Public Function FindById(idProvincia As Long) As provincia

    Dim col As New Collection
    Set col = FindAll("and p.id=" & idProvincia)


    Set FindById = col(1)
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, Optional tablaPais As String) As provincia

    Dim prov As provincia
    Dim Id As Long: Id = GetValue(rs, indice, tabla, "ID")

    If Id > 0 Then
        Set prov = New provincia
        prov.Id = Id
        prov.nombre = GetValue(rs, indice, tabla, "Nombre")
        If LenB(tablaPais) > 0 Then Set prov.pais = DAOPais.Map(rs, indice, tablaPais)
    End If

    Set Map = prov
End Function

Public Sub LlenarCombo( _
    ByRef cbo As XtremeSuiteControls.ComboBox, _
    ByVal idPais As Long _
)

    On Error GoTo ManejarError

    Dim P As provincia
    Dim col As Collection

    cbo.Clear

    Set col = FindAllByPais(idPais)

    If col Is Nothing Then
        Err.Raise _
            vbObjectError + 1101, _
            "DAOProvincias.LlenarCombo", _
            "No se pudo obtener la colección de provincias."
    End If

    For Each P In col

        If Not P Is Nothing Then

            cbo.AddItem P.nombre
            cbo.ItemData(cbo.NewIndex) = P.Id

        End If

    Next P

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
        "DAOProvincias.LlenarCombo", _
        descripcionError

End Sub


Public Function LlenarComboNoDefinido(cbo As XtremeSuiteControls.ComboBox, Id As Long, Optional incluirSinDefinir As Boolean = True)
    Dim P As provincia
    Dim col As Collection

    cbo.Clear
    cbo.Sorted = False
    
    If incluirSinDefinir Then
        cbo.AddItem "SIN DEFINIR"
        cbo.ItemData(cbo.NewIndex) = 0
    End If

    Set col = FindAllByPais(Id)

    For Each P In col
        If IsSomething(P) Then
            cbo.AddItem P.nombre
            cbo.ItemData(cbo.NewIndex) = P.Id
        End If
    Next

    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Function


Public Function Save(P As provincia) As Boolean
    Dim q As String
    On Error GoTo err1
    Dim n As Boolean
    If P.Id > 0 Then

        q = "UPDATE sp.Provincia  SET  idPais=" & P.pais.Id & ", Nombre = '" & UCase(P.nombre) & "'   WHERE   ID = '" & P.Id & "' "
        n = False
    Else
        q = "INSERT INTO sp.Provincia (Nombre,idPais)VALUES('" & UCase(P.nombre) & "'," & P.pais.Id & ")"
        n = True
    End If

    If Not conectar.execute(q) Then GoTo err1
    If n Then P.Id = conectar.UltimoId2

    Exit Function
err1:
    Save = False
End Function
