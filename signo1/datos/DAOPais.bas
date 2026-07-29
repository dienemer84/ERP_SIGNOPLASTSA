Attribute VB_Name = "DAOPais"
Option Explicit

Public Function FindById(ByVal Id As Long) As pais

    Dim col As Collection

    Set col = FindAll(" AND pa.id = " & CStr(Id))

    If Not col Is Nothing Then

        If col.count > 0 Then
            Set FindById = col.item(1)
        End If

    End If

End Function

Public Function FindAll( _
    Optional ByVal filtro As String = "" _
) As Collection

    Dim paso As String
    Dim numeroError As Long
    Dim descripcionError As String

    Dim idx As Dictionary
    Dim rs As Recordset
    Dim strsql As String
    Dim col As Collection
    Dim objPais As pais

    On Error GoTo ManejarError

    paso = "Crear colección"

    Set col = New Collection
    Set idx = New Dictionary

    paso = "Preparar consulta SQL"

    strsql = "SELECT * FROM Pais pa WHERE 1 = 1"

    If LenB(filtro) > 0 Then
        strsql = strsql & filtro
    End If

    paso = "Abrir Recordset. SQL: " & strsql

    Set rs = conectar.RSFactory(strsql)

    If rs Is Nothing Then

        Err.Raise _
            vbObjectError + 1000, _
            "DAOPais.FindAll", _
            "RSFactory devolvió un Recordset vacío."

    End If

    paso = "Construir índice de campos"

    conectar.BuildFieldsIndex rs, idx

    paso = "Recorrer países"

    Do While Not rs.EOF

        paso = "Mapear país"

        Set objPais = Map(rs, idx, "pa")

        If Not objPais Is Nothing Then
            col.Add objPais
        End If

        rs.MoveNext

    Loop

    Set FindAll = col
    Exit Function

ManejarError:

    numeroError = Err.Number
    descripcionError = Err.Description

    Err.Raise _
        numeroError, _
        "DAOPais.FindAll - " & paso, _
        descripcionError

End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As pais

    Dim pais As pais
    Dim Id As Long: Id = GetValue(rs, indice, tabla, "ID")

    If Id > 0 Then
        Set pais = New pais
        pais.Id = Id
        pais.nombre = GetValue(rs, indice, tabla, "Nombre")

    End If

    Set Map = pais
End Function

Public Sub LlenarCombo( _
    ByRef cbo As XtremeSuiteControls.ComboBox _
)

    Dim paso As String
    Dim numeroError As Long
    Dim descripcionError As String

    Dim P As pais
    Dim col As Collection

    On Error GoTo ManejarError

    paso = "Obtener colección de países"

    Set col = FindAll()

    If col Is Nothing Then

        Err.Raise _
            vbObjectError + 1001, _
            "DAOPais.LlenarCombo", _
            "FindAll devolvió Nothing."

    End If

    paso = "Limpiar combo"

    cbo.Clear

    paso = "Agregar países al combo"

    For Each P In col

        If Not P Is Nothing Then

            cbo.AddItem P.nombre
            cbo.ItemData(cbo.NewIndex) = P.Id

        End If

    Next P

    paso = "Seleccionar primer país"

    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If

    Exit Sub

ManejarError:

    numeroError = Err.Number
    descripcionError = Err.Description

    Err.Raise _
        numeroError, _
        "DAOPais.LlenarCombo - " & paso, _
        descripcionError

End Sub


Public Function Save(pais As pais) As Boolean
    Dim q As String
    Dim n As Boolean
    n = False
    On Error GoTo err1
    If pais.Id > 0 Then
        q = "UPDATE sp.Pais  SET   Nombre = '" & UCase(pais.nombre) & "'   WHERE   ID = '" & pais.Id & "' "
    Else
        q = "INSERT INTO sp.Pais (Nombre)VALUES('" & UCase(pais.nombre) & "')"
        n = True
    End If

    If Not conectar.execute(q) Then GoTo err1
    
    If n Then
        pais.Id = conectar.UltimoId2
    End If
    
    Save = True
    Exit Function
err1:
    Save = False
End Function
