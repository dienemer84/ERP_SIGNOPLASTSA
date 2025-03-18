Attribute VB_Name = "DAOContactoPpal"
Option Explicit

Public Const CAMPO_ID As String = "id"
Public Const CAMPO_EMPRESA As String = "empresa"
Public Const CAMPO_DIRECCION As String = "direccion"
Public Const CAMPO_LOCALIDAD As String = "localidad"
Public Const CAMPO_EMAIL As String = "email"
Public Const TABLA_AGENDA As String = "a"


Public Function FindAll(Optional ByRef filter As String = vbNullString, Optional ByRef order As String = vbNullString) As Collection
    Dim rs As ADODB.Recordset
    Dim q As String

    Dim contactos As New Collection

    q = "SELECT * FROM agenda a WHERE 1 = 1"

    If LenB(filter) > 0 Then
        q = q & "" & filter
    End If

    If LenB(order) > 0 Then
        q = q & " ORDER BY a.empresa"
    End If

    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Set FindAll = New Collection


    While Not rs.EOF
        contactos.Add Map(rs, fieldsIndex, TABLA_AGENDA)
        rs.MoveNext
    Wend

    Set FindAll = contactos
    
    
End Function



Public Function Map(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, ByRef tableNameOrAlias As String) As clsContactoPpal
    
    Dim C As clsContactoPpal
    Dim Id As Variant

    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ID)

    If Id > 0 Then
        Set C = New clsContactoPpal
        C.Id = Id
        C.Empresa = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_EMPRESA)
        C.direccion = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_DIRECCION)
        C.localidad = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_LOCALIDAD)
        C.Email = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_EMAIL)

    End If

    Set Map = C
End Function


Public Function Save(C As clsContactoPpal, Optional Cascade As Boolean = False, Optional NotificarObserver As Boolean = True) As Boolean
    conectar.BeginTransaction
    Save = Guardar(C, Cascade, NotificarObserver)
    
    If Not Save Then GoTo err1
    conectar.CommitTransaction
    Exit Function
err1:
    Save = False
    conectar.RollBackTransaction
End Function


Public Function Guardar(C As clsContactoPpal, Optional Cascade As Boolean = False, Optional NotificarObserver As Boolean = True) As Boolean
    
    Dim q As String
    
    Guardar = True

    Dim Nueva As Boolean
    If C.Id = 0 Then
        Nueva = True

        q = "INSERT INTO agenda (empresa, direccion, localidad, email) Values (" _
          & conectar.Escape(C.Empresa) & ", " _
          & conectar.Escape(C.direccion) & ", " _
          & conectar.Escape(C.localidad) & ", " _
          & conectar.Escape(C.Email) & ")"

    Else
        Nueva = False
        q = "UPDATE agenda " _
          & "SET " _
          & "empresa = " & conectar.Escape(C.Empresa) & " ," _
          & "direccion = " & conectar.Escape(C.direccion) & " ," _
          & "localidad = " & conectar.Escape(C.localidad) & " ," _
          & "email = " & conectar.Escape(C.Email) & "" _
          & " WHERE " _
          & "id = " & C.Id

    End If
    If Not conectar.execute(q) Then GoTo err1

    If C.Id = 0 Then
        C.Id = conectar.UltimoId2
    End If
    
    If Cascade Then
        If Not conectar.execute("DELETE FROM datos_agenda WHERE id_agenda=" & C.Id) Then GoTo err1

        Dim deta As clsContactoPpalDetalle
        
        For Each deta In C.Detalles
            deta.Id = 0
            deta.IdAgenda = C.Id
            If Not DAOContactoPpalDetalles.Guardar(deta) Then GoTo err1
        Next

    End If
    
    Exit Function

err1:
    Guardar = False
End Function

