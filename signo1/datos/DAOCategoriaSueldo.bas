Attribute VB_Name = "DAOCategoriaSueldo"
Option Explicit

Public Const CAMPO_ID As String = "id"
Public Const CAMPO_NOMBRE As String = "nombre"
Public Const CAMPO_VALOR As String = "valor"
Public Const CAMPO_PORCENTAJE_ESPECIALIZACION As String = "porcentaje_especializacion"

Public Const TABLA_CATEGORIA_SUELDO As String = "cs"


Public Function FindAll(Optional ByRef filter As String = vbNullString) As Collection
    Dim tickStart As Double
    Dim tickend As Double
    tickStart = GetTickCount

    Dim rs As ADODB.Recordset
    Dim q As String

    Dim categorias As New Collection

    q = "SELECT cs.*, mon.* FROM categoria_sueldo cs LEFT JOIN AdminConfigMonedas mon on cs.id_moneda=mon.id WHERE 1 = 1"

    If LenB(filter) > 0 Then
        q = q & " AND " & filter
    End If

    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim catSueldo As CategoriaSueldo

    While Not rs.EOF
        Set catSueldo = DAOCategoriaSueldo.Map(rs, fieldsIndex, DAOCategoriaSueldo.TABLA_CATEGORIA_SUELDO, "mon")
        categorias.Add catSueldo, CStr(catSueldo.Id)
        rs.MoveNext
    Wend

    tickend = GetTickCount

    '    Debug.Print tickEnd - tickStart, "ms elapsed"

    Set FindAll = categorias
End Function

Public Function Map(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, ByRef tableNameOrAlias As String, ByRef tablaMoneda As String) As CategoriaSueldo
    Dim c As CategoriaSueldo
    Dim Id As Variant

    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCategoriaSueldo.CAMPO_ID)

    If Id > 0 Then
        Set c = New CategoriaSueldo
        c.Id = Id
        c.nombre = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCategoriaSueldo.CAMPO_NOMBRE)
        c.PorcentajeEspecializacion = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCategoriaSueldo.CAMPO_PORCENTAJE_ESPECIALIZACION)
        c.Valor = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCategoriaSueldo.CAMPO_VALOR)
        If LenB(tablaMoneda) > 0 Then Set c.moneda = DAOMoneda.Map(rs, fieldsIndex, tablaMoneda)
    End If

    Set Map = c
End Function

Public Function Save(catSueldo As CategoriaSueldo) As Boolean
    On Error GoTo E
    Save = True
    Dim q As String

    If catSueldo.Id = 0 Then

        q = "INSERT INTO categoria_sueldo" _
          & " (nombre, " _
          & " valor," _
          & " porcentaje_especializacion)" _
          & " VALUES (" & conectar.Escape(UCase(catSueldo.nombre)) & ", " _
          & conectar.Escape(catSueldo.Valor) & "," _
          & conectar.Escape(catSueldo.PorcentajeEspecializacion) & ")"
    Else
        q = "update categoria_sueldo" _
          & " SET" _
          & " nombre = " & conectar.Escape(UCase(catSueldo.nombre)) & "," _
          & " valor = " & conectar.Escape(catSueldo.Valor) & "," _
          & " porcentaje_especializacion = " & conectar.Escape(catSueldo.PorcentajeEspecializacion) _
          & " Where" _
          & " id = " & catSueldo.Id
    End If

    If conectar.execute(q) Then

        If catSueldo.Id = 0 Then
            conectar.UltimoId "categoria_sueldo", catSueldo.Id
        End If

        q = "INSERT INTO categoria_sueldo_historico (id_categoria_sueldo, valor, fecha, id_usuario)   VALUES " _
          & "(" & catSueldo.Id & "," & catSueldo.Valor & "," & conectar.Escape(Now) & "," & funciones.getUser & ")"
        Save = conectar.execute(q)



    Else
        GoTo E
    End If
    Exit Function
E:
    Save = False
End Function

Public Function Delete(catSueldo As CategoriaSueldo) As Boolean
    On Error GoTo E
    Delete = conectar.execute("DELETE FROM categoria_sueldo WHERE id = " & catSueldo.Id)
    Exit Function
E:
    Delete = False
End Function
