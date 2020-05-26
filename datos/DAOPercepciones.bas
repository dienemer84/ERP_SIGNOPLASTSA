Attribute VB_Name = "DAOPercepciones"
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection


Public Function GetAll() As Collection
    Dim col As New Collection
    Dim a As clsPercepciones

    Set rs = conectar.RSFactory("select * from AdminConfigPercepciones")

    While Not rs.EOF


        Set a = New clsPercepciones
        a.id = rs!id
        a.Percepcion = rs!Percepcion
        a.Porcentaje = rs!Porcentaje
        a.valido = rs!valido

        col.Add a

        rs.MoveNext
    Wend
    Set a = Nothing
    Set GetAll = col
End Function
Public Function GetById(id_percepcion As Long) As clsPercepciones
    Dim a As clsPercepciones
    Set rs = conectar.RSFactory("select * from AdminConfigPercepciones where id=" & id_percepcion)
    If Not rs.EOF And Not rs.BOF Then
        Set a = New clsPercepciones
        a.id = rs!id
        a.Percepcion = rs!Percepcion
        a.Porcentaje = rs!Porcentaje
        a.valido = rs!valido
    Else
        Set a = Nothing
    End If
    Set GetById = a

End Function




Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As clsPercepciones

    Dim id As Long: id = GetValue(rs, indice, tabla, "id")
    Dim P As clsPercepciones

    If id > 0 Then
        Set P = New clsPercepciones
        P.id = id
        P.Porcentaje = GetValue(rs, indice, tabla, "porcentaje")
        P.valido = GetValue(rs, indice, tabla, "Valido")
        P.codigo = GetValue(rs, indice, tabla, "Codigo")
        P.Percepcion = GetValue(rs, indice, tabla, "Percepcion")
    End If
    Set Map = P
End Function


