Attribute VB_Name = "DAOPercepciones"
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection


Public Function GetAll() As Collection
    Dim col As New Collection
    Dim A As clsPercepciones

    Set rs = conectar.RSFactory("select * from AdminConfigPercepciones")

    While Not rs.EOF


        Set A = New clsPercepciones
        A.Id = rs!Id
        A.Percepcion = rs!Percepcion
        A.Porcentaje = rs!Porcentaje
        A.valido = rs!valido

        col.Add A

        rs.MoveNext
    Wend
    Set A = Nothing
    Set GetAll = col
End Function
Public Function GetById(id_percepcion As Long) As clsPercepciones
    Dim A As clsPercepciones
    Set rs = conectar.RSFactory("select * from AdminConfigPercepciones where id=" & id_percepcion)
    If Not rs.EOF And Not rs.BOF Then
        Set A = New clsPercepciones
        A.Id = rs!Id
        A.Percepcion = rs!Percepcion
        A.Porcentaje = rs!Porcentaje
        A.valido = rs!valido
    Else
        Set A = Nothing
    End If
    Set GetById = A

End Function




Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As clsPercepciones

    Dim Id As Long: Id = GetValue(rs, indice, tabla, "id")
    Dim P As clsPercepciones

    If Id > 0 Then
        Set P = New clsPercepciones
        P.Id = Id
        P.Porcentaje = GetValue(rs, indice, tabla, "porcentaje")
        P.valido = GetValue(rs, indice, tabla, "Valido")
        P.codigo = GetValue(rs, indice, tabla, "Codigo")
        P.Percepcion = GetValue(rs, indice, tabla, "Percepcion")
    End If
    Set Map = P
End Function


