Attribute VB_Name = "DAOObraSocial"
Option Explicit

Public Function Save(os As ObraSocial) As Boolean
    On Error GoTo err1
 
    conectar.BeginTransaction
    Dim q As String
    If os.id = 0 Then    'el insert se hace classPersonal
        q = "INSERT INTO ObraSocial " _
            & " (id," _
            & " nombre)"
        q = q & " Values" _
            & " ('id'," _
            & " 'nombre')"


        q = Replace$(q, "'id'", Escape(os.id))
        q = Replace$(q, "'nombre'", Escape(os.nombre))
        


    Else

        q = "Update ObraSocial" _
            & " SET" _
            & " nombre = " & Escape(os.nombre) _
            & " Where id = " & Escape(os.id)

    End If

    Save = conectar.execute(q)
    If Save Then If os.id = 0 Then os.id = conectar.UltimoId2

    conectar.CommitTransaction
    Exit Function
err1:
    conectar.RollBackTransaction

End Function






Public Sub llenarComboXtremeSuite(cbo As Xtremesuitecontrols.ComboBox)
    Dim col As Collection
    Set col = GetAll
    Dim o As ObraSocial
    cbo.Clear
    
    Dim i As Integer
    
    For i = 1 To col.count
        Set o = col(i)
    
            cbo.AddItem o.nombre
        
        cbo.ItemData(cbo.NewIndex) = o.id
    Next i
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub

Public Function GetById(id As Integer) As ObraSocial
    
    Set GetById = GetAll("ObraSocial.id=" & id)(1)
End Function
Public Function GetDefault() As ObraSocial
    
    Set GetDefault = GetAll("ObraSocial.es_default=1")(1)
End Function


Public Function GetAll(Optional filter As String = vbNullString) As Collection
    Dim col As New Collection
    Dim A As ObraSocial
    Dim rs As Recordset
    If LenB(filter) = 0 Then filter = "1=1"
    Set rs = conectar.RSFactory("SELECT * FROM ObraSocial  where 1 = 1 and " & filter)
    While Not rs.EOF
        Dim fieldsIndex As Dictionary
        BuildFieldsIndex rs, fieldsIndex
   
    
        col.Add Map(rs, fieldsIndex, "ObraSocial")
        rs.MoveNext
    Wend
    Set GetAll = col
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As ObraSocial
    Dim E As ObraSocial
    Dim id As Long

    id = GetValue(rs, indice, tabla, "id")

    If id <> 0 Then
        Set E = New ObraSocial
        E.id = id
        E.nombre = GetValue(rs, indice, tabla, "nombre")
        E.default = GetValue(rs, indice, tabla, "es_default")

    End If

    Set Map = E
End Function
