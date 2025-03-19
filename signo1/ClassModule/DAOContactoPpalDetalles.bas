Attribute VB_Name = "DAOContactoPpalDetalles"
Option Explicit
Public Const TABLA_AGENDA As String = "a"
Public Const TABLA_DETALLE_AGENDA As String = "da"

Public Const CAMPO_ID As String = "id"
Public Const CAMPO_IDCONTACTO As String = "id_agenda"
Public Const CAMPO_DETALLE As String = "detalle"
Public Const CAMPO_TELEFONO1 As String = "tel1"
Public Const CAMPO_TELEFONO2 As String = "tel2"
Public Const CAMPO_MAIL As String = "mail"
Public Const CAMPO_MAS As String = "mas"


Public Function Guardar(deta As clsContactoPpalDetalle) As Boolean
    On Error GoTo err1
    Guardar = True
    Dim q As String
    Dim n As Boolean
    If deta.Id = 0 Then
        n = True
        q = " INSERT INTO datos_agenda (" _
          & "id_agenda," _
          & "detalle, " _
          & "tel1, " _
          & "tel2, " _
          & "mail, " _
          & "mas ) " _
          & "Values  ( " _
          & conectar.Escape(deta.IdAgenda) & "," _
          & conectar.Escape(deta.detalle) & "," _
          & conectar.Escape(deta.Telefono1) & "," _
          & conectar.Escape(deta.Telefono2) & "," _
          & conectar.Escape(deta.mail) & "," _
          & conectar.Escape(deta.Mas) & ")"

    Else
        n = False
        q = "UPDATE datos_agenda SET " _
          & "id_agenda = 'id_agenda' , " _
          & "detalle = 'detalle' , " _
          & "tel1 = 'tel1' , " _
          & "tel2 = 'tel2' , " _
          & "mail = 'mail' , " _
          & "mas = 'mas' " _
          & "WHERE  id = 'id' "

        q = Replace$(q, "'id'", conectar.Escape(deta.Id))
        q = Replace$(q, "'id_agenda'", conectar.Escape(deta.IdAgenda))
        q = Replace$(q, "'detalle'", conectar.Escape(deta.detalle))
        q = Replace$(q, "'tel1'", conectar.Escape(deta.Telefono1))
        q = Replace$(q, "'tel2'", conectar.Escape(deta.Telefono2))
        q = Replace$(q, "'mail'", conectar.Escape(deta.mail))
        q = Replace$(q, "'mas'", conectar.Escape(deta.Mas))

    End If
    If Not conectar.execute(q) Then GoTo err1

    Exit Function
err1:
    Guardar = False
End Function

Public Function Delete(detalle As clsContactoPpalDetalle) As Boolean

    Delete = conectar.execute("delete from datos_agenda where id=" & detalle.Id)


' Stop


End Function

Public Function FindAllByContactoPpal(idContacto As Long) As Collection
    Set FindAllByContactoPpal = FindAll("AND " & TABLA_DETALLE_AGENDA & "." & CAMPO_IDCONTACTO & "=" & idContacto)
End Function


Public Function FindAll(Optional filtro As String = vbNullString, Optional WithCantidadEntregadas As Boolean = False, Optional WithDetallePedido As Boolean = False) As Collection
    Dim indice As Dictionary
    Dim rs As Recordset
    Dim col As New Collection
    Dim strsql As String
    strsql = "SELECT * FROM datos_agenda da LEFT JOIN agenda a ON da.id_agenda = a.id WHERE 1=1 "
    
    If LenB(filtro) > 0 Then strsql = strsql & filtro
    
    Set rs = conectar.RSFactory(strsql)
    
    conectar.BuildFieldsIndex rs, indice
    
    Dim detalle As clsContactoPpalDetalle

    While Not rs.EOF
        Set detalle = New clsContactoPpalDetalle

        Set detalle = Map(rs, indice)

        col.Add detalle
        rs.MoveNext
    Wend

    Set FindAll = col
    
End Function


Public Function Map(rs As Recordset, indice As Dictionary) As clsContactoPpalDetalle
    Dim Id As Variant
    Id = GetValue(rs, indice, TABLA_DETALLE_AGENDA, CAMPO_ID)
    If Id > 0 Then
        Dim dr As clsContactoPpalDetalle
        Set dr = New clsContactoPpalDetalle
        
        dr.IdAgenda = GetValue(rs, indice, TABLA_DETALLE_AGENDA, CAMPO_IDCONTACTO)
        dr.detalle = GetValue(rs, indice, TABLA_DETALLE_AGENDA, CAMPO_DETALLE)
        dr.Telefono1 = GetValue(rs, indice, TABLA_DETALLE_AGENDA, CAMPO_TELEFONO1)
        dr.Telefono2 = GetValue(rs, indice, TABLA_DETALLE_AGENDA, CAMPO_TELEFONO2)
        dr.mail = GetValue(rs, indice, TABLA_DETALLE_AGENDA, CAMPO_MAIL)
        dr.Mas = GetValue(rs, indice, TABLA_DETALLE_AGENDA, CAMPO_MAS)
        'dr.Id = GetValue(rs, indice, TABLA_DETALLE_AGENDA, CAMPO_ID)
        dr.Id = Id

        Set Map = dr
    End If

End Function









