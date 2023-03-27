Attribute VB_Name = "DAOCaja"
Option Explicit

Public Function FindById(Id As Long) As caja
    Dim col As Collection
    Set col = FindAll("c.id = " & Id)
    If col.count = 0 Then
        Set FindById = Nothing
    Else
        Set FindById = col.item(1)
    End If
End Function

Public Function FindAll(Optional ByVal filter As String = " 1 = 1 ") As Collection
    Dim q As String
    q = "SELECT *" _
      & " FROM cajas c WHERE " & filter

    Dim col As New Collection
    Dim rs As Recordset
    Set rs = conectar.RSFactory(q)
    Dim fieldsIndex As New Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Dim tmp As caja

    While Not rs.EOF
        Set tmp = Map(rs, fieldsIndex, "c")
        col.Add tmp, CStr(tmp.Id)
        rs.MoveNext
    Wend

    Set FindAll = col
End Function

Public Sub llenarComboXtremeSuite(cbo As Xtremesuitecontrols.ComboBox)
    Dim col As Collection
    Set col = FindAll
    Dim i As Integer
    Dim caja As caja
    cbo.Clear
    For i = 1 To col.count
        Set caja = col(i)
        cbo.AddItem caja.nombre
        cbo.ItemData(cbo.NewIndex) = caja.Id
    Next i
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As caja
    Dim c As caja
    Dim Id As Long

    Id = GetValue(rs, indice, tabla, "id")

    If Id > 0 Then
        Set c = New caja
        c.Id = Id
        c.nombre = GetValue(rs, indice, tabla, "nombre")
    End If

    Set Map = c
End Function

