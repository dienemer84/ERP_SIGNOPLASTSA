Attribute VB_Name = "DAOBancos"
Option Explicit

Public Const CAMPO_ID As String = "id"
Public Const CAMPO_NOMBRE As String = "Nombre"
Public Const TABLA_BANCO As String = "bco"
Private Banco As Banco
Dim rs As Recordset

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As Banco
    Dim Banco As Banco
    Set Banco = New Banco
    Dim Id As Long
    Id = GetValue(rs, indice, tabla, CAMPO_ID)
    If Id > 0 Then
        Banco.Id = Id
        Banco.nombre = GetValue(rs, indice, tabla, CAMPO_NOMBRE)

    End If
    Set Map = Banco
End Function


Public Function GetById(Id As Long) As Banco
    Dim col As Collection
    Set col = GetAll(DAOBancos.CAMPO_ID & "=" & Id)
    If col.count = 0 Then
        Set GetById = Nothing
    Else
        Set GetById = col(1)
    End If
End Function

Public Function GetAll(Optional filtro As String = Empty) As Collection

    On Error GoTo err1
    Dim col As New Collection
    Dim bco As Banco
    Dim indice As Dictionary
    Dim q As String

    q = "select * from AdminConfigBancos bco where 1=1"

    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If
    Set rs = conectar.RSFactory(q)

    conectar.BuildFieldsIndex rs, indice
    While Not rs.EOF
        Set bco = New Banco
        Set bco = Map(rs, indice, TABLA_BANCO)
        col.Add bco, CStr(bco.Id)
        rs.MoveNext
    Wend

    Set GetAll = col
    Exit Function
err1:
    Set GetAll = Nothing
End Function


Public Sub llenarComboXtremeSuite(cbo As XtremeSuiteControls.ComboBox)
    Dim col As New Collection
    Set col = DAOBancos.GetAll()
    Dim bco As Banco
    cbo.Clear
    For Each bco In col
        cbo.AddItem bco.nombre
        cbo.ItemData(cbo.NewIndex) = bco.Id
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub


Public Function Save(Banco As Banco) As Boolean
    On Error GoTo err1
    Dim q As String
    Save = True
    Dim n As Boolean
    If Banco.Id = 0 Then
        q = "INSERT INTO sp.AdminConfigBancos  (Nombre)  Values  ('Nombre')"
        n = True
    Else
        q = "UPDATE sp.AdminConfigBancos set Nombre='Nombre' where id='id'"
        n = False
    End If
    q = Replace(q, "'Nombre'", Escape(Banco.nombre))
    q = Replace(q, "'id'", Escape(Banco.Id))

    Save = conectar.execute(q)
    If n Then Banco.Id = conectar.UltimoId2
    Exit Function
err1:
    Save = False
End Function
