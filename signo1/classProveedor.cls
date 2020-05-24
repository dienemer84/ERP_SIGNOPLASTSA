VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim CN As ADODB.Connection
Dim Rs As ADODB.Recordset
Dim Cant_Rubros As Integer

Private Sub Class_Initialize()
    Set CN = New ADODB.Connection
    Set Rs = New ADODB.Recordset
    CN.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=192.168.0.1;uid=root;pwd=3l3c720n;database=compras;OPTION=3;connection=adUseClient"
    CN.Open
    
    Rs.Open "select count(id) as canti from rubros", CN, adOpenDynamic, adLockOptimistic, adCmdText
    Cant_Rubros = Rs!canti
    Rs.Close


    
    
End Sub


Function cantidad_total() As Integer
If Rs.State = 1 Then Rs.Close
Rs.Open "select count(id) as canti from stock", CN, adOpenDynamic, adLockOptimistic, adCmdText
cantidad_total = Rs!canti
Rs.Close
End Function
Public Sub llenar_lista_proveedores(lst As ListView, rubro As Integer, proveedor As String, Optional marcado As Integer)
lst.Visible = False
lst.ColumnHeaders.Clear
lst.ColumnHeaders.Add(1).Text = "Numero"
lst.ColumnHeaders(1).Width = 0
lst.ColumnHeaders.Add(2).Text = "Cód"
lst.ColumnHeaders(2).Width = 600
lst.ColumnHeaders(2).Alignment = lvwColumnCenter
lst.ColumnHeaders.Add(3).Text = "Razon Social"
lst.ColumnHeaders(3).Width = 3500
lst.ColumnHeaders(3).Alignment = lvwColumnCenter
lst.ColumnHeaders.Add(4).Text = "Direccion"
lst.ColumnHeaders(4).Width = 3500
lst.ColumnHeaders(4).Alignment = lvwColumnCenter
lst.ColumnHeaders.Add(5).Text = "Ciudad"
lst.ColumnHeaders(5).Width = 1400
lst.ColumnHeaders(5).Alignment = lvwColumnCenter
lst.ColumnHeaders.Add(6).Text = "CP"
lst.ColumnHeaders(6).Width = 1000
lst.ColumnHeaders(6).Alignment = lvwColumnCenter
lst.ColumnHeaders.Add(7).Text = "Telefono"
lst.ColumnHeaders(7).Width = 1800
lst.ColumnHeaders(7).Alignment = lvwColumnCenter
lst.ColumnHeaders.Add(8).Text = "Fax"
lst.ColumnHeaders(8).Width = 1200
lst.ColumnHeaders(8).Alignment = lvwColumnCenter
lst.ColumnHeaders.Add(9).Text = "E-Mail"
lst.ColumnHeaders(9).Width = 2000
lst.ColumnHeaders(9).Alignment = lvwColumnCenter
lst.ColumnHeaders.Add(10).Text = "Contácto"
lst.ColumnHeaders(10).Width = 0
lst.ColumnHeaders(10).Alignment = lvwColumnCenter
lst.ColumnHeaders.Add(11).Text = "Cond. Pago"
lst.ColumnHeaders(11).Width = 2000
lst.ColumnHeaders(11).Alignment = lvwColumnCenter
lst.ColumnHeaders.Add(12).Text = "P.C.E."
lst.ColumnHeaders(12).Width = 700
lst.ColumnHeaders(12).Alignment = lvwColumnCenter
lst.ColumnHeaders.Add(13).Text = "u$S"
lst.ColumnHeaders(13).Width = 700
lst.ColumnHeaders(13).Alignment = lvwColumnCenter
lst.ColumnHeaders.Add(14).Text = "Bonif"
lst.ColumnHeaders(14).Width = 800
lst.ColumnHeaders(14).Alignment = lvwColumnCenter
lst.ColumnHeaders.Add(15).Text = "estado"
lst.ColumnHeaders(15).Width = 0
lst.ColumnHeaders(15).Alignment = lvwColumnCenter

lst.ListItems.Clear


If rubro = -1 Then
'si  son todos los rubros
  If Trim(proveedor) = Empty Then
    ' si no hay filtro
    strsql = "Select * from proveedores where estado=1"
  Else
    'si hay filtro
    strsql = "Select * from proveedores where razon like '" & proveedor & "%' and estado=1"
  End If
ElseIf rubro = -2 Then
'si  son todos los rubros
  If Trim(proveedor) = Empty Then
    ' si no hay filtro
    strsql = "Select * from proveedores where estado=0"
  Else
    'si hay filtro
    strsql = "Select * from proveedores where razon like '" & proveedor & "%' and estado=0"
  End If
  Else
  'si es un rubro especifico
     If Trim(proveedor) = Empty Then
        ' si no hay filtro
        strsql = "select p.* from proveedores p, asignacion a where a.id_rubro =" & rubro & " and a.id_proveedor=p.id and estado=1"
     Else
        'si hay filtro
        strsql = "select p.* from proveedores p, asignacion a where a.id_rubro =" & rubro & " and a.id_proveedor=p.id and razon like '" & proveedor & "%' and estado=1"
  End If
End If
If Rs.State = 1 Then Rs.Close
Rs.Open strsql, CN, adOpenDynamic, adLockOptimistic, adCmdText
Dim X As ListItem
    While Not Rs.EOF
        Set X = lst.ListItems.Add(, , Rs!id)
        X.SubItems(1) = Format(Rs!id, "0000")
        X.SubItems(2) = Rs!razon
        X.SubItems(3) = Rs!direccion
        X.SubItems(4) = Rs!ciudad
        X.SubItems(5) = Rs!cp
        X.SubItems(6) = Rs!tel
        X.SubItems(7) = Rs!fax
        X.SubItems(8) = Rs!email
        X.SubItems(9) = Rs!contacto
        X.SubItems(10) = Rs!FP
        X.SubItems(11) = UCase(Rs!PCE)
        X.SubItems(12) = UCase(Rs!dolar)
        X.SubItems(13) = Rs!bonificacion
        X.SubItems(14) = Rs!estado   '1 significa activo
        X.Tag = Rs!razon
        If X = marcado Then  'deja marcado el q se vio antes
            X.Selected = True
            X.EnsureVisible
        End If
        Rs.MoveNext
    Wend
    lst.Visible = True
End Sub

Public Function llenar_lista_rubros(lst As ListView, marcado As Integer, tamaño As Integer, check As Boolean, Optional idp As Integer)
lst.ColumnHeaders.Clear
lst.ColumnHeaders.Add(1).Text = "" 'dejo el titulo en blanco, por si hay check
tam = 0
lst.Checkboxes = check
If check = True Then tam = 300 'si hay check, dejo espacio para verlo
lst.ColumnHeaders(1).Width = tam 'a
lst.ColumnHeaders.Add(2).Text = "Rubro"
lst.ColumnHeaders(2).Width = tamaño
lst.ColumnHeaders(2).Alignment = lvwColumnCenter
lst.ListItems.Clear

If idp = Empty Then 'la consulta depende si esta o no definido idp (id del proveedor)
strsql = "select * from rubros ORDER BY rubro ASC"
Else
strsql = "select a.id_rubro as id ,r.rubro from asignacion a, rubros r where a.id_proveedor=" & idp & " and a.id_rubro=r.id ORDER BY id_rubro ASC"
End If
If Rs.State = 1 Then Rs.Close
Rs.Open strsql, CN, adOpenDynamic, adLockOptimistic, adCmdText
cant = Cant_Rubros
If cant > 0 Then
Dim X As ListItem
While Not Rs.EOF
     Set X = lst.ListItems.Add(, , Rs!id)
     X.SubItems(1) = Rs!rubro
     If X = marcado Then 'deveria verse marcado el item q queremos
        X.Selected = True
        X.EnsureVisible
     End If
     Rs.MoveNext
Wend
Rs.Close
End If
End Function
Public Function ejecutar_consulta(consulta As String)
    If Rs.State = 1 Then Rs.Close
    Rs.Open consulta, CN, adOpenDynamic, adLockOptimistic, adCmdText
End Function

Public Function llenar_combo_rubros(cbo As ComboBox)

strsql = "select * from rubros ORDER BY rubro ASC"
If Rs.State = 1 Then Rs.Close
Rs.Open strsql, CN, adOpenDynamic, adLockOptimistic, adCmdText
cbo.Clear
cbo.AddItem "Todos"
cbo.ItemData(cbo.NewIndex) = -1
cbo.AddItem "Eliminados"
cbo.ItemData(cbo.NewIndex) = -2

While Not Rs.EOF
cbo.AddItem Rs!rubro
cbo.ItemData(cbo.NewIndex) = Rs!id
Rs.MoveNext
Wend
Rs.Close
If cbo.ListCount > 0 Then
    cbo.ListIndex = 0
End If
End Function

Public Function cargar_datos(listaRubros As ListView, accion As Integer, Optional frm2 As frmModificaProveedor, Optional frm1 As frmNuevoProveedor) As Integer
cargar_datos = 0
Dim frm As Form
If accion = 0 Then
    Set frm = frm1
ElseIf accion = 1 Then
    Set frm = frm2
End If
razon = normaliza(frm.Text1(0))
direccion = normaliza(frm.Text1(1))
ciudad = normaliza(frm.Text1(2))
cp = normaliza(frm.Text1(3))
tel = normaliza(frm.Text1(4))
fax = normaliza(frm.Text1(5))
email = normaliza(frm.Text1(6))
contacto = normaliza(frm.Text1(7))
FP = normaliza(frm.Text1(8))
bonificacion = CDbl(frm.Text1(9))
dolares = ""
PCE = ""
If frm.Check2 Then dolares = "X"
If frm.Check1 Then PCE = "X"

CN.BeginTrans
tra = True
On Error GoTo e

If accion = 0 Then
'doy de alta al proveedor

strsql = "insert into proveedores (razon,direccion,ciudad,cp,tel,fax,email,contacto,FP,PCE,dolar,bonificacion) VALUES "
strsql = strsql & " ('" & razon & "','" & direccion & "','" & ciudad & "','" & cp & "','" & tel & "','" & fax & "','" & email & "','" & contacto & "','" & FP & "','" & PCE & "','" & dolar & "'," & bonificacion & ")"
ElseIf accion = 1 Then
'modifico datos
id = frm.idproveedor
strsql = "update proveedores set razon='" & razon & "', direccion='" & direccion & "', ciudad = '" & ciudad & "', cp='" & cp & "', tel= '" & tel & "', fax='" & fax & "', contacto='" & contacto & "',FP='" & FP & "', PCE='" & PCE & "', dolar='" & dolares & "', bonificacion='" & bonificacion & "' where id =" & id
End If

CN.Execute strsql
If accion = 0 Then
    'si doy de alta, esta rutina y pesco el ultimo id para ingresar los rubros asignados
    
    Rs.Open "select last_insert_id() as idd from proveedores", CN, adOpenDynamic, adLockOptimistic, adCmdTextid
    ultid = Rs!idd
        
ElseIf accion = 1 Then
    'si modifico, esta otra y borro todos los rubros q tenia asignados, para ingresarlos de vuelta
    ultid = id
    CN.Execute "delete from asignacion where id_proveedor=" & ultid
End If
        'cargo todos los rubros nuevos
        For I = 1 To listaRubros.ListItems.Count
            id_rubro = listaRubros.ListItems(I)
            strsql2 = "insert into asignacion (id_proveedor,id_rubro) values (" & ultid & "," & id_rubro & ")"
            CN.Execute strsql2
        Next I

CN.CommitTrans
'If accion = 1 Then Unload frm
'Rs.Close
Exit Function
e:
If tra = True Then
MsgBox Err.Description
CN.RollbackTrans
cargar_datos = -1
End If
End Function

Public Function cambiar_estado(id As Integer, accion As Integer) As Boolean
strsql = "select estado from proveedores where id = " & id
If Rs.State = 1 Then Rs.Close
Rs.Open strsql, CN, adOpenDynamic, adLockOptimistic, adCmdText
c = 0
While Not Rs.EOF
c = c + 1
Rs.MoveNext
Wend
If c = 1 Then
 Rs.MoveFirst
 estado = Rs!estado
If accion = 1 Then 'si manda a deshabilitar
  If estado = 0 Then
     MsgBox "Este proveedor ya fue inhabilitado", vbCritical, "Error"
  Else
   CN.Execute "update proveedores set estado=0 where id=" & id
  End If
ElseIf accion = 0 Then
  If estado = 1 Then
     MsgBox "Este proveedor ya fue habilitado", vbCritical, "Error"
  Else
   CN.Execute "update proveedores set estado=1 where id=" & id
  End If
End If
Else
 MsgBox "Se produjo un error de redundancia de datos", vbCritical, "Error"
End If
End Function


Public Function llenar_para_modificar(frm As frmModificaProveedor)
id = CInt(frm.idproveedor)
strsql = "select * from proveedores where id=" & id
If Rs.State = 1 Then Rs.Close
Rs.Open strsql, CN, adOpenDynamic, adLockOptimistic, adCmdText
c = 0
While Not Rs.EOF
c = c + 1
Rs.MoveNext
Wend

If c = 1 Then

Rs.MoveFirst
razon = Rs!razon
direccion = Rs!direccion
ciudad = Rs!ciudad
cp = Rs!cp
tel = Rs!tel
fax = Rs!fax
email = Rs!email
contacto = Rs!contacto
FP = Rs!FP
If UCase(Trim(Rs!PCE)) = "X" Then frm.Check1.value = 1
If UCase(Trim(Rs!dolar)) = "X" Then frm.Check2.value = 1
bonificacion = Rs!bonificacion
frm.Text1(0) = razon
frm.Text1(1) = direccion
frm.Text1(2) = ciudad
frm.Text1(3) = cp
frm.Text1(4) = tel
frm.Text1(5) = fax
frm.Text1(6) = email
frm.Text1(7) = contacto
frm.Text1(8) = FP
frm.Text1(9) = bonificacion
End If


End Function
