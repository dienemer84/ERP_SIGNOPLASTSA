Attribute VB_Name = "DAOProveedor"
Option Explicit
Dim Proveedor As clsProveedor


Public Function FindAllProveedoresWithFacturasImpagas() As Collection    'of proveedor
    Set FindAllProveedoresWithFacturasImpagas = FindAll("proveedores.id IN (SELECT DISTINCT id_proveedor from AdminComprasFacturasProveedores WHERE estado = " & EstadoFacturaProveedor.pagoParcial & " or estado = " & EstadoFacturaProveedor.Aprobada & ")")
End Function


Public Function Map2(ByRef rs As Recordset, Index As Dictionary, ByRef tableNameOrAlias As String, _
                     Optional ByRef AdminConfigIVAProveedor As String = vbNullString, _
                     Optional ByRef AdminConfigFacturasProveedor As String = vbNullString) As clsProveedor
    Dim P As clsProveedor
    Dim Id As Long: Id = GetValue(rs, Index, tableNameOrAlias, "id")

    If Id > 0 Then
        Set P = New clsProveedor
        P.Id = Id
        P.razonFantasia = GetValue(rs, Index, tableNameOrAlias, "razon_fantasia")
        P.IIBB = GetValue(rs, Index, tableNameOrAlias, "iibb")
        P.Ciudad = GetValue(rs, Index, tableNameOrAlias, "ciudad")
        P.contacto = GetValue(rs, Index, tableNameOrAlias, "contacto")
        P.cp = GetValue(rs, Index, tableNameOrAlias, "cp")
        P.direccion = GetValue(rs, Index, tableNameOrAlias, "direccion")
        P.Email = GetValue(rs, Index, tableNameOrAlias, "email")
        P.estado = GetValue(rs, Index, tableNameOrAlias, "estado")
        P.Fax = GetValue(rs, Index, tableNameOrAlias, "fax")
        P.FormaPago = GetValue(rs, Index, tableNameOrAlias, "FP")
        P.Cuit = GetValue(rs, Index, tableNameOrAlias, "cuit")
        P.pagocontraEntrega = GetValue(rs, Index, tableNameOrAlias, "PCE")
        P.pagoDolares = GetValue(rs, Index, tableNameOrAlias, "DOLAR")
        P.bonificacion = GetValue(rs, Index, tableNameOrAlias, "bonificacion")
        P.RazonSocial = GetValue(rs, Index, tableNameOrAlias, "razon")
        P.tel = GetValue(rs, Index, tableNameOrAlias, "tel")

        P.CBU = GetValue(rs, Index, tableNameOrAlias, "cbu")
        P.ALIAS = GetValue(rs, Index, tableNameOrAlias, "alias")
        P.TitularCta = GetValue(rs, Index, tableNameOrAlias, "titularcta")
        
        '    Set P.Moneda = DAOMoneda.GetById(GetValue(rs, Index, tableNameOrAlias, "id_moneda"))

        Set P.moneda = DAOMoneda.Map(rs, Index, "AdminConfigMonedas")
    End If
    Set Map2 = P
End Function

Public Function FindAll(Optional filtro As String = vbNullString, _
                        Optional WithContactos As Boolean = False, _
                        Optional WithCuentasContables As Boolean = False, _
                        Optional WithAlicuotas As Boolean = False, _
                        Optional EstadoCtaCte As Boolean = True, _
                        Optional EstadoContado As Boolean = True, _
                        Optional EstadoEliminado As Boolean = False, _
                        Optional WithRubros As Boolean = False, _
                        Optional WithMoneda As Boolean = False) As Collection
    Dim col As New Collection
    Dim cuenta As clsCuentaContable
    Dim rs As Recordset
    Dim rubros As String: rubros = vbNullString
    Dim cuentas As String: cuentas = vbNullString
    Dim asignacion As String: asignacion = vbNullString
    Dim cuentaProv As String: cuentaProv = vbNullString
    Dim contacto As String: contacto = vbNullString
    Dim alicuotas As String: alicuotas = vbNullString
    Dim tipIva As clsTipoIvaProveedor
    Dim q As String
    q = "SELECT * from sp.proveedores " _
      & "LEFT JOIN sp.AdminConfigIVAProveedor        ON (proveedores.id_iva = AdminConfigIVAProveedor.id) " _
      & "LEFT JOIN sp.AdminConfigFacturasProveedor   ON (AdminConfigFacturasProveedor.id_iva = AdminConfigIVAProveedor.id) " _

If WithCuentasContables Then
        cuentas = " AdminComprasCuentasContables"
        cuentaProv = " AdminComprasCuentasProveedores"
        q = q & " LEFT JOIN sp.AdminComprasCuentasProveedores ON (AdminComprasCuentasProveedores.id_proveedor = proveedores.id) " _
          & " LEFT JOIN sp.AdminComprasCuentasContables   ON (AdminComprasCuentasProveedores.id_cuenta = AdminComprasCuentasContables.id) "
    End If
    If WithRubros Then
        rubros = "rubros"
        asignacion = "asignacion"
        q = q & " LEFT JOIN sp.asignacion  ON (asignacion.id_proveedor = proveedores.id) " _
          & " LEFT JOIN sp.rubros  ON (asignacion.id_rubro = rubros.id) "
    End If
    If WithContactos Then
        contacto = "contactos"
        q = q & " LEFT JOIN sp.contactos    ON (contactos.idCliente = proveedores.id) "
    End If
    If WithAlicuotas Then
        alicuotas = "AdminConfigIvaAlicuotas"
        q = q & " LEFT JOIN sp.AdminConfigIvaAlicuotas        ON (AdminConfigIvaAlicuotas.id_config_factura = AdminConfigFacturasProveedor.id) "
    End If

    'moneda
    q = q & " LEFT JOIN sp.AdminConfigMonedas        ON (proveedores.id_moneda= AdminConfigMonedas.id) "


    q = q & " WHERE 1=1"
    If LenB(filtro) > 0 Then
        q = q & " AND " & filtro
    End If


    If EstadoCtaCte Or EstadoContado Or EstadoEliminado Then
        q = q & " and proveedores.estado IN (-1"
    End If

    If EstadoCtaCte Then
        q = q & "," & EstadoProveedor.EstadoProveedorCuentaCorriente
    End If

    If EstadoContado Then
        q = q & "," & EstadoProveedor.EstadoProveedorContado
    End If
    If EstadoEliminado Then
        q = q & ", " & EstadoProveedor.EstadoProveedorEliminado
    End If

    If EstadoCtaCte Or EstadoContado Or EstadoEliminado Then
        q = q & ")"
    End If

    Set rs = conectar.RSFactory(q)
    
    Dim indice As New Dictionary
    Dim prov As clsProveedor
    BuildFieldsIndex rs, indice
    Dim cont As clsContacto
    Dim rub As clsRubros
    Dim ali As clsAlicuotas
    Dim confac As clsConfigFacturaProveedor

    While Not rs.EOF


        Set prov = Map2(rs, indice, "proveedores", "AdminConfigIVAProveedor", "AdminConfigFacturasProveedor")
        If funciones.BuscarEnColeccion(col, CStr(prov.Id)) Then
            Set prov = col.item(CStr(prov.Id))
        Else
            col.Add prov, CStr(prov.Id)
        End If

        If WithContactos Then
            Set cont = DAOContacto.Map(rs, indice, contacto)
            If IsSomething(cont) Then
                If Not funciones.BuscarEnColeccion(prov.contactos, CStr(cont.Id)) Then
                    prov.contactos.Add cont, CStr(cont.Id)
                End If
            End If
        End If

        If WithRubros Then
            Set rub = DAORubros.Map(rs, indice, "rubros")
            If IsSomething(rub) Then
                If Not funciones.BuscarEnColeccion(prov.rubros, CStr(rub.Id)) Then
                    prov.rubros.Add rub, CStr(rub.Id)
                End If
            End If
        End If
        Set tipIva = DAOTipoIvaProveedor.Map(rs, indice, "AdminConfigIVAProveedor")
        If IsSomething(tipIva) And Not IsSomething(prov.TipoIVA) Then
            Set prov.TipoIVA = tipIva
        End If

        If IsSomething(prov.TipoIVA) Then
            Set confac = DAOConfigFacturaProveedor.Map(rs, indice, "AdminConfigFacturasProveedor")
            If IsSomething(confac) Then

                If funciones.BuscarEnColeccion(prov.TipoIVA.configFacturas, CStr(confac.Id)) Then
                    Set confac = prov.TipoIVA.configFacturas.item(CStr(confac.Id))
                Else
                    prov.TipoIVA.configFacturas.Add confac, CStr(confac.Id)
                End If

                If WithAlicuotas Then
                    Set ali = DAOAlicuotas.Map(rs, indice, "AdminConfigIvaAlicuotas")
                    If IsSomething(ali) Then
                        If Not funciones.BuscarEnColeccion(confac.alicuotas, CStr(ali.Id)) Then
                            confac.alicuotas.Add ali, CStr(ali.Id)
                        End If
                    End If
                End If

                If WithCuentasContables Then
                    Set cuenta = DAOCuentaContable.Map(rs, indice, "AdminComprasCuentasContables")

                    If IsSomething(cuenta) Then
                        If Not funciones.BuscarEnColeccion(prov.cuentasContables, CStr(cuenta.Id)) Then
                            prov.cuentasContables.Add cuenta, CStr(cuenta.Id)
                        End If
                    End If
                End If
            End If
        End If
        rs.MoveNext
    Wend
    Set FindAll = col
End Function


Public Function FindAllByRubro(rubroId As Long) As Collection
    Set FindAllByRubro = FindAll("proveedores.id IN (SELECT DISTINCT id_proveedor FROM asignacion WHERE asignacion.id_rubro = " & rubroId & ")", , , , True)
End Function


Public Function FindById(Id As Long, _
                         Optional WithContactos As Boolean = True, _
                         Optional WithCuentasContables As Boolean = True, _
                         Optional WithAlicuotas As Boolean = True, _
                         Optional WithRubros As Boolean = True _
                       ) As clsProveedor
    On Error GoTo err1
    Set FindById = DAOProveedor.FindAll("proveedores.id = " & Id, WithContactos, WithCuentasContables, WithAlicuotas, , , , WithRubros).item(1)
    Exit Function
err1:
    Set FindById = Nothing
End Function

Public Function Save(Proveedor As clsProveedor) As Boolean
    On Error GoTo err1
    conectar.BeginTransaction
    If Not Guardar(Proveedor) Then GoTo err1
    conectar.CommitTransaction
    Save = True
    Exit Function
err1:
    Save = False
    conectar.RollBackTransaction
End Function

Public Function Guardar(Proveedor As clsProveedor) As Boolean

    On Error GoTo E
    Dim strsql As String
    Dim rs As Recordset
    Dim n As Boolean
    If Proveedor.Id = 0 Then
        n = True
        strsql = "insert into proveedores (id_moneda,razon,direccion,ciudad,cp,tel,fax,email,contacto,FP,PCE,dolar,bonificacion,id_iva,razon_fantasia,iibb,cuit,estado,cbu,alias,titularcta) VALUES "
        If Not IsSomething(Proveedor.moneda) Then

            Set Proveedor.moneda = DAOMoneda.GetById(0)
        End If

        strsql = strsql & " ( " & Proveedor.moneda.Id & "," & conectar.Escape(Proveedor.RazonSocial) & "," & conectar.Escape(Proveedor.direccion) & "," & conectar.Escape(Proveedor.Ciudad) & "," & conectar.Escape(Proveedor.cp) & "," & conectar.Escape(Proveedor.tel) & "," & conectar.Escape(Proveedor.Fax) & "," & conectar.Escape(Proveedor.Email) & "," & conectar.Escape(Proveedor.contacto) & "," & conectar.Escape(Proveedor.FormaPago) & "," & conectar.Escape(Proveedor.pagocontraEntrega) & "," & conectar.Escape(Proveedor.pagoDolares) & "," & conectar.Escape(Proveedor.bonificacion) & "," & conectar.Escape(Proveedor.TipoIVA.Id) & "," & conectar.Escape(Proveedor.razonFantasia) & "," & conectar.Escape(Proveedor.IIBB) & "," & conectar.Escape(Proveedor.Cuit) & "," & conectar.Escape(Proveedor.estado) & "," & conectar.Escape(Proveedor.CBU) & "," & conectar.Escape(Proveedor.ALIAS) & "," & conectar.Escape(Proveedor.TitularCta) & ")"
        If Not conectar.execute(strsql) Then GoTo E

        Set rs = conectar.RSFactory("select last_insert_id() as idd from proveedores")
        Dim ultid As Long
        ultid = rs!idd
        Proveedor.Id = ultid
        'cargo todos los rubros nuevos
        Dim ru As clsRubros
        Dim F As Long
        For F = 1 To Proveedor.rubros.count
            Set ru = New clsRubros
            If Not conectar.execute("insert into asignacion (id_proveedor,id_rubro) values (" & ultid & "," & ru.Id & ")") Then GoTo E
        Next F
        Proveedor.Id = ultid
    Else
        n = False
        strsql = "update proveedores set " & _
         "id_moneda = " & Proveedor.moneda.Id & ", " & _
         "id_iva = " & conectar.Escape(Proveedor.TipoIVA.Id) & ", " & _
         "cuit = " & conectar.Escape(Proveedor.Cuit) & ", " & _
         "razon = " & conectar.Escape(Proveedor.RazonSocial) & ", " & _
         "direccion = " & conectar.Escape(Proveedor.direccion) & ", " & _
         "ciudad = " & conectar.Escape(Proveedor.Ciudad) & ", " & _
         "cp = " & conectar.Escape(Proveedor.cp) & ", " & _
         "tel = " & conectar.Escape(Proveedor.tel) & ", " & _
         "fax = " & conectar.Escape(Proveedor.Fax) & ", " & _
         "contacto = " & conectar.Escape(Proveedor.contacto) & ", " & _
         "FP = " & conectar.Escape(Proveedor.FormaPago) & ", " & _
         "PCE = " & conectar.Escape(Proveedor.pagocontraEntrega) & ", " & _
         "dolar = " & conectar.Escape(Proveedor.pagoDolares) & ", " & _
         "bonificacion = " & conectar.Escape(Proveedor.bonificacion) & ", " & _
         "razon_fantasia = " & conectar.Escape(Proveedor.razonFantasia) & ", " & _
         "iibb = " & conectar.Escape(Proveedor.IIBB) & ", " & _
         "estado = " & conectar.Escape(Proveedor.estado) & ", " & _
         "email = " & conectar.Escape(Proveedor.Email) & ", " & _
         "cbu = " & conectar.Escape(Proveedor.CBU) & ", " & _
         "alias = " & conectar.Escape(Proveedor.ALIAS) & ", " & _
         "titularcta = " & conectar.Escape(Proveedor.TitularCta) & " " & _
         "where id = " & conectar.Escape(Proveedor.Id)
        
                
        If Not conectar.execute(strsql) Then GoTo E
        If Not conectar.execute("delete from asignacion where id_proveedor=" & Proveedor.Id) Then GoTo E
        'cargo todos los rubros nuevos

        For F = 1 To Proveedor.rubros.count
            Set ru = New clsRubros
            Set ru = Proveedor.rubros(F)
            If Not conectar.execute("insert into asignacion (id_proveedor,id_rubro) values (" & Proveedor.Id & "," & ru.Id & ")") Then GoTo E
        Next F

    End If
    Guardar = True
    Dim EVENTO As New clsEventoObserver
    Set EVENTO.Elemento = Proveedor
    If n Then
        EVENTO.EVENTO = agregar_
    Else
        EVENTO.EVENTO = modificar_
    End If
    EVENTO.tipo = Proveedores_
    Channel.Notificar EVENTO, Proveedores_
    Exit Function



E:
    Guardar = False
    If n Then Proveedor.Id = 0
    'MsgBox Err.Description

End Function
Public Function CambiarEstado(Proveedor As clsProveedor) As Boolean
    On Error GoTo err1
    CambiarEstado = True
    If Proveedor.estado = 1 Then
        conectar.execute "update proveedores set estado=0 where id=" & Proveedor.Id
    ElseIf Proveedor.estado = 0 Then
        conectar.execute "update proveedores set estado=1 where id=" & Proveedor.Id
    End If
    Exit Function
err1:
    CambiarEstado = False
End Function


Public Sub LlenarCombo(cbo As ComboBox, _
                       Optional EstadoCtaCte As Boolean = False, _
                       Optional EstadoContado As Boolean = True, _
                       Optional EstadoEliminado As Boolean = False, Optional rubroId As Long = 0)
    Dim col As Collection
    Dim F As String
    F = vbNullString
    If rubroId <> 0 Then
        F = "proveedores.id IN (SELECT id_proveedor FROM asignacion WHERE id_rubro = " & rubroId & ")"
    End If

    Set col = DAOProveedor.FindAll(F, , , , EstadoCtaCte, EstadoContado, EstadoEliminado)
    Dim prov As clsProveedor
    cbo.Clear
    Dim i As Integer
    For i = 1 To col.count
        Set prov = col(i)
        cbo.AddItem prov.RazonSocial
        cbo.ItemData(cbo.NewIndex) = prov.Id
    Next i
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub


Public Sub llenarComboXtremeSuite(cbo As Xtremesuitecontrols.ComboBox, _
                                  Optional EstadoCtaCte As Boolean = False, _
                                  Optional EstadoContado As Boolean = True, _
                                  Optional EstadoEliminado As Boolean = False)
    Dim col As Collection
    Set col = DAOProveedor.FindAll(, , , , EstadoCtaCte, EstadoContado, EstadoEliminado)
    Dim prov As clsProveedor
    cbo.Clear
    Dim i As Integer
    For i = 1 To col.count
        Set prov = col(i)
        cbo.AddItem prov.RazonSocial
        cbo.ItemData(cbo.NewIndex) = prov.Id
    Next i
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub


Public Function ValidarCuit(Proveedor As clsProveedor) As Boolean
    Dim q As String
    Dim rs As Recordset

    If Proveedor.Cuit = "999999999999999" Then
        ValidarCuit = True
        Exit Function
    End If


    If Proveedor.TipoIVA.detalle = "Exterior" Then
        ValidarCuit = True
        Exit Function
    End If

    q = "select count(id) as cantidad from proveedores where cuit=" & Proveedor.Cuit

    If Proveedor.Id <> 0 Then q = q & " and proveedores.id <> " & Proveedor.Id
    Set rs = conectar.RSFactory(q)
    If Not rs.EOF And Not rs.BOF Then
        If rs!Cantidad >= 0 Then
            ValidarCuit = funciones.VerificarCUIT(Proveedor.Cuit)
        End If
    End If

    Exit Function
err4:

End Function

