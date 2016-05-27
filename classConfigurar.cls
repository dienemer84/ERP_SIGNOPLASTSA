VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classConfigurar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim strsql As String
Dim rs As ADODB.Recordset
Dim vcodigoMaterial As String

Public Function QueIdMaterial(codigos As String) As Long
    Dim rs1 As New Recordset
    On Error GoTo err1

    strsql = "select id  from materiales where codigo=" & conectar.Escape(UCase(codigos))
    Set rs1 = conectar.RSFactory(strsql)
    c = 0
    While Not rs1.EOF
        c = c + 1
        rs1.MoveNext
    Wend
    If c = 1 Then
        rs1.MoveFirst
        QueIdMaterial = rs1!id  'si exist codigo, dvuelvo id
    Else
        QueIdMaterial = -1    'sino devuelvo -1
    End If

    Set rs1 = Nothing
    Exit Function
err1:
    QueIdMaterial = -1
End Function
Private Sub Class_Initialize()
    On Error GoTo err22
    Set cn = conectar.obternerConexion  'New ADODB.Connection
    Exit Sub
err22:
    MsgBox Err.Description
End Sub





Property Get codigoMaterial() As String
    codigoMaterial = rs!codigo
End Property
Property Get manteOferta() As Long
    manteOferta = rs!manteOferta
End Property
Property Get descripcionMaterial() As String
    descripcionMaterial = rs!descripcion
End Property
Public Function rubro_terminacion() As Integer
    strsql = "select rubro from terminacion_cuentas"
    Me.ejecutar_consulta strsql
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend
    If c = 1 Then
        rs.MoveFirst
        rubro_terminacion = rs!Rubro
    Else
        MsgBox "Se produjo un error interno en las BBDD", vbCritical, "Error"
    End If
    ', cboCantPint As ComboBox, cboFosfatos As ComboBox, rubro As Integer
End Function


Public Function sector_terminacion() As Integer
    strsql = "select sector from terminacion_cuentas"
    Me.ejecutar_consulta strsql
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend

    If c = 1 Then
        rs.MoveFirst
        sector_terminacion = rs!Sector
    Else
        MsgBox "Se produjo un error interno en las BBDD", vbCritical, "Error"
    End If
    ', cboCantPint As ComboBox, cboFosfatos As ComboBox, rubro As Integer
End Function



Function ver_datos_pintura(ByRef cantPintM2, ByRef CantForfatosM2, ByRef TpoPrepSupM2, ByRef TpoPinturaM2, ByRef TpoHorneado, ByRef CantKg, ByRef factorAumentoMat, ByRef factotAumentoMDO, ByRef Espesor)
    strsql = "select ID as id,count(id) as cantidad from Terminacion group by id"
    Me.ejecutar strsql

    If rs!Cantidad = 1 Then
        rs.MoveFirst
        id = rs!id
        'selecciono
        strsql = "select * from Terminacion where id=" & id
        Me.ejecutar strsql
        cantPintM2 = rs!cantPintM2
        CantForfatosM2 = rs!CantForfatosM2
        TpoPrepSupM2 = rs!TpoPrepSupM2
        TpoPinturaM2 = rs!TpoPinturaM2
        TpoHorneado = rs!TpoHorneado
        'CantKg = rs!CantKg
        factorAumentoMat = rs!factorAumentoMat
        factotAumentoMDO = rs!factotAumentoMDO
        Espesor = rs!Espesor
    Else
        'creo por default
        strsql = "insert into Terminacion values(1,0,0,0,0,0,0,0,0,0)"
        Me.ejecutar strsql
    End If
End Function
Function DEVOLVER_AUMENTO(m2 As Double, aumento As Double) As Double
    aum = aumento
    au = 1
    If m2 >= 0.75 Then
        c = 1
    ElseIf m2 >= 0.5 Then
        c = 2
    ElseIf m2 >= 0.25 Then
        c = 3
    ElseIf m2 >= 0.125 Then
        c = 4
    ElseIf m2 >= 0.05 Then
        c = 5
    Else
        c = 6
    End If

    If c <> 1 Then
        For i = 1 To c
            au = au * aum
        Next i
    End If
    DEVOLVER_AUMENTO = au
End Function


Public Sub LlenarListaGastos(lista As ListView)
    lista.ListItems.Clear    '
    strsql = "select * from gastos"
    ejecutar strsql
    While Not rs.EOF
        Set x = lista.ListItems.Add(, , rs!id)
        x.SubItems(1) = rs!Concepto
        x.SubItems(2) = rs!porcentual & "%"
        x.SubItems(3) = rs!porcentual
        rs.MoveNext
    Wend
    rs.Close
End Sub
'Public Function mostrar_conf(conf As Integer) As Double
'    If conf = 1 Then    'muestro MU mdo
'        strsql = "Select PorcMO as a from configuracion limit 1"
'    ElseIf conf = 2 Then
'        strsql = "Select PorMAMenos10 as a from configuracion limit 1"
'    ElseIf conf = 3 Then
'        strsql = "Select PorMAMenos15 as a from configuracion limit 1"
'    ElseIf conf = 4 Then
'        strsql = "Select PorMAMas15 as a from configuracion limit 1"
'    End If
'    Me.ejecutar strsql
'    c = 0
'    While Not rs.EOF
'        c = c + 1
'        rs.MoveNext
'    Wend
'    If c = 1 Then
'        rs.MoveFirst
'        mostrar_conf = rs!a
'    End If
'
'End Function
Public Function Gastos() As Double
    strsql = "select sum(porcentual) as totales from gastos"
    ejecutar strsql
    Gastos = rs!totales
End Function
Public Sub ejecutar(consulta)
    On Error GoTo err15
    Set rs = conectar.RSFactory(consulta)

    Exit Sub
err15:
    MsgBox "Se produjo un error: " & Err.Description
End Sub
Public Sub LlenarDatosVarios(frm As frmDatosVarios)

    strsql = "Select * from configuracion where id=0"
    Set rs = conectar.RSFactory(strsql)

    With rs
        frm.txtMOM = rs!Mano_obra_muerta
        frm.txtManteOferta = rs!manteOferta
        frm.txtMDO = rs!PorcMO
        frm.txtPintura = rs!PintM2
        frm.txtMenos10 = rs!PorMAMenos10
        frm.txtMenos15 = rs!PorMAMenos15
        frm.txtMas15 = rs!PorMaMas15
    End With
End Sub

Public Sub Calcular_terminacion(lst As ListView, aumMdo As Double, cantPintM2 As Double, Espesor As Double, ByRef cantPintREF As Double, ByRef cantFosfREF, fosfatos As Double, tpoPrep As Double, ByRef supREF As Double, ByRef aplicPintREF, tiempoAplic As Double, ByRef hornoREF As Double, tpoHorno As Double, Optional largoTot As Double, Optional anchoTot As Double)
    'Cantidad de pintura
    Dim m2 As Double, Largo As Double, Ancho As Double
    For i = 1 To lst.ListItems.count
        m2 = CDbl(lst.ListItems(i).ListSubItems(3))
        canti = CDbl(lst.ListItems(i).ListSubItems(7))
        aumento = Me.DEVOLVER_AUMENTO(m2, aumMdo)
        Largo = CDbl(lst.ListItems(i).ListSubItems(8))
        largoTot = largoTot + Largo
        Ancho = CDbl(lst.ListItems(i).ListSubItems(9))
        anchoTot = anchoTot + Ancho
        caras = CDbl(lst.ListItems(i).ListSubItems(5))
        capas = CDbl(lst.ListItems(i).ListSubItems(6))
        cantPintREF = ((((Largo * Ancho * caras * capas) / 1000000) / 1000 * (cantPintM2 * aumento * Espesor * 2)) * canti) + cantPintREF
        ';D112*E112/1000000*D18
        cantFosfREF = (((((Largo * Ancho) / 1000000) / 1000 * fosfatos * aumento)) * canti) + cantFosfREF    'reemplazar 1 x canti si no funca


        supREF = ((((Largo * Ancho / 1000000) * tpoPrep * aumento)) * canti) + supREF
        aplicPintREF = ((((Largo * Ancho * caras * capas) / 1000000) * (tiempoAplic * aumento)) * canti) + aplicPintREF
        hornoREF = ((((Largo * Ancho / 1000000)) * (capas * caras) * tpoHorno) * canti) + hornoREF
    Next i
End Sub
Public Sub VerParaModificar(id As Integer, frm As frmConfigurarGastos)
    If rs.State = 1 Then rs.Close
    strsql = "select * from gastos where id=" & id
    rs.Open strsql, cn, adOpenDynamic, adLockOptimistic, adCmdText
    frm.Text4 = rs!Concepto
    frm.Text3 = rs!porcentual
End Sub
Public Function ejecutar_consulta(consulta As String)
    On Error GoTo err34
    Set rs = conectar.RSFactory(consulta)
    'If rs.State = 1 Then rs.Close
    'rs.Open consulta, CN, adOpenDynamic, adLockOptimistic, adCmdText
    Exit Function
err34:
    MsgBox "Se produjo un error: " & Err.Description
End Function

Public Sub llenarCboMonedas(cbo As ComboBox)
    'cbo.AddItem "AR$"
    'cbo.ItemData(cbo.NewIndex) = 0
    'cbo.AddItem "US$", 1
    'cbo.ItemData(cbo.NewIndex) = 1
    'cbo.ListIndex = 0
    Me.ejecutar "select nombre_corto,id from AdminConfigMonedas"
    cbo.Clear
    pri = -999
    While Not rs.EOF
        cbo.AddItem rs!Nombre_corto
        cbo.ItemData(cbo.NewIndex) = rs!id
        If pri = -999 Then pri = rs!id
        rs.MoveNext
    Wend
    If pri <> -999 Then
        cbo.ListIndex = pri
    End If
End Sub
Private Sub Class_Terminate()
    Set cn = Nothing
    Set rs = Nothing
End Sub
Public Function ejecutar_comando(strComando As String) As Boolean
    On Error GoTo err545
    ejecutar_comando = True
    cn.BeginTrans
    cn.execute strComando
    cn.CommitTrans
    Exit Function
err545:
    cn.RollbackTrans
    ejecutar_comando = False
End Function
Public Function llenarComboAlmacenes(cbo As ComboBox)
    'Dim rs As Recordset
    Me.ejecutar "select * from materialesAlmacenes"    '-1 es porq el datalist usa como valor bool -1 en true
    While Not rs.EOF
        muestra = Format(rs!id, "000") & " - " & rs!detalle
        cbo.AddItem muestra
        cbo.ItemData(cbo.NewIndex) = rs!id
        rs.MoveNext
    Wend
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Function

