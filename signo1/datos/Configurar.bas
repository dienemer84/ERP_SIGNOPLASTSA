Attribute VB_Name = "Configurar"
Option Explicit
Public IdCtaPercepcionesIVA As Long
Public IdCtaRedondeo As Long
Public IdCtaPercepcionesIIBB As Long
Public IdCtaCombustible As Long
Public IdCtaIVACredito As Long
Private rs_iibb As Recordset
Private rs As Recordset
Private iPorcMO As Double
Private iPorMAMenos10 As Double
Private iPorMAMenos15 As Double
Private iPorMaMas15 As Double
Private iPintM2 As Double
Private iDolar As Double
Private iManteOferta As Long
Private iSueldo As Double
Private iMano_obra_muerta As Double
Private iidPercepcionIva As Long
Private percepcionesIIBB As Collection
Public IdCtaPercepcionesIIBBResto As Long



Public Property Get idPercepcionesIIBB() As Collection
    Set idPercepcionesIIBB = percepcionesIIBB
End Property
Public Property Get Mano_obra_muerta() As Double
    Mano_obra_muerta = iMano_obra_muerta
End Property
Public Property Get PorMAMenos10() As Double
    PorMAMenos10 = iPorMAMenos10
End Property
Public Property Get PorMAMenos15() As Double
    PorMAMenos15 = iPorMAMenos15
End Property
Public Property Get PorMaMas15() As Double
    PorMaMas15 = iPorMaMas15
End Property
Public Property Get PintM2() As Double
    PintM2 = iPintM2
End Property
Public Property Get Dolar() As Double
    Dolar = iDolar
End Property
Public Property Get manteOferta() As Long
    manteOferta = iManteOferta
End Property
Public Property Get Sueldo() As Double
    Sueldo = iSueldo
End Property
Public Property Get PorcMO() As Double
    PorcMO = iPorcMO
End Property
Public Property Get IdPercepcionIVA() As Long
    IdPercepcionIVA = iidPercepcionIva
End Property



Public Function EstaActualizando(cs As String) As Boolean

  Dim cn As ADODB.Connection

  

    Set cn = New ADODB.Connection

    cn.ConnectionString = cs
    cn.Open


Dim rstmp As New ADODB.Recordset

Dim rs As Recordset
    If rstmp.State = 1 Then rstmp.Close
    rstmp.Open "select actualizando from configuracion", cn, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockOptimistic, adCmdText
    Set rs = rstmp
 
    EstaActualizando = rs!Actualizando
    rs.Close
    cn.Close
    


End Function




Public Function LoadConfiguration() As Boolean
    Set rs = conectar.RSFactory("select * from configuracion")
    LoadConfiguration = False
    If Not rs.BOF And Not rs.EOF Then

        IdCtaPercepcionesIVA = rs!IdCtaPercepcionesIVA
        IdCtaPercepcionesIIBB = rs!IdCtaPercepcionesIIBB
        IdCtaIVACredito = rs!IdCtaIVACredito
        IdCtaRedondeo = rs!IdCtaRedondeo
        IdCtaCombustible = rs!IdCtaCombustible

        IdCtaPercepcionesIIBBResto = rs!IdCtaPercepcionesIIBBResto


        iPorcMO = rs!PorcMO
        iPorMAMenos10 = rs!PorMAMenos10
        iPorMAMenos15 = rs!PorMAMenos15
        iPorMaMas15 = rs!PorMaMas15
        iPintM2 = rs!PintM2
        iDolar = rs!Dolar
        iManteOferta = rs!manteOferta
        iSueldo = rs!Sueldo
        iMano_obra_muerta = rs!Mano_obra_muerta
        iidPercepcionIva = rs!IdPercepcionIVA
        Dim P As clsPercepciones

        Set rs_iibb = conectar.RSFactory("select * from configuracion_percepcionesIIBB")
        Set percepcionesIIBB = New Collection
        While Not rs_iibb.EOF
            Set P = New clsPercepciones
            P.id = rs_iibb!idPercepcion
            percepcionesIIBB.Add P, CStr(P.id)
            rs_iibb.MoveNext
        Wend

        LoadConfiguration = True
    End If
End Function
