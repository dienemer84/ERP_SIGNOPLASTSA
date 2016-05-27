Attribute VB_Name = "DAOdtoCuentasTerminacion"
Option Explicit
Dim dto As New DTOCuentasTerminacion


Public Function GetConfigTerminacion() As DTOCuentasTerminacion

    Dim q As String
    q = "SELECT  id, CantPint, Fosfatos, superficie,  aplicacion, horneado, sector,  rubro  FROM  sp.terminacion_cuentas "
    Dim rs As Recordset
    Set rs = conectar.RSFactory(q)
    If Not rs.EOF And Not rs.BOF Then
        Set dto.Sector = DAOSectores.GetById(rs!Sector)
        Set dto.Rubro = DAORubros.FindById(rs!Rubro)
        Set dto.CantidadFosfatos = DAOMateriales.FindById(rs!fosfatos)
        Set dto.CantidadPintura = DAOMateriales.FindById(rs!cantPint)
        Set dto.Aplicacion = DAOTareas.FindById(rs!Aplicacion)
        Set dto.Horneado = DAOTareas.FindById(rs!Horneado)
        Set dto.Limpieza = DAOTareas.FindById(rs!superficie)
    End If

    Set GetConfigTerminacion = dto

End Function




Public Function SaveConfigTerminacion(dto As DTOCuentasTerminacion) As Boolean
    Dim strsql As String
    strsql = "update terminacion_cuentas set CantPint=" & dto.CantidadPintura.id & ", Fosfatos=" & dto.CantidadFosfatos.id & "," _
             & " superficie=" & dto.Limpieza.id & ", aplicacion=" & dto.Aplicacion.id & ", horneado=" & dto.Horneado.id & ", sector=" _
             & dto.Sector.id & ", rubro= " & dto.Rubro.id

    SaveConfigTerminacion = conectar.execute(strsql)

End Function
