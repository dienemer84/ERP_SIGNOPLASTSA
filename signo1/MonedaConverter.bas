Attribute VB_Name = "MonedaConverter"
Option Explicit
Private Monedas As Collection
Private monedaPatron As clsMoneda
Private lastUpdate As Date
Public MonedaDN As clsMoneda


Public Property Get Patron() As clsMoneda
    If Not IsSomething(monedaPatron) Then ActualizarMonedas
    Set Patron = monedaPatron
End Property





Public Function Convertir(Valor As Double, monedaOrigenId As Long, monedaDestinoId As Long) As Double
    If monedaOrigenId = monedaDestinoId Then
        Convertir = Valor
        Exit Function
    End If


    'si no se inicializaron las monedas o se actualizaron hace mas de 15 minutos, actualizo las monedas
    'If (monedas Is Nothing Or monedaPatron Is Nothing) Or (DateDiff("n", lastUpdate, Now) > 15) Then
    If DateDiff("n", lastUpdate, Now) > 15 Then ActualizarMonedas

    Dim monOrigen As clsMoneda
    Dim monDestino As clsMoneda

    Set monOrigen = Monedas.item(CStr(monedaOrigenId))
    Set monDestino = Monedas.item(CStr(monedaDestinoId))

    Dim Cambio As Double
    Cambio = monDestino.MonedaCambio.Cambio
    If monDestino.MonedaCambio.Id = monDestino.Id Then Cambio = 1
    If monedaPatron.Id = monDestino.Id Then Cambio = 1
    If monOrigen.Id = monDestino.Id Then Cambio = 1

    '    If monOrigen.Id <> monDestino.Id Then
    '    MsgBox ("Tenemos monedas distintas")
    '        MsgBox ("Moneda de la OT del Remito asociado: " & monOrigen.NombreCorto & " | Moneda del Comprobante: " & monDestino.NombreCorto)
    '        MsgBox ("Moneda de la OT del Remito asociado: " & monOrigen.NombreCorto & vbCrLf & "" _
             '        & "La Moneda del Comprobante que se está cargando es: " & monDestino.NombreCorto & vbCrLf & "" _
             '       & "Se procede a realizar la conversión correspondiente." & vbCrLf & "" _
             '        & "Cálculo:" & vbCrLf & "" _
             '       & "Importe del item: " & monOrigen.NombreCorto & " " & Valor & vbCrLf & "" _
             '       & " * Valor de Moneda de OT: " & monDestino.NombreCorto & " " & monOrigen.Cambio)

    Convertir = ((Valor * monOrigen.Cambio) / monDestino.Cambio * Cambio) / monedaPatron.Cambio
    '    Else
    '    MsgBox ("Tenemos las mismas monedas")
    '
    '    End If




End Function




Public Function ConvertirForzado2(Valor As Double, monedaOrigenId As Long, monedaDestinoId As Long, cambioforzado As Double) As Double
'si no se inicializaron las monedas o se actualizaron hace mas de 15 minutos, actualizo las monedas
'If (monedas Is Nothing Or monedaPatron Is Nothing) Or (DateDiff("n", lastUpdate, Now) > 15) Then
    If DateDiff("n", lastUpdate, Now) > 15 Then ActualizarMonedas

    Dim monOrigen As clsMoneda
    Dim monDestino As clsMoneda

    Set monOrigen = Monedas.item(CStr(monedaOrigenId))
    Set monDestino = Monedas.item(CStr(monedaDestinoId))


    Dim Cambio As Double
    Cambio = cambioforzado

    If (monedaOrigenId = monedaDestinoId) Then

        ConvertirForzado2 = Valor

    Else
        If (monDestino.Id = monedaPatron.Id) Then
            ConvertirForzado2 = Valor / cambioforzado

        Else
            ConvertirForzado2 = Valor * cambioforzado
        End If

    End If

End Function

Public Function ConvertirForzado(Valor As Double, monedaOrigenId As Long, cambioforzado As Double) As Double
'si no se inicializaron las monedas o se actualizaron hace mas de 15 minutos, actualizo las monedas
'If (monedas Is Nothing Or monedaPatron Is Nothing) Or (DateDiff("n", lastUpdate, Now) > 15) Then
    If DateDiff("n", lastUpdate, Now) > 15 Then ActualizarMonedas

    Dim monOrigen As clsMoneda
    Dim monDestino As clsMoneda

    Set monOrigen = Monedas.item(CStr(monedaOrigenId))


    Dim Cambio As Double
    Cambio = cambioforzado  'monDestino.MonedaCambio.Cambio
    'If monDestino.MonedaCambio.Id = monDestino.Id Then Cambio = 1
    ' If monedaPatron.Id = monDestino.Id Then Cambio = 1
    '  If monOrigen.Id = monDestino.Id Then Cambio = 1

    ConvertirForzado = ((Valor * monOrigen.Cambio) / Cambio) / monedaPatron.Cambio

End Function


Public Sub ActualizarMonedas()
    Dim mon As clsMoneda
    Set Monedas = DAOMoneda.GetAll

    Set monedaPatron = Nothing
    For Each mon In Monedas
        If mon.Patron Then
            Set monedaPatron = mon
            Exit For
        End If
    Next mon

    lastUpdate = Now
End Sub


