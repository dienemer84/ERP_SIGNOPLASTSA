Attribute VB_Name = "LabelHelper"
Option Explicit





Public Function PrintEtiquetaLegajos() As Boolean


    Dim c As New Collection
    Dim E As clsEmpleado
    Set c = DAOEmpleados.GetAll(" and legajo>1 and legajo<200 order by Legajo ASC")

    For Each E In c
        Printer.CurrentX = CentrarImpresion("SIGNO:PLAST")
        Printer.Print "SIGNO:PLAST"
        Printer.Font.Size = 2
        Printer.Print " "
        Printer.Font.Size = 22
        Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
        Printer.Font.Size = 22
        Printer.Font = "ARIAL BLACK"
        Printer.FontBold = True
        Printer.CurrentX = 50


        Printer.Print "Legajo: " & E.legajo
        Printer.CurrentX = 50
        Printer.Font.Size = 12
        Printer.Print E.NombreCompleto
        Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)

        Printer.EndDoc
    Next

    Exit Function


End Function




Public Function PrintEtiquetaPedido(Pedido As OrdenTrabajo) As Boolean
    On Error GoTo err1
    PrintEtiquetaPedido = True
    Printer.Font.Size = 14
    Printer.Font = "ARIAL BLACK"
    Printer.FontBold = True
    Printer.CurrentX = CentrarImpresion(Pedido.Cliente.razon)
    Printer.Print Pedido.Cliente.razon
    Printer.CurrentX = 50
    Printer.Print "OT:  " & Pedido.id


    Printer.CurrentX = 50
    Printer.Print "OC :" & Pedido.descripcion
    Printer.Font.Size = 8
    Printer.CurrentX = 50

    Dim c As Double
    Dim d As DetalleOrdenTrabajo

    c = 0
    For Each d In Pedido.Detalles

        c = c + d.CantidadPedida
    Next d

    Printer.Print "TOTAL: " & Pedido.Detalles.count & " items, " & funciones.FormatearDecimales(c) & " elementos"
    Printer.EndDoc
    Exit Function
err1:
    PrintEtiquetaPedido = False
End Function


Public Function PrintEtiquetaDetallePedido(Pedido As OrdenTrabajo, deta As DetalleOrdenTrabajo, CantidadPedida As Double, Optional detaConj As DetalleOTConjuntoDTO = Nothing) As Boolean
    On Error GoTo err1
    PrintEtiquetaDetallePedido = True
    Dim twTamano As Long
    Dim header As String
    Dim x As Long
    'para una etiuqeutas de 1 banda!
    Dim Total As Long

    If IsSomething(detaConj) Then
        Total = detaConj.Cantidad * CantidadPedida
    Else
        Total = CantidadPedida
    End If

    If Total > 5 Then

        Printer.Font.Size = 15
        Printer.Font = "ARIAL BLACK"
        Printer.FontBold = True
        Printer.CurrentX = CentrarImpresion(Trim(Pedido.Cliente.razon))
        Printer.Print Pedido.ClienteFacturar.razon
        Printer.CurrentX = 50
        Printer.Print "OT| ITEM:  " & Pedido.id & " | " & deta.item
        Printer.CurrentX = 50
        Printer.Print "Cantidad :" & Total
        Printer.Font.Size = 10
        Printer.CurrentX = 50
        Printer.Print "Pieza: " & deta.Pieza.nombre
        Printer.EndDoc
    End If

    For x = 1 To Total
        twTamano = FormHelper.ConvertCmToTwip(6)
        Printer.Font = "ARIAL BLACK"
        header = Pedido.ClienteFacturar.razon
        Printer.CurrentY = 50
        Printer.CurrentX = 50
        Printer.Font.Bold = True
        Printer.Font.Size = 11
        Printer.Print header
        Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
        Printer.Line (0, Printer.CurrentY + 10)-(Printer.ScaleWidth, Printer.CurrentY + 10)


        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 20
        Printer.Font = "Arial"
        Printer.Font.Size = "7"
        Printer.CurrentX = 50


        Printer.Print "O.C: " & Trim(Replace(Pedido.descripcion, vbNewLine, "", , , vbTextCompare))
        Printer.CurrentX = 50
        Printer.Font.Size = "8"
        Printer.CurrentX = 50

        'If (deta.FechaEntrega >= Now) Then Printer.Print "Entrega:  " & deta.FechaEntrega;


        If IsSomething(detaConj) Then

            Printer.Print vbTab & "UM: Subconjunto"

        Else

            If deta.Pieza.EsConjunto Then
                Printer.Print vbTab & "UM: Conjunto"
            Else
                Printer.Print vbTab & "UM: Unidad"
            End If

        End If




        Printer.CurrentX = 50
        If IsSomething(detaConj) Then
            Printer.Print deta.item & " | " & detaConj.Pieza.nombre
        Else
            Printer.Print deta.item & " |  " & deta.Pieza.nombre
        End If
        Printer.CurrentX = 0
        Printer.Font.Size = "7"

        Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
        Printer.Line (0, Printer.CurrentY + 10)-(Printer.ScaleWidth, Printer.CurrentY + 10)

        Dim barra As String
        barra = "*" & deta.id & "*"


        Printer.CurrentX = funciones.CentrarImpresion(barra) + 130
        Printer.CurrentY = Printer.CurrentY + 20
        Printer.FontSize = 11
        Printer.Font = "IDAUTOMATIONHC39M"
        Printer.Print barra
        Printer.Font = "arial black"
        Printer.FontBold = True
        Printer.FontSize = 12



        Printer.CurrentX = 100
        Printer.CurrentY = 900
        Printer.Print "SIGNO"
        Printer.CurrentY = 1150
        Printer.CurrentX = 100
        Printer.Print "PLAST"
        Printer.CurrentY = 1400
        Printer.CurrentX = 100
        Printer.FontSize = 9
        Printer.Print deta.OrdenTrabajo.id & "|" & deta.item
        Printer.Line (1200, 950)-(1200, 1700)
        Printer.Line (1205, 950)-(1205, 1700)

        'incremento la cant de etiquetas del item?
        DAODetalleOrdenTrabajo.CountPrintedLabels deta.id

        Printer.EndDoc

    Next x
    Exit Function
err1:
    PrintEtiquetaDetallePedido = False
End Function




