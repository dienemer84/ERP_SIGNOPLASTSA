VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function ImprimirLista(titulo, Lst As istView, cd As CommonDialog, Optional linea2 = Empty, Optional fondo1 = Empty, Optional fondo2 = Empty) As Boolean
On Error GoTo err91
cd.ShowPrinter

AnchoCol = 0

  For I = 1 To Lst.ColumnHeaders.count
     AnchoCol = AnchoCol + Lst.ColumnHeaders(I).Width
  Next
  Espacio = 0

Printer.Font.Size = 12
Printer.Font.Bold = True
'Printer.Line  (8800, 1400)-(10100, 1400)
Printer.Print UCase(titulo)
If linea2 <> Empty Then
    Printer.Print UCase(linea2)
End If

Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
  With Lst
  
  'Acá se imprimen los encabezados del ListView
  For I = 1 To .ColumnHeaders.count
      Espacio = Espacio + CInt(.ColumnHeaders(I).Width * Printer.ScaleWidth / AnchoCol)
      If Lst.ColumnHeaders(I).Width > 1 Then
        'Printer.Print i
        Printer.Print Lst.ColumnHeaders(I).Text;
      End If
      Printer.CurrentX = Espacio
  Next
Printer.Font.Bold = False
  Printer.Print
    'Imprime una línea
  Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
  Printer.Print
  
  'Este bucle recorre los items y subitems del ListView  y los imprime

  For I = 1 To .ListItems.count
       Espacio = 0
       
      Set litem = .ListItems(I)
      Printer.Print litem.Text;
       'Recorremos las columnas
       For X = 1 To .ColumnHeaders.count - 1
             Espacio = Espacio + CInt(.ColumnHeaders(X).Width * Printer.ScaleWidth / AnchoCol)
             Printer.CurrentX = Espacio
             If Lst.ColumnHeaders(X + 1).Width > 1 Then
           
                Printer.Print litem.SubItems(X);
 
            
            
            End If
       Next
       
       'Otro espacio en blanco
       Printer.Print
       
    
  Next
  
  End With
  
  Printer.Print
  Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
 
    
Printer.Print
If fondo1 <> Empty Then
Printer.Print fondo1
End If

If fondo2 <> Empty Then
Printer.Print fondo2
End If
  
Printer.Print "Fecha emision: " & Format(Now, "dd-mm-yy")

Printer.EndDoc
ImprimirLista = True

Exit Function
err91:
ImprimirLista = False
If tra Then CN.RollbackTrans

End Function
