VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoriaSueldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Moneda As clsMoneda
Public id As Long
Public nombre As String
Public Valor As Double
Public PorcentajeEspecializacion As Double

Public Sub EspecificarSueldo()
    On Error GoTo E

    'ver de hacer bien esto
    Dim q As String: q = "SELECT sueldo FROM configuracion LIMIT 1"
    Dim r As Recordset
    Set r = RSFactory(q)
    Dim valorSueldo As Double
    While Not r.EOF
        valorSueldo = IIf(IsNull(r!Sueldo), 0, r!Sueldo)
        r.MoveNext
    Wend

    Dim ret As String
    Dim nuevoValor As Double
    ret = InputBox("Especifique el sueldo de los empleados", "Sueldo", valorSueldo)
    If LenB(ret) > 0 Then
        If IsNumeric(ret) Then
            nuevoValor = Val(ret)
            If MsgBox("¿Desea actualizar el valor del sueldo con el nuevo valor [" & nuevoValor & "]?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                q = "UPDATE configuracion SET sueldo = " & conectar.Escape(nuevoValor)
                If conectar.execute(q) Then
                    MsgBox "El valor fue actualizado", vbInformation
                Else
                    MsgBox "No se pudo actualizar el valor", vbCritical
                End If
            End If
        Else
            MsgBox "El valor no es numerico.", vbCritical
        End If
    End If
    Exit Sub
E:
    MsgBox Err.Description, vbCritical
End Sub
