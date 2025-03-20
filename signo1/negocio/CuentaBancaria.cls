VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuentaBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Id As Long
Public Banco As Banco
Public numero As String
Public TipoCuenta As TipoCuentaBancaria
Public moneda As clsMoneda
Public CBU As String
Public BancoNombre As String


Public Property Get DescripcionCBUFormateada() As String
    Dim col As New Collection

    If IsSomething(Me.Banco) Then
        If LenB(Me.Banco.nombre) > 0 Then
            col.Add Me.Banco.nombre
        End If
    End If

    If LenB(Me.numero) > 0 Then
        col.Add "Nº " & Me.numero
    End If

    If TipoCuenta = TipoCuentaBancaria.CajaAhorro Then
        col.Add "Caja de ahorro"
    ElseIf TipoCuenta = TipoCuentaBancaria.CuentaCorriente Then
        col.Add "Cuenta Corriente"
    End If

    If IsSomething(Me.moneda) Then
        col.Add Me.moneda.NombreCorto
    End If


    col.Add "CBU: " & Me.CBU

    Dim i As Long
    Dim arr() As String
    If col.count > 0 Then ReDim arr(col.count - 1) As String
    For i = 1 To col.count
        arr(i - 1) = col.item(i)
    Next

    DescripcionCBUFormateada = Join(arr, " | ")
End Property


Public Property Get DescripcionFormateada() As String
    Dim col As New Collection

    If IsSomething(Me.Banco) Then
        If LenB(Me.Banco.nombre) > 0 Then
            col.Add Me.Banco.nombre
        End If
    End If

    If LenB(Me.numero) > 0 Then
        col.Add "Nº " & Me.numero
    End If

    If TipoCuenta = TipoCuentaBancaria.CajaAhorro Then
        col.Add "Caja de ahorro"
    ElseIf TipoCuenta = TipoCuentaBancaria.CuentaCorriente Then
        col.Add "Cuenta Corriente"
    End If

    If IsSomething(Me.moneda) Then
        col.Add Me.moneda.NombreCorto
    End If

    Dim i As Long
    Dim arr() As String
    If col.count > 0 Then ReDim arr(col.count - 1) As String
    For i = 1 To col.count
        arr(i - 1) = col.item(i)
    Next

    DescripcionFormateada = Join(arr, " | ")
End Property


Public Property Get DescripcionFormateadaCompleta() As String
    Dim col As New Collection

    If IsSomething(Me.Banco) Then
        If LenB(Me.Banco.nombre) > 0 Then
            col.Add Me.Banco.nombre
        End If
    End If

    If LenB(Me.numero) > 0 Then
        col.Add "Nº " & Me.numero
    End If

    If TipoCuenta = TipoCuentaBancaria.CajaAhorro Then
        col.Add "Caja de ahorro"
    ElseIf TipoCuenta = TipoCuentaBancaria.CuentaCorriente Then
        col.Add "Cuenta Corriente"
    End If

    If IsSomething(Me.moneda) Then
        col.Add Me.moneda.NombreCorto
    End If

    Dim i As Long
    Dim arr() As String
    If col.count > 0 Then ReDim arr(col.count - 1) As String
    For i = 1 To col.count
        arr(i - 1) = col.item(i)
    Next

    DescripcionFormateadaCompleta = Join(arr, " | ")
End Property
