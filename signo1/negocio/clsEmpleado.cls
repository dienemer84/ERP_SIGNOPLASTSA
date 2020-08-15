VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public legajo As Long
Public documento As Long
Public Apellido As String
Public Nombre As String
Public Nombres As String
Public direccion As String
Public localidad As String
Public Telefono1 As String
Public Telefono2 As String
Public estado As EstadoUsuario
Public email As String
Public sectores As Collection
Public id As Long
Public FechaIngreso As Date
Public FechaNacimiento As Date
Public GrupoSanguineo As String
Public Foto As Variant
Public idFoto As Long

Public Cuil As String
Public ObraSocial As ObraSocial
Public UltimaActualizacion As Date


Public Property Get NombreAbreviado() As String
    Dim nom As String
    If LenB(Me.Apellido) > 0 Then nom = nom & Me.Apellido & " "
    If LenB(Me.Nombre) > 0 Then nom = nom & Left(Me.Nombre, 1) & "."

    NombreAbreviado = nom
End Property

Public Property Get NombreCompleto() As String
    Dim nom As String
    If LenB(Me.Apellido) > 0 Then nom = nom & Me.Apellido & " "
    If LenB(Me.Nombre) > 0 Then nom = nom & Me.Nombre

    NombreCompleto = nom
End Property


Public Property Get LegajoAndNombreCompleto()
    Dim nc As String
    nc = Me.legajo & " - " & NombreCompleto
    LegajoAndNombreCompleto = nc
End Property


Public Property Get DireccionCompleta()
    Dim ret As String
    ret = Me.direccion
    If LenB(Me.localidad) > 0 Then ret = ret & " (" & Me.localidad & ")"
    If LenB(Me.Telefono1) > 0 Or LenB(Me.Telefono2) > 0 Then
        ret = ret & " [Tel: "
        If LenB(Me.Telefono1) > 0 Then
            ret = ret & Me.Telefono1
            If LenB(Me.Telefono2) > 0 Then
                ret = ret & ", " & Me.Telefono2
            End If
        Else
            ret = ret & Me.Telefono2
        End If
        ret = ret & "]"
    End If
    DireccionCompleta = ret
End Property





