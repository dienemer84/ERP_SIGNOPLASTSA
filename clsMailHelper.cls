VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMailHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents oMail As clsCDOmail
Attribute oMail.VB_VarHelpID = -1

Public Event Error(descripcion As String, numero As Variant)
Public Event EnvioCompleto()

Public Function EnviarEmail(asunto As String, mensaje As String, destino As String, Optional value As ISuscriber) As Boolean
    On Error GoTo err1
    If Not IsMissing(value) Then
        Dim EVENTO As New clsEventoObserver
        Set EVENTO.Elemento = value


        EVENTO.EVENTO = agregar_

        EVENTO.Tipo = EnvioMail_
    End If


    EnviarEmail = True
    Set oMail = New clsCDOmail
    With oMail
        .servidor = "smtp.gmail.com"
        .puerto = 465
        .UseAuntentificacion = True
        .ssl = True
        .usuario = "signoplast@gmail.com"
        .PassWord = "3l3c720n"
        .asunto = asunto

        .de = "SISTEMA Signo Plast S.A."
        .para = destino
        .mensaje = mensaje
        .Enviar_Backup
    End With
    Set oMail = Nothing
    Exit Function
    RaiseEvent EnvioCompleto
    If Not IsMissing(value) Then
        Channel.Notificar EVENTO, EnvioMail_
    End If

err1:
    RaiseEvent Error(Err.Description, Err.Number)
    Channel.Notificar EVENTO, EnvioMail_
End Function

Public Function EnviarMailEvento(descripcion As String, usuario As String, destino As String, descripcion2 As String, Optional value As ISuscriber) As Boolean
    Dim asunto As String
    Dim mensaje As String

    asunto = "NUEVO EVENTO EN  SIGNO PLAST"
    mensaje = "<html>El usuario " & usuario & " a generado un nuevo evento. " & vbNewLine & "<p><b>" & descripcion & "</b></td>"
    mensaje = mensaje & vbNewLine & "<p>" & descripcion2 & "</p></html>"


    'If IsMissing(value) Then
    'EnviarEmail asunto, mensaje, destino

    ERPHelper.SendMail asunto, mensaje, destino

    'Else
    '  EnviarEmail asunto, mensaje, destino, value
    '  ERPHelper.SendMail asunto, mensaje, destino, value
    'End If



End Function

Public Function isEmail(email As String) As Boolean
    Dim myAt As Integer
    Dim myDot As Integer
    Dim myDotDot As Integer

    isEmail = True
    myAt = InStr(1, email, "@", vbTextCompare)
    myDot = InStr(myAt + 2, email, ".", vbTextCompare)
    myDotDot = InStr(myAt + 2, email, "..", vbTextCompare)
    If myAt = 0 Or myDot = 0 Or Not myDotDot = 0 Or Right(email, 1) = "." Then isEmail = False
End Function

