VERSION 5.00
Begin VB.Form frmFacturaElectronica 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmFacturaElectronica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function factura_electronica()

    Me.WSAFIPFEx1.iniciar SCModoFiscal.Test, "23172338909", "c:\certificados\cert.pfx", ""

    Dim F As Boolean

    F = Me.WSAFIPFEx1.f1ObtenerTicketAcceso
    Debug.Print F

End Function
