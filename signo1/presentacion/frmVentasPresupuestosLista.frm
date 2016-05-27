VERSION 5.00
Begin VB.Form frmVentasPresupuestosLista 
   Caption         =   "Lista de Presupuestos"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15300
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   15300
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   3240
      TabIndex        =   0
      Top             =   2520
      Width           =   2055
   End
End
Attribute VB_Name = "frmVentasPresupuestosLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim presupuestos As Collection


Private Sub Command1_Click()
Set presupuestos = DAOPresupuestos.GetAll
End Sub

