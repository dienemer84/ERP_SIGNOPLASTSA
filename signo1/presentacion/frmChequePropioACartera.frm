VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmChequePropioACartera 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheque Propio a Cartera"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChequePropioACartera.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnPasar 
      Height          =   435
      Left            =   945
      TabIndex        =   3
      Top             =   3780
      Width           =   1590
      _Version        =   786432
      _ExtentX        =   2805
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Pasar a Cartera"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtOrigenDestino 
      Height          =   825
      Left            =   1365
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2805
      Width           =   2880
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1365
      TabIndex        =   0
      Top             =   1920
      Width           =   1245
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFechaVenc 
      Height          =   330
      Left            =   1365
      TabIndex        =   1
      Top             =   2340
      Width           =   1245
      _Version        =   786432
      _ExtentX        =   2196
      _ExtentY        =   582
      _StockProps     =   68
      Format          =   1
      CurrentDate     =   40184.6576388889
   End
   Begin XtremeSuiteControls.PushButton btnCerrar 
      Height          =   435
      Left            =   2655
      TabIndex        =   4
      Top             =   3780
      Width           =   1590
      _Version        =   786432
      _ExtentX        =   2805
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Cerrar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFDBBF&
      X1              =   135
      X2              =   4230
      Y1              =   1725
      Y2              =   1725
   End
   Begin VB.Label lblMoneda 
      AutoSize        =   -1  'True
      Caption         =   "Moneda: "
      Height          =   195
      Left            =   645
      TabIndex        =   11
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Observación:"
      Height          =   195
      Left            =   300
      TabIndex        =   10
      Top             =   2775
      Width           =   960
   End
   Begin VB.Label lblBanco 
      AutoSize        =   -1  'True
      Caption         =   "Banco: "
      Height          =   195
      Left            =   765
      TabIndex        =   9
      Top             =   570
      Width           =   540
   End
   Begin VB.Label lblChequera 
      AutoSize        =   -1  'True
      Caption         =   "Chequera: "
      Height          =   195
      Left            =   495
      TabIndex        =   8
      Top             =   945
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Monto:"
      Height          =   195
      Left            =   750
      TabIndex        =   7
      Top             =   1950
      Width           =   510
   End
   Begin VB.Label lblFechaVenc 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Venc:"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   2385
      Width           =   885
   End
   Begin VB.Label lblNumero 
      AutoSize        =   -1  'True
      Caption         =   "Número: "
      Height          =   195
      Left            =   645
      TabIndex        =   5
      Top             =   195
      Width           =   660
   End
End
Attribute VB_Name = "frmChequePropioACartera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cheque As cheque


Private Sub btnCerrar_Click()
    Unload Me
End Sub

Private Sub btnPasar_Click()

    If Val(Me.txtMonto.Text) <= 0 Then
        MsgBox "El monto no es valido.", vbCritical
        Exit Sub
    End If

    If MsgBox("¿Confirma el pasaje del cheque propio a cartera con los sgtes datos?" & vbNewLine & "Monto: " & Me.cheque.moneda.NombreCorto & " " & funciones.FormatearDecimales(Val(Me.txtMonto.Text)) & vbNewLine & "Fecha Vencimiento: " & Me.dtpFechaVenc.value, vbQuestion + vbYesNo, "Pasaje") = vbYes Then
        Me.cheque.FechaVencimiento = Me.dtpFechaVenc.value
        Me.cheque.Monto = Val(Me.txtMonto.Text)
        Me.cheque.observaciones = Me.txtOrigenDestino.Text
        Me.cheque.EnCartera = True

        If DAOCheques.Guardar(Me.cheque) Then
            Unload Me
            Dim ev As New clsEventoObserver
            Set ev.Elemento = Me.cheque
            ev.EVENTO = agregar_
            ev.Tipo = PasajeChequePropioCartera
            Channel.Notificar ev, PasajeChequePropioCartera
        Else
            MsgBox "Hubo un problema a guardar el cheque.", vbCritical
        End If
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me

    Dim ch As chequera
    Set ch = DAOChequeras.GetById(Me.cheque.IdChequera)

    Me.lblNumero.caption = "Número: " & Me.cheque.numero
    Me.lblBanco.caption = "Banco: " & Me.cheque.Banco.nombre
    Me.lblChequera.caption = "Chequera: " & ch.numero & " (" & ch.NumeroDesde & " - " & ch.NumeroHasta & ")"
    Me.lblMoneda.caption = "Moneda: " & cheque.moneda.NombreCorto

    Me.dtpFechaVenc.value = DateAdd("d", 30, CLng(Now))

    Me.btnPasar.Enabled = Not Me.cheque.EnCartera
End Sub
