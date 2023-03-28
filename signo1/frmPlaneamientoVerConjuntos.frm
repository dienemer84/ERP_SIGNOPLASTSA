VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDesarrolloVerConjuntos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ver conjunto"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   11880
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   5760
      Width           =   975
   End
End
Attribute VB_Name = "frmDesarrolloVerConjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsS As New classStock
Dim vIdConjunto As Long
Public Property Let idConjunto(nIdConjunto As Long)
    vIdConjunto = nIdConjunto
End Property
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    cargarConjunto
End Sub
Private Sub cargarConjunto()
    Dim rs3 As Recordset  'nivel 2
    Dim rs2 As Recordset  'nivel 1
    Dim rs As Recordset   'raiz
    Dim clsP As New classPlaneamiento
    Me.TreeView1.Nodes.Clear
    With Me.TreeView1.Nodes
        xy = 0
        xyz = 0
        xyzw = 0
        Set rs = conectar.RSFactory("select detalle,conjunto from stock where id=" & vIdConjunto)
        If rs!conjunto = 0 Then

            'agrego el nombre
            A = rs!detalle
            .Add , tvwroot, "'" & vIdConjunto & "'", A
            'agrego las ramas nivel 1

            Set rs = conectar.RSFactory("select sc.cantidad,sc.idPiezaHija,s.detalle,s.conjunto from stock s inner join stockConjuntos sc on s.id=sc.idPiezaHija where idPiezaPadre=" & vIdConjunto)
            While Not rs.EOF
                xy = xy + 1
                B = rs!idPiezaHija
                .Add "'" & vIdConjunto & "'", tvwChild, "'" & B & xy & "'", rs!Cantidad & " -> " & rs!detalle
                If rs!conjunto > -1 Then
                    Set rs2 = conectar.RSFactory("select sc.cantidad,sc.idPiezaHija,s.detalle,s.conjunto from stock s inner join stockConjuntos sc on s.id=sc.idPiezaHija where idPiezaPadre=" & B)
                    While Not rs2.EOF
                        xyz = xyz + 1
                        c = rs2!idPiezaHija
                        .Add "'" & B & xy & "'", tvwChild, "'" & c & xyz & "'", rs2!Cantidad * rs!Cantidad & " -> " & rs2!detalle
                        If rs2!conjunto > -1 Then
                            Set rs3 = conectar.RSFactory("select sc.cantidad,sc.idPiezaHija,s.detalle,s.conjunto from stock s inner join stockConjuntos sc on s.id=sc.idPiezaHija where idPiezaPadre=" & c)
                            While Not rs3.EOF
                                xyzw = xyzw + 1
                                d = rs3!idPiezaHija
                                .Add "'" & c & xyz & "'", tvwChild, "'" & d & xyzw & "'", rs3!Cantidad * rs2!Cantidad * rs!Cantidad & " -> " & rs3!detalle
                                rs3.MoveNext
                            Wend
                        End If


                        rs2.MoveNext
                    Wend
                End If
                rs.MoveNext
            Wend
        End If
    End With

    'Me.Text1 = clsS.calcularConjunto(vIdConjunto, 1)
End Sub

