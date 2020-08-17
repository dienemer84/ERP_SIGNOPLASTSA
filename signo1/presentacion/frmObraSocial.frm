VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmObraSocial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Obras Sociales"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   Begin GridEX20.GridEX grid 
      Height          =   3915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   6906
      Version         =   "2.0"
      PreviewRowIndent=   500
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewRowLines =   2
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowColumnDrag =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      RowHeaders      =   -1  'True
      ItemCount       =   1
      DataMode        =   99
      AllowAddNew     =   -1  'True
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   1
      Column(1)       =   "frmObraSocial.frx":0000
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmObraSocial.frx":0118
      FormatStyle(2)  =   "frmObraSocial.frx":0250
      FormatStyle(3)  =   "frmObraSocial.frx":0300
      FormatStyle(4)  =   "frmObraSocial.frx":03B4
      FormatStyle(5)  =   "frmObraSocial.frx":048C
      FormatStyle(6)  =   "frmObraSocial.frx":0544
      ImageCount      =   0
      PrinterProperties=   "frmObraSocial.frx":0624
   End
End
Attribute VB_Name = "frmObraSocial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim os As Collection
Private Sub Form_Load()
    FormHelper.Customize Me

  

    CargarOS
End Sub

Private Sub CargarOS()
    Set os = DAOObraSocial.GetAll
    Me.grid.ItemCount = os.count
End Sub

Private Sub grid_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
 Dim o As ObraSocial
 Set o = New ObraSocial
    With o
   
        .nombre = Values(1)
        
    End With
    If DAOObraSocial.Save(o) Then
        os.Add o, CStr(o.id)
    Else
        MsgBox "Hubo un error al guardar los valores"
    End If

End Sub

Private Sub grid_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
   If RowIndex > os.count Then Exit Sub
    
  Dim o As ObraSocial
    Set o = os.item(RowIndex)
    With o
        Values(1) = .nombre
        
    End With
End Sub

Private Sub grid_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
If RowIndex > os.count Then Exit Sub
    
    
  Dim o As ObraSocial
    Set o = os.item(RowIndex)
    With o
    
        .nombre = Values(1)
        
    End With
    If Not DAOObraSocial.Save(o) Then MsgBox "Hubo un error al guardar los valores"
End Sub
