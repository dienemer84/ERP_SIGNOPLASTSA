Attribute VB_Name = "GridEXHelper"
Option Explicit
Public Const GRID_FORMATSTYLE_ANULADA As String = "ANULADA__"    '__ porque ya habia estilos con esos nombres
Public Const GRID_FORMATSTYLE_VERDE As String = "VERDE__"
Public Const GRID_FORMATSTYLE_ROJO As String = "ROJO__"

Public Sub AutoSizeColumns(ByRef grid As GridEX20.GridEX, Optional autoSize As Boolean = True)
    grid.ColumnAutoResize = autoSize
    grid.ContinuousScroll = True
    Dim jscol As GridEX20.JSColumn
    For Each jscol In grid.Columns
        jscol.autoSize
    Next jscol
End Sub

Public Sub CustomizeGrid(grid As GridEX20.GridEX, Optional GroupByVisible As Boolean = False, Optional Editable As Boolean = False)
    grid.GroupByBoxVisible = GroupByVisible
    grid.ContinuousScroll = True
    CustomizeGridColors grid
    grid.AllowEdit = Editable
    grid.ColumnAutoResize = True
    grid.Font = "Tahoma"
    grid.ColumnHeaderFont = "Tahoma"
    BuildStyles grid
End Sub
Public Sub AddColumnReportControl(ReportControl As ReportControl, ByVal index As Long, ByVal caption As String, Optional align As XTPColumnAlignment = xtpAlignmentLeft, Optional ByVal tree As Boolean = False, Optional ByVal Width As Double = 25)
    Dim col As ReportColumn
    Set col = ReportControl.Columns.Add(index, caption, Width, True)
    col.Icon = 0
    col.Sortable = True
    col.AllowDrag = False
    col.AllowRemove = False
    col.Alignment = align
    col.TreeColumn = tree
    If Width <> 25 Then
        col.autoSize = True
        col.BestFitMode = XTPColumnBestFitMode.xtpBestFitModeAllData
    End If
End Sub
Private Sub CustomizeGridColors(grid As GridEX20.GridEX)
    grid.FormatStyles(5).BackColor = FormHelper.FondoCeleste
    grid.FormatStyles(5).ForeColor = FormHelper.LetraAzul     'vbBlack  'vbWhite 'FormHelper.FondoCeleste
    grid.ContinuousScroll = True
    grid.FormatStyles(5).FontBold = True
    grid.ColumnHeaderFont = "Tahoma"
    grid.ColumnHeaderFont.Bold = False
    grid.BackColor = vbWhite    ' FormHelper.FondoCeleste
    grid.BackColorGBBox = FormHelper.FondoAzul    'GRILLA_BACKCOLOR_GBBOX_INFOTEXT
    grid.BackColorInfoText = FormHelper.FondoAzul    'GRILLA_BACKCOLOR_GBBOX_INFOTEXT
    grid.ForeColorHeader = FormHelper.LetraAzul
    grid.ForeColorRowGroup = vbBlack    'GRILLA_BACKCOLOR_GBBOX_INFOTEXT
    grid.BackColorRowGroup = &HF5E6DC
    grid.BackColorHeader = FormHelper.GridBGHeader  'BACK_COLOR_HEADER
    grid.ForeColorInfoText = FormHelper.FondoCeleste    ' &HFFFFFFF 'FORE_COLOR_INFO_TEXT
    grid.GroupByBoxInfoText = "Arrastre una columna aquí para ordenar por dicha columna."
End Sub


Public Sub ColumnHeaderClick(ByRef grid As GridEX20.GridEX, ByRef Column As GridEX20.JSColumn)
    Dim grTemp As JSGroup
    Dim SortOrder As jgexSortOrderConstants

    If Column.IsGrouped Then

        For Each grTemp In grid.Groups
            If grTemp.ColIndex = Column.index Then
                GroupByBoxHeaderClick grTemp
                Exit For
            End If
        Next
    Else
        SortOrder = Column.SortOrder
        grid.SortKeys.Clear
        If SortOrder = jgexSortAscending Then
            grid.SortKeys.Add Column.index, jgexSortDescending
        Else
            grid.SortKeys.Add Column.index, jgexSortAscending
        End If
    End If
    grid.row = 0
End Sub

Public Sub GroupByBoxHeaderClick(ByRef Group As GridEX20.JSGroup)
    'When clicking in a group by box header we change SortOrder for that group
    Group.SortOrder = -Group.SortOrder
End Sub
Public Sub Grid2Clipboard(grid As GridEX)
    Dim texto As String
    'On Error Resume Next

    Dim i As Long
    Dim j As Long

    Dim oldSelected As Long: oldSelected = grid.row

    Dim tmpString As String
    Dim cabeceras As String

    For i = 1 To grid.rowcount
        grid.row = i
        For j = 1 To grid.Columns.count
            If grid.Columns(j).Visible = True Then
                If i = 1 Then cabeceras = cabeceras & grid.Columns(j).caption & vbTab


                If IsNull(grid.value(j)) Then
                    texto = texto & vbTab
                Else
                    If IsNumeric(grid.value(j)) Then
                        texto = texto & Replace$(funciones.RedondearDecimales(Val(grid.value(j))), vbNewLine, vbNullString) & vbTab

                    Else
                        texto = texto & Replace$(grid.value(j), vbNewLine, vbNullString) & vbTab
                    End If
                End If

            End If



        Next j
        texto = texto & vbNewLine
    Next i

    grid.row = oldSelected

    Clipboard.SetText vbNullString
    Clipboard.Clear
    Clipboard.SetText cabeceras & vbNewLine & texto
End Sub


Private Sub BuildStyles(grid As GridEX)
    Dim Style As JSFormatStyle

    '''''''''''''''''''''''''''anulada
    Set Style = grid.FormatStyles.Add(GRID_FORMATSTYLE_ANULADA)
    Style.ForeColor = vbRed
    Style.FontName = "Tahoma"
    Style.FontStrikethru = True
    Style.FontItalic = True

    '''''''''''''''''''''''''''verde
    Set Style = grid.FormatStyles.Add(GRID_FORMATSTYLE_VERDE)
    Style.ForeColor = &H3B9905
    Style.FontName = "Tahoma"

    '''''''''''''''''''''''''''rojo
    Set Style = grid.FormatStyles.Add(GRID_FORMATSTYLE_ROJO)
    Style.ForeColor = &HB29FF
    Style.FontName = "Tahoma"

End Sub





Public Function ReportControlAddColumn(ReportControl As ReportControl, ByVal index As Long, ByVal caption As String, Optional align As XTPColumnAlignment = xtpAlignmentLeft, Optional ByVal tree As Boolean = False, Optional ByVal Width As Double = 25) As ReportColumn
    Dim col As ReportColumn
    Set col = ReportControl.Columns.Add(index, caption, Width, True)
    col.Icon = 0
    col.Sortable = True
    col.AllowDrag = False
    col.AllowRemove = False
    col.Alignment = align
    col.TreeColumn = tree
    If Width <> 25 Then
        col.autoSize = True
        col.BestFitMode = XTPColumnBestFitMode.xtpBestFitModeAllData
    End If
    Set ReportControlAddColumn = col
End Function
