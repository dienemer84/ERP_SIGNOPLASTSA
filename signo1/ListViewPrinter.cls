VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classListViewPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'+==+==+==+==+==+==+==+==+==+==+==+==+==+
'  ListView Printer Class
'  Author: Opal Raj Ghimire, buna48@hotmail.com
'  You are free to use this class in your projects.
'  Credit me where it deserves
'  A notification is appreciated.
'+==+==+==+==+==+==+==+==+==+==+==+==+==+
'  This class comes with absolutely
'  NO WAR ranty ! Use it at your own risk !!!
'  May posess errors
'+==+==+==+==+==+==+==+==+==+==+==+==+==+

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public HasPicture As Boolean: Public PicHeight As Long: Public PicWidth As Long
Attribute PicHeight.VB_VarUserMemId = 1073938432
Attribute PicWidth.VB_VarUserMemId = 1073938432
Public RowFrom As Long: Public RowTo As Long: Public RowHeight As Long: Public LastRowPrinted As Boolean
Attribute RowFrom.VB_VarUserMemId = 1073938435
Attribute RowTo.VB_VarUserMemId = 1073938435
Attribute RowHeight.VB_VarUserMemId = 1073938435
Attribute LastRowPrinted.VB_VarUserMemId = 1073938435
Public NumOfRowsPerPage As Integer: Dim lngLastRow As Long
Attribute NumOfRowsPerPage.VB_VarUserMemId = 1073938439
Attribute lngLastRow.VB_VarUserMemId = 1073938439
Public DrawHorizontalLines As Boolean: Public DrawVerticalLines As Boolean: Dim bHeadPrinted As Boolean
Attribute DrawHorizontalLines.VB_VarUserMemId = 1073938441
Attribute DrawVerticalLines.VB_VarUserMemId = 1073938441
Attribute bHeadPrinted.VB_VarUserMemId = 1073938441
Public BorderDistance As Long: Public DrawBorder As Boolean
Attribute BorderDistance.VB_VarUserMemId = 1073938444
Attribute DrawBorder.VB_VarUserMemId = 1073938444
Public PosX As Long: Public PosY As Long:
Attribute PosX.VB_VarUserMemId = 1073938446
Attribute PosY.VB_VarUserMemId = 1073938446
Public CurrentX As Long
Attribute CurrentX.VB_VarUserMemId = 1073938448
Public CurrentY As Long
Attribute CurrentY.VB_VarUserMemId = 1073938449
Dim objListView As Object:
Attribute objListView.VB_VarUserMemId = 1073938450
Dim lngWidthHor As Long, lngWidthVer As Long, lngWidthBor As Long
Attribute lngWidthHor.VB_VarUserMemId = 1073938451
Attribute lngWidthVer.VB_VarUserMemId = 1073938451
Attribute lngWidthBor.VB_VarUserMemId = 1073938451
Dim lngColorHor As Long, lngColorVer As Long, lngColorBor As Long
Attribute lngColorHor.VB_VarUserMemId = 1073938454
Attribute lngColorVer.VB_VarUserMemId = 1073938454
Attribute lngColorBor.VB_VarUserMemId = 1073938454

Private Const DT_END_ELLIPSIS = 32768: Const DT_MODIFYSTRING = 65536

Enum LineID
    VerticalLines = 1
    HorizontalLines = 2
    BorderLines = 3
End Enum

Private Sub Class_Initialize()
    PicHeight = 16: PicWidth = 16
    BorderDistance = 5
    DrawHorizontalLines = True: DrawVerticalLines = True: DrawBorder = True
    lngWidthBor = 1: lngWidthHor = 1: lngWidthVer = 1
    lngColorHor = vbBlack: lngColorVer = vbBlack: lngColorBor = vbBlack
    NumOfRowsPerPage = 10
    lngLastRow = 1
    RowHeight = 250

End Sub

'  Prints the Head(Column header of ListView Control)

Public Sub PrintHead(Obj As Object)
    Dim lngKount As Long, rectBox As RECT, strTmp As String, lngBUpScalemode As Long
    Dim intAlign As Integer, intIconNum As Integer, intSpace_Btn_Txt_n_Pic As Integer
    Dim lngPosY As Long, lngPosX As Long, dx As Long, dY As Long
    Dim lngNumOfColumns As Long, lngNumOfSubItems As Long, lngRowHeight As Long
    Dim sngXFac As Single, sngYFac As Single, intPad As Integer

    sngXFac = 1: sngYFac = 1
    intPad = 3
    Obj.ForeColor = objListView.ForeColor
    lngNumOfColumns = objListView.ColumnHeaders.count
    lngNumOfSubItems = lngNumOfColumns - 1
    lngRowHeight = Obj.ScaleY(RowHeight, vbTwips, vbPixels)
    If TypeOf Obj Is Printer Then
        Obj.Print
        sngXFac = Screen.TwipsPerPixelX / Obj.TwipsPerPixelX
        sngYFac = Screen.TwipsPerPixelY / Obj.TwipsPerPixelY
        intPad = 3 * sngXFac
    End If  'TypeOf Obj Is
    lngBUpScalemode = Obj.scaleMode
    Obj.scaleMode = vbPixels
    intSpace_Btn_Txt_n_Pic = Obj.ScaleX(400, vbTwips, vbPixels)

    lngPosX = Obj.ScaleX(PosX, vbTwips, vbPixels)
    lngPosY = Obj.ScaleY(PosY, vbTwips, vbPixels)

    If DrawBorder Then
        lngPosX = lngPosX + (BorderDistance * sngXFac)
        lngPosY = lngPosY + (BorderDistance * sngYFac)
    End If

    With rectBox
        .Left = lngPosX
        .Top = lngPosY

        .Bottom = lngRowHeight + .Top

        For lngKount = 1 To lngNumOfColumns
            .Right = (objListView.ColumnHeaders(lngKount).Width * sngXFac) + .Left
            strTmp = objListView.ColumnHeaders(lngKount).text
            intAlign = objListView.ColumnHeaders(lngKount).Alignment
            If intAlign = 1 Then intAlign = 2 Else If intAlign = 2 Then intAlign = 1
            .Left = .Left + intPad: .Right = .Right - intPad
            .Top = .Top + (2 * sngYFac)

            If HasPicture = True And intAlign = 0 Then
                intIconNum = objListView.ColumnHeaders(lngKount).Icon
                If intIconNum <> 0 And Sgn(.Right - .Left) <> -1 Then
                    dx = (.Right - .Left) + (4 * sngXFac): dY = PicWidth
                    If dx < dY Then dY = dx
                    Obj.PaintPicture objListView.ColumnHeaderIcons.ListImages(intIconNum).Picture, .Left, .Top, (dY - 1) * sngXFac, (PicHeight - 1) * sngYFac
                    .Left = .Left + intSpace_Btn_Txt_n_Pic
                End If    'intIconNum
            End If  'HasPicture
            DrawText Obj.hdc, strTmp, Len(strTmp), rectBox, intAlign Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
            .Left = .Left - intPad: .Right = .Right + intPad
            .Top = .Top - (2 * sngYFac)
            .Left = .Right
        Next  'lngKount = 1 To lngNumOfColumns

        If DrawHorizontalLines Then
            Obj.DrawWidth = lngWidthHor
            Obj.Line (lngPosX, .Top)-(.Left, .Bottom - lngRowHeight), lngColorHor    'first line
            Obj.Line (lngPosX, .Top + lngRowHeight)-(.Left, .Bottom), lngColorHor    'second line
        End If  'DrawHorizontalLines
        If DrawBorder Then .Right = .Right + BorderDistance * sngXFac: .Bottom = .Bottom + BorderDistance * sngYFac

        If TypeOf Obj Is Printer Then
            CurrentX = .Left * Obj.TwipsPerPixelX
            CurrentY = .Top * Obj.TwipsPerPixelY
        Else
            CurrentX = .Right * Screen.TwipsPerPixelX
            CurrentY = .Bottom * Screen.TwipsPerPixelY
        End If

    End With  'RectBox
    bHeadPrinted = True
    Obj.scaleMode = lngBUpScalemode
End Sub


'  Prints the Body (Listitems and SubItems of the Listview Control)

Public Sub PrintBody(Obj As Object)
    Dim intAlign As Integer, lngCol As Long
    Dim lngRow As Long, rectBox As RECT, lngSubItems As Long, strText As String, lngBottom As Long
    Dim lngColWidth As Long, lngBUpScalemode As Long, intIconNum As Integer, intSpace_Btn_Txt_n_Pic As Integer
    Dim dx As Long, dY As Long, lngMaxRow As Long
    Dim lngPosY As Long, lngPosX As Long
    Dim lngNumOfColumns As Long, lngNumOfSubItems As Long, lngRowHeight As Long
    Dim sngXFac As Single, sngYFac As Single, intPad As Integer


    Obj.ForeColor = objListView.ForeColor
    intPad = 3: sngXFac = 1: sngYFac = 1
    lngNumOfColumns = objListView.ColumnHeaders.count
    lngNumOfSubItems = lngNumOfColumns - 1
    lngRowHeight = Obj.ScaleY(RowHeight, vbTwips, vbPixels)

    If TypeOf Obj Is Printer Then
        Obj.Print
        sngXFac = Screen.TwipsPerPixelX / Obj.TwipsPerPixelX
        sngYFac = Screen.TwipsPerPixelY / Obj.TwipsPerPixelY
        intPad = 3 * sngXFac
    End If

    lngBUpScalemode = Obj.scaleMode
    Obj.scaleMode = vbPixels
    intSpace_Btn_Txt_n_Pic = Obj.ScaleX(260, vbTwips, vbPixels)
    Obj.DrawWidth = lngWidthHor
    lngMaxRow = objListView.ListItems.count
    LastRowPrinted = False

    lngPosX = Obj.ScaleX(PosX, vbTwips, vbPixels)
    lngPosY = Obj.ScaleY(PosY, vbTwips, vbPixels)

    If DrawBorder Then
        lngPosX = lngPosX + (BorderDistance * sngXFac)
        lngPosY = lngPosY + (BorderDistance * sngYFac)
    End If

    With rectBox

        .Left = lngPosX: .Top = lngPosY

        If bHeadPrinted Then .Top = lngPosY + lngRowHeight

        For lngRow = RowFrom To RowTo
            For lngCol = 1 To lngNumOfColumns    'objListView.ListItems(lngRow).ListSubItems.Count + 1
                If lngCol = 1 Then
                    strText = objListView.ListItems(lngRow).text
                    Obj.FontBold = objListView.ListItems(lngRow).Bold
                    Obj.ForeColor = objListView.ListItems(lngRow).ForeColor          '
                Else
                    strText = objListView.ListItems(lngRow).SubItems(lngCol - 1)
                    Obj.FontBold = objListView.ListItems(lngRow).ListSubItems(lngCol - 1).Bold
                    Obj.ForeColor = objListView.ListItems(lngRow).ListSubItems(lngCol - 1).ForeColor
                End If  'lngCol = 1
                intAlign = objListView.ColumnHeaders(lngCol).Alignment
                If intAlign = 1 Then intAlign = 2 Else If intAlign = 2 Then intAlign = 1
                .Right = .Left + (objListView.ColumnHeaders(lngCol).Width * sngXFac)
                .Bottom = .Top + lngRowHeight
                .Left = .Left + intPad: .Right = .Right - intPad
                .Top = .Top + (2 * sngYFac)

                If HasPicture Then

                    If lngCol = 1 Then intIconNum = objListView.ListItems(lngRow).SmallIcon Else intIconNum = objListView.ListItems(lngRow).ListSubItems(lngCol - 1).ReportIcon
                    If intIconNum <> 0 Then
                        dx = (.Right - .Left) + (4 * sngYFac): dY = PicWidth
                        If dx < dY Then dY = dx
                        Obj.PaintPicture objListView.SmallIcons.ListImages(intIconNum).Picture, .Left, .Top, (dY - 1) * sngXFac, (PicHeight - 1) * sngYFac
                        .Left = .Left + intSpace_Btn_Txt_n_Pic
                    End If  ' intIconNum <> 0

                End If  'HasPicture

                DrawText Obj.hdc, strText, Len(strText), rectBox, intAlign Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
                Obj.FontBold = False: Obj.ForeColor = objListView.ForeColor
                .Left = .Left - intPad: .Right = .Right + intPad
                .Top = .Top - (2 * sngYFac)

                .Left = .Right

            Next  'lngCol = 1 To lngNumOfColumns
            If DrawHorizontalLines Then Obj.Line (lngPosX, .Top + lngRowHeight)-(.Left, .Bottom), lngColorHor

            .Left = lngPosX: lngBottom = .Bottom
            .Top = .Bottom
            If lngRow = lngMaxRow Then LastRowPrinted = True: Exit For
        Next  'lngRow = RowFrom To RowTo

        If DrawHorizontalLines And (Not bHeadPrinted) Then Obj.Line (lngPosX, lngPosY)-(.Right, lngPosY), lngColorHor

        If DrawVerticalLines Then
            Obj.DrawWidth = lngWidthVer
            Obj.Line (lngPosX, lngPosY)-(lngPosX, lngBottom), lngColorVer
            For lngCol = 1 To lngNumOfColumns
                lngColWidth = lngColWidth + (objListView.ColumnHeaders(lngCol).Width * sngXFac)
                Obj.Line (lngPosX + lngColWidth, lngPosY)-(lngPosX + lngColWidth, lngBottom), lngColorVer
            Next  'lngCol = 1 To lngNumOfColumns
        End If  'DrawVerticalLines

        If DrawBorder Then
            Obj.DrawWidth = lngWidthBor
            lngPosY = lngPosY - (BorderDistance * sngYFac)
            lngPosX = lngPosX - (BorderDistance * sngXFac)
            dx = .Right + (BorderDistance * sngXFac)
            dY = .Bottom + (BorderDistance * sngYFac)

            Obj.Line (dx, dY)-(dx, lngPosY), lngColorBor
            Obj.Line -(lngPosX, lngPosY), lngColorBor
            Obj.Line -(lngPosX, dY), lngColorBor
            Obj.Line -(dx, dY), lngColorBor
            .Right = .Right + (BorderDistance * sngXFac): .Bottom = .Bottom + (BorderDistance * sngYFac)
        End If  'DrawBorder
        If TypeOf Obj Is Printer Then
            CurrentX = .Right * Obj.TwipsPerPixelX
            CurrentY = .Bottom * Obj.TwipsPerPixelY
        Else
            CurrentX = .Right * Screen.TwipsPerPixelX
            CurrentY = .Bottom * Screen.TwipsPerPixelY
        End If
    End With  'rectbox
    bHeadPrinted = False
    Obj.scaleMode = lngBUpScalemode
End Sub

'Sets the ListViewControl
Public Property Set ListViewName(ByVal oNewValue As Object)
    Set objListView = oNewValue
End Property

'LineWidth and Color setting, Printer can have non integer value such as 1.2, 1.5, 2.1 etc.
Public Sub SetLine(Obj As Object, id As LineID, Optional Width As Single = 1#, Optional Color As Long = vbBlack)
    Dim intWidth As Integer
    Dim sngMul As Single
    If Width = 1 Then intWidth = 1: GoTo DownThere
    If TypeOf Obj Is Printer Then
        sngMul = 15 / Printer.TwipsPerPixelX  'Trying to make drawwidth similar in different resolution printer
        intWidth = Abs(sngMul * Width)
    Else
        intWidth = Abs(Int(Width))
        If intWidth = 0 Then intWidth = 1
    End If
DownThere:
    Select Case id
        Case HorizontalLines
            lngWidthHor = intWidth
            lngColorHor = Color
        Case VerticalLines
            lngWidthVer = intWidth
            lngColorVer = Color
        Case BorderLines
            lngWidthBor = intWidth
            lngColorBor = Color
    End Select
End Sub

'Set RowTo and RowFrom, to make printing several page easier.
Public Sub SetRows()
    RowFrom = lngLastRow
    RowTo = lngLastRow + (NumOfRowsPerPage - 1)
    lngLastRow = RowTo + 1

    If lngLastRow > objListView.ListItems.count Then lngLastRow = 1
End Sub


