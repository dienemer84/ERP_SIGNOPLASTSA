Attribute VB_Name = "FormHelper"
Option Explicit

Public Const FondoCeleste As Long = &HF0E1E1
'Public Const FondoCeleste As Long = &HFFDBBF

Public Const FondoAzul As Long = &HD6AA88
Public Const LetraAzul As Long = &H8B4215
Public Const PreviewFondoCeleste As Long = &HE9FB9E
Public Const GridBGHeader As Long = &HF7DCC8


Public Const TWIP_CM As Double = 0.001763889

Public Const PIXEL_CM As Double = 0.026458333
Public Function ConvertTwipToCm(ValueInTwips As Long) As Double
    ConvertTwipToCm = ValueInTwips * TWIP_CM
End Function

Public Function ConvertCmToTwip(ValueInCm As Long) As Double
    ConvertCmToTwip = ValueInCm / TWIP_CM
End Function


Public Function ConvertPixelToCm(ValueInPixels As Long) As Double
    ConvertPixelToCm = ValueInPixels * PIXEL_CM
End Function

Public Function ConvertCmToPixel(ValueInCm As Long) As Double
    ConvertCmToPixel = ValueInCm / PIXEL_CM
End Function



Public Sub Customize(F As Form)
    On Error Resume Next

    F.backColor = FondoCeleste
    F.Font = "Tahoma"

    Set F.Icon = Nothing

    Dim ctrl As Control
    Dim but As CommandButton
    Dim rep As ReportControl
    Dim push As PushButton
    Dim gbox As Xtremesuitecontrols.GroupBox
    Dim cbox As Xtremesuitecontrols.ComboBox
    Dim rb As Xtremesuitecontrols.RadioButton


    For Each ctrl In F.Controls
        'http://www.elguille.info/colabora/vb2005/galegre_ClearTypeVB.htm
        'CustomizeControl ctrl
        ctrl.Font = "MS Shell Dlg 2"    '"Tahoma"

        If TypeOf ctrl Is Label Or TypeOf ctrl Is Frame Or TypeOf ctrl Is OptionButton Or TypeOf ctrl Is CheckBox Or TypeOf ctrl Is Xtremesuitecontrols.CheckBox Or TypeOf ctrl Is Xtremesuitecontrols.Label Then

            ctrl.backColor = FondoCeleste
            ctrl.ForeColor = LetraAzul

        ElseIf TypeOf ctrl Is CommandButton Then
            Dim fontt As New stdfont
            fontt.Bold = False
            'http://www.elguille.info/colabora/vb2005/galegre_ClearTypeVB.htm
            fontt.Name = "MS Shell Dlg 2"    '"Tahoma"

            Set but = ctrl
            Set but.Font = fontt
            but.backColor = FondoAzul
        ElseIf TypeOf ctrl Is ComboBox Then
            ctrl.backColor = vbWhite
            ctrl.Sorter = True
        ElseIf TypeOf ctrl Is ComboBox Then
            ctrl.backColor = &H80000005
        ElseIf TypeOf ctrl Is ReportControl Then
            Dim fontt2 As New stdfont
            fontt2.Bold = False
            fontt2.Name = "MS Shell Dlg 2"    '"Tahoma"
            'fontt2.Size = 8 '8 por defecto

            Set rep = ctrl
            rep.PaintManager.CaptionBackColor = FormHelper.GridBGHeader
            rep.PaintManager.CaptionForeColor = FormHelper.LetraAzul
            Set rep.PaintManager.TextFont = fontt2


            rep.PaintManager.HorizontalGridStyle = xtpGridSmallDots
            rep.PaintManager.VerticalGridStyle = xtpGridSmallDots

            rep.PaintManager.PreviewTextColor = FormHelper.LetraAzul

            If rep.PaintManager.NoItemsText = "There are no items to show." Then
                rep.PaintManager.NoItemsText = "No hay items para mostrar"
            End If

            rep.PaintManager.HighlightBackColor = Information.RGB(62, 157, 232)    '3E 9D E8
        ElseIf TypeOf ctrl Is PushButton Then
            Set push = ctrl
            If push.ForeColor <> LetraAzul Then push.ForeColor = LetraAzul
            If push.Appearance <> xtpAppearanceOffice2007 Then push.Appearance = xtpAppearanceOffice2007

        ElseIf TypeOf ctrl Is Xtremesuitecontrols.GroupBox Then
            Set gbox = ctrl
            gbox.backColor = FormHelper.FondoCeleste
        ElseIf TypeOf ctrl Is Xtremesuitecontrols.ComboBox Then
            Set cbox = ctrl
            cbox.Appearance = xtpAppearanceOffice2003    'xtpAppearanceOffice2007
            cbox.AutoComplete = True
            cbox.Sorted = True
            cbox.Style = xtpComboDropDownList
            cbox.UseVisualStyle = True
        ElseIf TypeOf ctrl Is Xtremesuitecontrols.RadioButton Then
            Set rb = ctrl
            rb.Appearance = xtpAppearanceOffice2007
            rb.backColor = FormHelper.FondoCeleste
        ElseIf TypeOf ctrl Is Xtremesuitecontrols.CheckBox Then
            ctrl.backColor = FormHelper.FondoCeleste


        End If

    Next ctrl

End Sub

Private Sub CustomizeControl(ctrl As Control)
    On Error Resume Next

    If TypeOf ctrl Is Label Or TypeOf ctrl Is Frame Then
        ctrl.backColor = FondoCeleste
    ElseIf TypeOf ctrl Is CommandButton Then
        ctrl.backColor = FondoAzul
    End If

    Dim ctrl2 As Control
    For Each ctrl2 In ctrl.Controls
        CustomizeControl ctrl2
    Next ctrl2
End Sub

