VERSION 5.00
Begin VB.UserControl stCalendar 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "stCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *** Lists ***
Enum CellStyleList
    NormalBlanco
    BlueSelect
    DarkGray
    LightGray
    BevelIN
    BevelOUT
End Enum
Enum BorderStyleList
    NoBorder
    SingleBorder
End Enum
Enum LanguageList
    Dutch
    French
    English
    Spanish
End Enum
Enum CellTypeList
    header
    EmptyC   ' to make sure there'll be no conflict with vbkeyword Empty
    Normal
    Selected
End Enum
Enum SelectTypeList
    Single_Cell
    Multi_Cell
End Enum

' *** variables ***
Dim szX, szY                     ' size of a cell
Attribute szY.VB_VarUserMemId = 1073938432
Dim CellTxt() As String * 3      ' cell text
Attribute CellTxt.VB_VarUserMemId = 1073938434
Dim CellTpe() As CellTypeList    ' cell type
Attribute CellTpe.VB_VarUserMemId = 1073938435
Dim CellMrk() As Integer         ' cell marking state
Attribute CellMrk.VB_VarUserMemId = 1073938436
Dim MarkColor(4) As Long         ' mark colors
Attribute MarkColor.VB_VarUserMemId = 1073938437
Dim OffsetCell As Integer        ' starting from cell ~
Attribute OffsetCell.VB_VarUserMemId = 1073938438
Dim FirstDayWD As Integer        ' first day week day
Attribute FirstDayWD.VB_VarUserMemId = 1073938439
Dim CurrentCell As Integer       ' always one is selected
Attribute CurrentCell.VB_VarUserMemId = 1073938440

'Default Property Values:
Const m_def_SelectionType = 0
Const m_def_DayCount = 31
Const m_def_ViewEmptyCell = 3    ' LightGray
Const m_def_ViewHeaderChar = 2   ' number of characters
Const m_def_ViewHeaderLang = 2   ' English
Const m_def_ViewHeaderCell = 2   ' DarkGray
Const m_def_ViewDayCell = 0      ' NormalBlanco
Const m_def_ViewSelCell = 1      ' BlueSelect

'Property Variables:
Dim m_cYear As Integer
Attribute m_cYear.VB_VarUserMemId = 1073938441
Dim m_cMonth As Integer
Attribute m_cMonth.VB_VarUserMemId = 1073938442
Dim m_cDay As Integer
Attribute m_cDay.VB_VarUserMemId = 1073938443
Dim m_SelectionType As SelectTypeList
Attribute m_SelectionType.VB_VarUserMemId = 1073938444
Dim m_DayCount As Long
Attribute m_DayCount.VB_VarUserMemId = 1073938445
Dim m_ViewEmptyCell As CellStyleList
Attribute m_ViewEmptyCell.VB_VarUserMemId = 1073938446
Dim m_ViewHeaderChar As Long
Attribute m_ViewHeaderChar.VB_VarUserMemId = 1073938447
Dim m_ViewHeaderLang As LanguageList
Attribute m_ViewHeaderLang.VB_VarUserMemId = 1073938448
Dim m_ViewHeaderCell As CellStyleList
Attribute m_ViewHeaderCell.VB_VarUserMemId = 1073938449
Dim m_ViewDayCell As CellStyleList
Attribute m_ViewDayCell.VB_VarUserMemId = 1073938450
Dim m_ViewSelCell As CellStyleList
Attribute m_ViewSelCell.VB_VarUserMemId = 1073938451

'Event Declarations:
Event DayClicked(ByVal Button As Integer, ByVal Shift As Integer, ByVal iDay As Integer, ByRef Cancel As Boolean)
Attribute DayClicked.VB_Description = "An event passing the clicked day Set Cancel to True to prevent property  changes."
Event SelChanged()
Attribute SelChanged.VB_Description = "An event raised after DayClicked."

' (de)marks a certain calendar day
Public Sub DayMarking(ByVal iDay As Integer, ByVal MarkTpe As Integer, OnOff As Boolean)
Attribute DayMarking.VB_Description = "You can (re)set 4 markings for each day, using this function."
    Dim m As Integer
    If MarkTpe < 0 Then MarkTpe = 0
    If MarkTpe > 3 Then MarkTpe = 3
    m = CellMrk(OffsetCell + iDay)
    If OnOff Then
        m = m Or (2 ^ MarkTpe)
    Else
        m = m And Not (2 ^ MarkTpe)
    End If
    CellMrk(OffsetCell + iDay) = m
End Sub

' (de)selects a day
Public Sub DaySelect(ByVal iDay As Integer, ByVal OnOff As Boolean)
Attribute DaySelect.VB_Description = "(De)Selects a day. ReturnValue = False --> change failed."
    If OnOff Then
        CellTpe(OffsetCell + iDay) = Selected
    Else
        CellTpe(OffsetCell + iDay) = Normal
    End If
End Sub

' when a header (weekday) is clicked in Multi selection mode
Private Sub DaySelectAll(ByVal WeekcDay As Integer)
    Dim cl As Integer
    For cl = OffsetCell + 1 To OffsetCell + DayCount
        If (cl Mod 7) = WeekcDay Then
            CellTpe(cl) = Selected
        Else
            CellTpe(cl) = Normal
        End If
    Next cl
    CalendarRedraw
End Sub

Private Sub DoBevel(mX1, mY1, mX2, mY2, BT, BW)
    Dim i, Under, Above
    If BT = 0 Then
        Above = QBColor(8): Under = QBColor(15)
    Else
        Above = QBColor(15): Under = QBColor(8)
    End If
    Line (mX1, mY1)-(mX2, mY2), QBColor(7), BF
    For i = 15 To BW * 15 Step 15
        Line (mX1 + i, mY1 + i)-(mX2 - i, mY1 + i), Above
        Line (mX1 + i, mY1 + i)-(mX1 + i, mY2 - i), Above
        Line (mX2 - i, mY1 + i)-(mX2 - i, mY2 - i), Under
        Line (mX1 + i, mY2 - i)-(mX2 - i, mY2 - i), Under
    Next i
End Sub

' switch to Single selection mode
' if more than one day is selected at the present
' deselect all but current calendar day
Private Sub SetSingleSelect()
    Dim d As Integer
    For d = 1 To DayCount
        If d = m_cDay Then
            CellTpe(OffsetCell + d) = Selected
            CurrentCell = OffsetCell + d
        Else
            CellTpe(OffsetCell + d) = Normal
        End If
    Next d
    CalendarRedraw
End Sub

' multi langual weekday names
Private Sub GetHeaderText()
    Dim cl As Integer
    Select Case ViewHeaderLang
        Case Dutch
            For cl = 0 To 6
                CellTxt(cl) = Left(Choose(cl + 1, "Maa", "Din", "Woe", "Don", "Vri", "Zat", "Zon"), ViewHeaderChar)
            Next cl
        Case French
            For cl = 0 To 6
                CellTxt(cl) = Left(Choose(cl + 1, "Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"), ViewHeaderChar)
            Next cl
        Case English
            For cl = 0 To 6
                CellTxt(cl) = Left(Choose(cl + 1, "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"), ViewHeaderChar)
            Next cl
        Case Spanish
            For cl = 0 To 6
                CellTxt(cl) = Left(Choose(cl + 1, "Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "Dom"), ViewHeaderChar)
            Next cl
    End Select
    For cl = 0 To 6
        CellTpe(cl) = header
    Next cl
End Sub

' returns True if that day is selected indeed
Public Function IsDaySel(ByVal iDay As Integer) As Boolean
Attribute IsDaySel.VB_Description = "Function returns True when day is selected, False if not so."
    If CellTpe(OffsetCell + iDay) = Selected Then
        IsDaySel = True
    Else
        IsDaySel = False
    End If
End Function

' routine to draw cells in different modes
Private Sub DrawCell(ByVal x As Long, ByVal y As Long, _
                     ByVal szX As Long, ByVal szY As Long, _
                     ByVal txt As String, _
                     ByVal mode As Integer)
    Dim cx, cy
    cx = x + (szX - TextWidth(Trim(txt))) / 2
    cy = y + (szY - TextHeight(Trim(txt))) / 2

    Select Case mode
        Case NormalBlanco
            Line (x + 15, y + 15)-Step(szX - 30, szY - 30), QBColor(15), BF
            CurrentX = cx: CurrentY = cy: Print Trim(txt)
        Case BlueSelect
            Line (x + 15, y + 15)-Step(szX - 30, szY - 30), QBColor(9), BF
            Line (x, y)-Step(szX, szY), QBColor(15), B
            ForeColor = QBColor(15)
            CurrentX = cx: CurrentY = cy: Print Trim(txt)
            ForeColor = 0
        Case LightGray
            Line (x + 15, y + 15)-Step(szX - 30, szY - 30), QBColor(7), BF
            ForeColor = QBColor(15)
            CurrentX = cx: CurrentY = cy: Print Trim(txt)
            ForeColor = 0
        Case DarkGray
            Line (x, y)-Step(szX, szY), QBColor(7), BF
            Line (x + 15, y + 15)-Step(szX - 30, szY - 30), QBColor(8), BF
            ForeColor = QBColor(15)
            CurrentX = cx: CurrentY = cy: Print Trim(txt)
            ForeColor = 0
        Case BevelIN
            DoBevel x, y, x + szX, y + szY, 0, 1
            CurrentX = cx: CurrentY = cy: Print Trim(txt)
        Case BevelOUT
            DoBevel x, y, x + szX, y + szY, 1, 1
            CurrentX = cx: CurrentY = cy: Print Trim(txt)
    End Select
End Sub

' month or year changed, so reset calendar content
Private Sub CalcCalendar()
    Dim i As Integer, stat As Integer, cl As Integer
    Dim wd As Variant
    Dim mode As Integer
    Dim FirstDay As Variant
    Dim LastDay As Variant
    Dim DayCounter As Variant

    ReDim CellTxt(49)
    ReDim CellTpe(49)
    ReDim CellMrk(49)

    FirstDay = DateSerial(m_cYear, m_cMonth, 1)
    LastDay = DateSerial(m_cYear, m_cMonth + 1, 1)
    FirstDayWD = Weekday(FirstDay, vbMonday)

    GetHeaderText
    cl = 7
    i = 1: stat = 0
    While i < 43
        If stat = 0 Then                 ' first empty cells part
            If FirstDayWD = i Then        ' cell is first day?
                stat = 1                   ' start filling days
                DayCounter = FirstDay
                OffsetCell = i + 6 - 1     ' store offset
            Else
                CellTxt(cl) = "   "        ' still empty cells
                CellTpe(cl) = EmptyC
            End If
        End If
        If stat = 1 Then
            If DayCounter >= LastDay Then    ' stop at last day
                stat = 2
                CellTxt(cl) = "   "        ' again empty cells from here on
                CellTpe(cl) = EmptyC
            Else
                CellTxt(cl) = Day(DayCounter)
                If Val(CellTxt(cl)) = m_cDay Then
                    CellTpe(cl) = Selected  ' if current day Type is Selected
                    CurrentCell = cl
                Else
                    CellTpe(cl) = Normal
                End If
                DayCounter = DateAdd("d", 1, DayCounter)
            End If
        End If
        If stat = 2 Then CellTxt(cl) = "   ": CellTpe(cl) = EmptyC
        i = i + 1: cl = cl + 1
    Wend
    'debug.Print "CalcCalendar"
    m_DayCount = Day(DateAdd("d", -1, LastDay))
    CalendarRedraw
End Sub

' is like a refresh
Public Sub CalendarRedraw()
Attribute CalendarRedraw.VB_Description = "To use after the functions DaySelect, DagMarking, ..."
Attribute CalendarRedraw.VB_UserMemId = -550
    Dim cl As Integer
    Dim x, y

    Cls
    For cl = 0 To 48
        x = (cl Mod 7) * szX
        y = (cl \ 7) * szY
        Select Case CellTpe(cl)
            Case header: DrawCell x, y, szX, szY, CellTxt(cl), m_ViewHeaderCell
            Case Normal: DrawCell x, y, szX, szY, CellTxt(cl), m_ViewDayCell
            Case EmptyC: DrawCell x, y, szX, szY, CellTxt(cl), m_ViewEmptyCell
            Case Selected: DrawCell x + 15, y + 15, szX - 30, szY - 30, CellTxt(cl), m_ViewSelCell
        End Select
        If CellMrk(cl) <> 0 Then DrawMarkers x, y, CellMrk(cl)
    Next cl
    'debug.Print "CalendarRedraw"
End Sub

' routine to draw one of four types of markers
Private Sub DrawMarkers(ByVal x As Long, ByVal y As Long, _
                        ByVal Marker As Integer)
    Dim b As Integer
    Dim dx As Integer, dY As Integer
    dx = szX * 0.15
    dY = szY * 0.15
    b = 45
    If (Marker And 1) = 1 Then
        Line (x + b, y + b)-Step(dx, dY), MarkColor(0), BF
    End If
    If (Marker And 2) = 2 Then
        Line (x + szX - dx - b, y + b)-Step(dx, dY), MarkColor(1), BF
    End If
    If (Marker And 4) = 4 Then
        Line (x + szX - dx - b, y + szY - dY - b)-Step(dx, dY), MarkColor(2), BF
    End If
    If (Marker And 8) = 8 Then
        Line (x + b, y + szY - dY - b)-Step(dx, dY), MarkColor(3), BF
    End If
End Sub

Private Sub UserControl_Initialize()
    'debug.Print "initialize"
    SetMarkColors
    CalcCalendar
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim cl As Integer
    Dim Cancel As Boolean
    cl = (y \ szY) * 7 + x \ szX
    If cl > UBound(CellTxt) Or CellTpe(cl) = EmptyC Then Exit Sub

    If CellTpe(cl) <> header Then
        RaiseEvent DayClicked(Button, Shift, CInt(CellTxt(cl)), Cancel)
        If Cancel = True Then Exit Sub    ' user code canceled
    End If
    Select Case SelectionType
        Case Single_Cell
            If CellTpe(cl) = header Or cl = CurrentCell Then Exit Sub
            DrawCell (CurrentCell Mod 7) * szX, (CurrentCell \ 7) * szY, szX, szY, CellTxt(CurrentCell), ViewDayCell
            CellTpe(CurrentCell) = Normal
            DrawMarkers (CurrentCell Mod 7) * szX, (CurrentCell \ 7) * szY, CellMrk(CurrentCell)
            DrawCell (x \ szX) * szX + 15, (y \ szY) * szY + 15, szX - 30, szY - 30, CellTxt(cl), ViewSelCell
            CellTpe(cl) = Selected
            DrawMarkers (x \ szX) * szX, (y \ szY) * szY, CellMrk(cl)
            CurrentCell = cl
            m_cDay = Val(CellTxt(cl))
        Case Multi_Cell
            Select Case CellTpe(cl)
                Case header: DaySelectAll cl
                Case Selected
                    CellTpe(cl) = Normal
                    DrawCell (x \ szX) * szX, (y \ szY) * szY, szX, szY, CellTxt(cl), ViewDayCell
                Case Normal
                    CellTpe(cl) = Selected
                    DrawCell (x \ szX) * szX + 15, (y \ szY) * szY + 15, szX - 30, szY - 30, CellTxt(cl), ViewSelCell
            End Select
            If CellTpe(cl) <> header Then m_cDay = Val(CellTxt(cl))
            DrawMarkers (x \ szX) * szX, (y \ szY) * szY, CellMrk(cl)
    End Select

    RaiseEvent SelChanged
End Sub

Private Sub UserControl_Resize()
    'debug.Print "Resize"
    szX = (ScaleWidth - 15) / 7
    szY = (ScaleHeight - 15) / 7
    CalendarRedraw
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleList
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleList)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get ViewHeaderCell() As CellStyleList
Attribute ViewHeaderCell.VB_Description = "View style of header-cells or weekday names."
Attribute ViewHeaderCell.VB_ProcData.VB_Invoke_Property = ";Calendar"
    ViewHeaderCell = m_ViewHeaderCell
End Property

Public Property Let ViewHeaderCell(ByVal New_ViewHeaderCell As CellStyleList)
    m_ViewHeaderCell = New_ViewHeaderCell
    PropertyChanged "ViewHeaderCell"
    CalendarRedraw
End Property

Public Property Get ViewDayCell() As CellStyleList
Attribute ViewDayCell.VB_Description = "View style of non-selected calender day-cells."
Attribute ViewDayCell.VB_ProcData.VB_Invoke_Property = ";Calendar"
    ViewDayCell = m_ViewDayCell
End Property

Public Property Let ViewDayCell(ByVal New_ViewDayCell As CellStyleList)
    m_ViewDayCell = New_ViewDayCell
    PropertyChanged "ViewDayCell"
    CalendarRedraw
End Property

Public Property Get ViewSelCell() As CellStyleList
Attribute ViewSelCell.VB_Description = "View style of selected days/cells."
Attribute ViewSelCell.VB_ProcData.VB_Invoke_Property = ";Calendar"
    ViewSelCell = m_ViewSelCell
End Property

Public Property Let ViewSelCell(ByVal New_ViewSelCell As CellStyleList)
    m_ViewSelCell = New_ViewSelCell
    PropertyChanged "ViewSelCell"
    CalendarRedraw
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ViewHeaderCell = m_def_ViewHeaderCell
    m_ViewDayCell = m_def_ViewDayCell
    m_ViewSelCell = m_def_ViewSelCell
    m_ViewHeaderChar = m_def_ViewHeaderChar
    m_ViewHeaderLang = m_def_ViewHeaderLang
    m_ViewEmptyCell = m_def_ViewEmptyCell
    m_SelectionType = m_def_SelectionType
    m_DayCount = m_def_DayCount
    m_cYear = Year(Now)
    m_cMonth = Month(Now)
    m_cDay = Day(Now)
    CalcCalendar
    Set Font = Ambient.Font
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'debug.Print "ReadProperties"
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_ViewHeaderCell = PropBag.ReadProperty("ViewHeaderCell", m_def_ViewHeaderCell)
    m_ViewDayCell = PropBag.ReadProperty("ViewDayCell", m_def_ViewDayCell)
    m_ViewSelCell = PropBag.ReadProperty("ViewSelCell", m_def_ViewSelCell)
    m_ViewHeaderChar = PropBag.ReadProperty("ViewHeaderChar", m_def_ViewHeaderChar)
    m_ViewHeaderLang = PropBag.ReadProperty("ViewHeaderLang", m_def_ViewHeaderLang)
    m_ViewEmptyCell = PropBag.ReadProperty("ViewEmptyCell", m_def_ViewEmptyCell)
    m_SelectionType = PropBag.ReadProperty("SelectionType", m_def_SelectionType)
    m_DayCount = PropBag.ReadProperty("DayCount", m_def_DayCount)
    m_cYear = PropBag.ReadProperty("cYear", Year(Now))
    m_cMonth = PropBag.ReadProperty("cMonth", Month(Now))
    m_cDay = PropBag.ReadProperty("cDay", Day(Now))
    CalcCalendar
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub

Private Sub UserControl_Show()
    'debug.Print "Show"
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ViewHeaderCell", m_ViewHeaderCell, m_def_ViewHeaderCell)
    Call PropBag.WriteProperty("ViewDayCell", m_ViewDayCell, m_def_ViewDayCell)
    Call PropBag.WriteProperty("ViewSelCell", m_ViewSelCell, m_def_ViewSelCell)
    Call PropBag.WriteProperty("ViewHeaderChar", m_ViewHeaderChar, m_def_ViewHeaderChar)
    Call PropBag.WriteProperty("ViewHeaderLang", m_ViewHeaderLang, m_def_ViewHeaderLang)
    Call PropBag.WriteProperty("ViewEmptyCell", m_ViewEmptyCell, m_def_ViewEmptyCell)
    Call PropBag.WriteProperty("SelectionType", m_SelectionType, m_def_SelectionType)
    Call PropBag.WriteProperty("DayCount", m_DayCount, m_def_DayCount)
    Call PropBag.WriteProperty("cYear", m_cYear, Year(Now))
    Call PropBag.WriteProperty("cMonth", m_cMonth, Month(Now))
    Call PropBag.WriteProperty("cDay", m_cDay, Day(Now))
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
End Sub

Public Property Get ViewHeaderChar() As Long
Attribute ViewHeaderChar.VB_Description = "Number of character used in the header-cells  (1-3)."
Attribute ViewHeaderChar.VB_ProcData.VB_Invoke_Property = ";Calendar"
    ViewHeaderChar = m_ViewHeaderChar
End Property

Public Property Let ViewHeaderChar(ByVal New_ViewHeaderChar As Long)
    If New_ViewHeaderChar < 1 Then New_ViewHeaderChar = 1
    If New_ViewHeaderChar > 3 Then New_ViewHeaderChar = 3
    m_ViewHeaderChar = New_ViewHeaderChar
    PropertyChanged "ViewHeaderChar"
    GetHeaderText
    CalendarRedraw
End Property

Public Property Get ViewHeaderLang() As LanguageList
Attribute ViewHeaderLang.VB_Description = "The language to be used for the weekday names."
Attribute ViewHeaderLang.VB_ProcData.VB_Invoke_Property = ";Calendar"
    ViewHeaderLang = m_ViewHeaderLang
End Property

Public Property Let ViewHeaderLang(ByVal New_ViewHeaderLang As LanguageList)
    m_ViewHeaderLang = New_ViewHeaderLang
    PropertyChanged "ViewHeaderLang"
    GetHeaderText
    CalendarRedraw
End Property

Public Property Get ViewEmptyCell() As CellStyleList
Attribute ViewEmptyCell.VB_Description = "View style of Empty cells"
Attribute ViewEmptyCell.VB_ProcData.VB_Invoke_Property = ";Calendar"
    ViewEmptyCell = m_ViewEmptyCell
End Property

Public Property Let ViewEmptyCell(ByVal New_ViewEmptyCell As CellStyleList)
    m_ViewEmptyCell = New_ViewEmptyCell
    PropertyChanged "ViewEmptyCell"
    CalendarRedraw
End Property
Public Property Get SelectionType() As SelectTypeList
Attribute SelectionType.VB_Description = "Can be Single or Multi."
Attribute SelectionType.VB_ProcData.VB_Invoke_Property = ";Calendar"
    SelectionType = m_SelectionType
End Property

Public Property Let SelectionType(ByVal New_SelectionType As SelectTypeList)
    '   If Ambient.UserMode Then Err.Raise 393
    m_SelectionType = New_SelectionType
    PropertyChanged "SelectionType"
    If m_SelectionType = Single_Cell Then SetSingleSelect
End Property
Public Sub SetMarkColors(Optional m1 As Long = 16711935, Optional m2 As Long = 255, Optional m3 As Long = 16776960, Optional m4 As Long = 65280)
Attribute SetMarkColors.VB_Description = "Set marking colors. Without parameters --> reset"
    MarkColor(0) = m1
    MarkColor(1) = m2
    MarkColor(2) = m3
    MarkColor(3) = m4
End Sub

Public Property Get DayCount() As Long
Attribute DayCount.VB_Description = "Number of days in the current month, read-only."
Attribute DayCount.VB_MemberFlags = "400"
    DayCount = m_DayCount
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get cYear() As Integer
Attribute cYear.VB_Description = "The value of the calendar year."
Attribute cYear.VB_ProcData.VB_Invoke_Property = ";Calendar"
    cYear = m_cYear
End Property

Public Property Let cYear(ByVal New_Year As Integer)
    If New_Year < 1500 Then New_Year = 1500
    If New_Year > 3000 Then New_Year = 3000
    m_cYear = New_Year
    PropertyChanged "cYear"
    CalcCalendar
End Property

Public Property Get cMonth() As Integer
Attribute cMonth.VB_Description = "Value for the current calendar month."
Attribute cMonth.VB_ProcData.VB_Invoke_Property = ";Calendar"
    cMonth = m_cMonth
End Property

Public Property Let cMonth(ByVal New_cMonth As Integer)
    If New_cMonth < 1 Then New_cMonth = 1
    If New_cMonth > 12 Then New_cMonth = 12
    If m_cDay > 28 Then cDay = 1    ' for security reasons
    m_cMonth = New_cMonth
    PropertyChanged "cMonth"
    CalcCalendar
End Property

Public Property Get cDay() As Integer
Attribute cDay.VB_Description = "Last clicked calendar day-value."
Attribute cDay.VB_ProcData.VB_Invoke_Property = ";Calendar"
Attribute cDay.VB_UserMemId = 0
    cDay = m_cDay
End Property

Public Property Let cDay(ByVal New_Day As Integer)
    If New_Day > DayCount Then Exit Property
    If SelectionType = Single_Cell Then
        DaySelect m_cDay, False
        DaySelect New_Day, True
    Else
        DaySelect New_Day, IIf(IsDaySel(New_Day), False, True)
    End If
    m_cDay = New_Day
    PropertyChanged "cDay"
    CalendarRedraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    CalendarRedraw
End Property

