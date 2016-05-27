VERSION 5.00
Begin VB.UserControl ACPRibbon 
   Alignable       =   -1  'True
   BackColor       =   &H00404040&
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   ControlContainer=   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   7095
   Begin VB.Label ButMouse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   0
      Left            =   2040
      TabIndex        =   4
      ToolTipText     =   "çlll"
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Glip_on 
      Height          =   60
      Index           =   0
      Left            =   2280
      Top             =   1560
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Glip_off 
      Height          =   60
      Index           =   0
      Left            =   2160
      Top             =   1560
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Button_left_over 
      Height          =   990
      Index           =   0
      Left            =   2520
      Top             =   1800
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image Button_center_over 
      Height          =   990
      Index           =   0
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Button_right_over 
      Height          =   990
      Index           =   0
      Left            =   3480
      Top             =   1800
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image Cat_Dlg_over 
      Height          =   210
      Index           =   0
      Left            =   4080
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Cat_Dlg_on 
      Height          =   210
      Index           =   0
      Left            =   3840
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Cat_Dlg 
      Height          =   210
      Index           =   0
      Left            =   3600
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Button_Icon 
      Appearance      =   0  'Flat
      Height          =   495
      Index           =   0
      Left            =   1320
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Button_Caption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image Button_right 
      Height          =   990
      Index           =   0
      Left            =   1920
      Top             =   1800
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image Button_center 
      Height          =   990
      Index           =   0
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Button_left 
      Height          =   990
      Index           =   0
      Left            =   960
      Top             =   1800
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label TabMouse 
      Height          =   360
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Tab_caption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aba 01"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image Tab_right 
      Height          =   360
      Index           =   0
      Left            =   2280
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Tab_center 
      Height          =   360
      Index           =   0
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Tab_left 
      Height          =   360
      Index           =   0
      Left            =   1680
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Tab_left_over 
      Height          =   360
      Index           =   0
      Left            =   1680
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Tab_center_over 
      Height          =   360
      Index           =   0
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Tab_right_over 
      Height          =   360
      Index           =   0
      Left            =   2280
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label CatMouse 
      Height          =   1350
      Index           =   0
      Left            =   4560
      TabIndex        =   3
      Top             =   150
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Cat_Caption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   5040
      TabIndex        =   2
      Tag             =   "sadf"
      Top             =   1200
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image Cat_Right_on 
      Height          =   1335
      Index           =   0
      Left            =   6120
      Top             =   150
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Cat_Center_on 
      Height          =   1335
      Index           =   0
      Left            =   5880
      Stretch         =   -1  'True
      Top             =   150
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Cat_Left_on 
      Height          =   1335
      Index           =   0
      Left            =   5760
      Top             =   150
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Cat_Right_off 
      Height          =   1335
      Index           =   0
      Left            =   5400
      Top             =   150
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Cat_Left_off 
      Height          =   1335
      Index           =   0
      Left            =   5040
      Top             =   150
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Cat_Center_off 
      Height          =   1335
      Index           =   0
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   150
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image BarraLeft 
      Height          =   2130
      Left            =   0
      Top             =   0
      Width           =   165
   End
   Begin VB.Image BarraRight 
      Height          =   2130
      Left            =   480
      Top             =   0
      Width           =   165
   End
   Begin VB.Image Barra2 
      Height          =   2130
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   405
   End
End
Attribute VB_Name = "ACPRibbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'#######################################
'#                                     #
'#           ACP Ribbon 2007           #
'#            (Ribbon Only)            #
'#                  by                 #
'#      adrianopaladini@gmail.com      #
'#                                     #
'#                                     #
'#  Visual from Office 2007 Beta 2 TR  #
'#                                     #
'#   Please Don´t Remove Author Info!  #
'#                                     #
'#######################################


'------------------------------------------------
' Bugs:
'
' Please report to:
'
'         adrianopaladini@gmail.com
'
'------------------------------------------------


Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_PATH = 260

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Dim TotalButton As Integer
Dim TotalTabs As Integer
Dim TotalCats As Integer
Dim TabSelected As String
Dim TabID(30) As String
Dim TabC(30) As String
Dim CatsID(30) As String
Dim CatsC(30) As String
Dim CatsT(30) As String
Dim CatsD(30) As Boolean

Dim TopBuID(90) As String
Dim TopBuS(90) As String
Dim TopBuC(90) As String
Dim TopBuI(90) As Picture
Dim TopBuT(90) As String
Dim TopBuG(90) As Boolean

Dim ms As Boolean
Dim Mx, My As Integer
Attribute My.VB_VarUserMemId = 1073938449
Event TabClick(ByVal Id As String, ByVal Caption As String)
Event CatClick(ByVal Id As String, ByVal Caption As String)
Event ButtonClick(ByVal Id As String, ByVal Caption As String)
Const m_def_Theme = 0
Const m_def_BC = False
Dim m_Theme As Variant
Attribute m_Theme.VB_VarUserMemId = 1073938451
Dim m_BC As Boolean
Attribute m_BC.VB_VarUserMemId = 1073938452
Dim zImg As ImageList
Attribute zImg.VB_VarUserMemId = 1073938453

Dim TAB_NORMAL
Attribute TAB_NORMAL.VB_VarUserMemId = 1073938454
Dim TAB_SELECTED
Attribute TAB_SELECTED.VB_VarUserMemId = 1073938455
Private Sub TabNone(Optional index As Integer = -1)
    If index <> -1 Then
        For i = 0 To index - 1
            If Tab_center_over(i).Visible = True Then
                Tab_center_over(i).Visible = False
                Tab_left_over(i).Visible = False
                Tab_right_over(i).Visible = False
            End If
        Next
        If Tab_center(index).Visible = False Then
            Tab_center_over(index).Visible = True
            Tab_left_over(index).Visible = True
            Tab_right_over(index).Visible = True
        End If
        For i = index + 1 To TabMouse.UBound
            If Tab_center_over(i).Visible = True Then
                Tab_center_over(i).Visible = False
                Tab_left_over(i).Visible = False
                Tab_right_over(i).Visible = False
            End If
        Next
    Else
        For i = 0 To TabMouse.UBound
            If Tab_center_over(i).Visible = True Then
                Tab_center_over(i).Visible = False
                Tab_left_over(i).Visible = False
                Tab_right_over(i).Visible = False
            End If
        Next
    End If
End Sub
Private Sub CatNone(Optional index As Integer = -1)
    If index <> -1 Then
        For i = 0 To index - 1
            If Cat_Center_on(i).Visible = True Then
                Cat_Center_on(i).Visible = False
                Cat_Left_on(i).Visible = False
                Cat_Right_on(i).Visible = False
                If Cat_Dlg(i).Visible = True Then
                    Cat_Dlg_on(i).Visible = False
                    Cat_Dlg_over(i).Visible = False
                End If
            End If
        Next
        Cat_Center_on(index).Visible = True
        Cat_Left_on(index).Visible = True
        Cat_Right_on(index).Visible = True
        If Cat_Dlg(index).Visible = True Then
            Cat_Dlg_on(index).Visible = True
            Cat_Dlg_over(index).Visible = False
        End If
        For i = index + 1 To CatMouse.UBound
            If Cat_Center_on(i).Visible = True Then
                Cat_Center_on(i).Visible = False
                Cat_Left_on(i).Visible = False
                Cat_Right_on(i).Visible = False
                If Cat_Dlg(i).Visible = True Then
                    Cat_Dlg_on(i).Visible = False
                    Cat_Dlg_over(i).Visible = False
                End If
            End If
        Next
    Else
        For i = 0 To CatMouse.UBound
            If Cat_Center_on(i).Visible = True Then
                Cat_Center_on(i).Visible = False
                Cat_Left_on(i).Visible = False
                Cat_Right_on(i).Visible = False
                If Cat_Dlg(i).Visible = True Then
                    Cat_Dlg_on(i).Visible = False
                    Cat_Dlg_over(i).Visible = False
                End If
            End If
        Next
    End If
End Sub
Private Sub ButNone(Optional index As Integer = -1)
    If index <> -1 Then
        For KL = 0 To index - 1
            If Button_center(KL).Visible = True Then
                Button_left(KL).Visible = False
                Button_right(KL).Visible = False
                Button_center(KL).Visible = False
                If Glip_off(i).Visible = True Then
                    Glip_on(i).Visible = False
                End If
            End If
        Next
        If Button_left(index).Visible = False Then
            Button_left(index).Visible = True
            Button_center(index).Visible = True
            Button_right(index).Visible = True
            If Glip_off(index).Visible = True Then
                Glip_on(index).Visible = True
            End If
        End If
        For KL = index + 1 To ButMouse.UBound
            If Button_center(KL).Visible = True Then
                Button_left(KL).Visible = False
                Button_right(KL).Visible = False
                Button_center(KL).Visible = False
                If Glip_off(i).Visible = True Then
                    Glip_on(i).Visible = False
                End If
            End If
        Next
    Else
        For KL = 0 To ButMouse.UBound
            If Button_center(KL).Visible = True Then
                Button_left(KL).Visible = False
                Button_right(KL).Visible = False
                Button_center(KL).Visible = False
                If Glip_off(i).Visible = True Then
                    Glip_on(i).Visible = False
                End If
            End If
        Next
    End If
End Sub
Private Sub Barra2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TabNone
    CatNone
    ButNone
End Sub
Private Sub BarraLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TabNone
    CatNone
    ButNone
End Sub
Private Sub BarraRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TabNone
    CatNone
    ButNone
End Sub
Private Sub ButMouse_Click(index As Integer)
    RaiseEvent ButtonClick(ButMouse(index).Tag, Button_Caption(index).Caption)
End Sub
Private Sub ButMouse_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Button_left_over(index).Visible = True
    Button_center_over(index).Visible = True
    Button_right_over(index).Visible = True
End Sub
Private Sub ButMouse_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TabNone
    CatNone Button_center(index).Tag
    ButNone index
End Sub
Private Sub ButMouse_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Button_left_over(index).Visible = False
    Button_center_over(index).Visible = False
    Button_right_over(index).Visible = False
End Sub
Private Sub Cat_Dlg_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TabNone
    CatNone index
    ButNone
End Sub
Private Sub Cat_Dlg_on_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TabNone
    CatNone index
    ButNone
    Cat_Dlg_over(index).Visible = True
End Sub
Private Sub Cat_Dlg_over_Click(index As Integer)
    RaiseEvent CatClick(Cat_Caption(index).Tag, Cat_Caption(index).Caption)
End Sub
Private Sub CatMouse_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TabNone
    CatNone index
    ButNone
End Sub
Private Sub TabMouse_Click(index As Integer)
    TabNone
    For i = 0 To index - 1
        Tab_center(i).Visible = False
        Tab_left(i).Visible = False
        Tab_right(i).Visible = False
        Tab_caption(i).ForeColor = TAB_NORMAL
    Next
    Tab_caption(index).ForeColor = TAB_SELECTED
    Tab_center(index).Visible = True
    Tab_left(index).Visible = True
    Tab_right(index).Visible = True
    For i = index + 1 To TabMouse.UBound
        Tab_center(i).Visible = False
        Tab_left(i).Visible = False
        Tab_right(i).Visible = False
        Tab_caption(i).ForeColor = TAB_NORMAL
    Next
    TabSelected = TabID(index)
    CatsUpdate
    RaiseEvent TabClick(TabID(index), TabC(index))
End Sub
Private Sub TabMouse_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TabNone index
    CatNone
    ButNone
End Sub
Private Sub UserControl_Initialize()

    Barra2.Top = -(26 * 15)
    BarraLeft.Top = Barra2.Top
    BarraRight.Top = Barra2.Top

    UserControl.Height = Barra2.Height
    Barra2.Width = 2048 * 15
    TotalTopButton = 0
    TotalButton = 0
    TotalTabs = 0
    TotalCats = 0
    TabSelected = ""
    TabMouse(0).BackStyle = 0
    CatMouse(0).BackStyle = 0
    ButMouse(0).BackStyle = 0
End Sub
Private Sub TabsUpdate()
    On Error Resume Next
    For i = 1 To (TotalTabs - 1)
        Unload Tab_caption(i)
        Unload Tab_left(i)
        Unload Tab_center(i)
        Unload Tab_right(i)
        Unload Tab_left_over(i)
        Unload Tab_center_over(i)
        Unload Tab_right_over(i)
        Unload TabMouse(i)
    Next
    For i = 0 To (TotalTabs - 1)
        If i <> 0 Then
            Load Tab_caption(i)
            Load Tab_left(i)
            Load Tab_center(i)
            Load Tab_right(i)
            Load Tab_left_over(i)
            Load Tab_center_over(i)
            Load Tab_right_over(i)
            Load TabMouse(i)
            Tab_left(i).Left = Tab_right(i - 1).Left + Tab_right(i).Width
        Else
            Tab_left(0).Left = 90
        End If
        TabMouse(i).Left = Tab_left(i).Left

        Tab_caption(i).Top = 0 + 60
        Tab_center(i).Top = 0
        Tab_left(i).Top = 0
        Tab_right(i).Top = 0
        Tab_center_over(i).Top = 0
        Tab_left_over(i).Top = 0
        Tab_right_over(i).Top = 0
        TabMouse(i).Top = 0

        Tab_caption(i) = TabC(i)
        Tab_center(i).Width = Tab_caption(i).Width
        Tab_center(i).Left = Tab_left(i).Left + Tab_left(i).Width
        Tab_caption(i).Left = Tab_center(i).Left
        Tab_right(i).Left = Tab_center(i).Left + Tab_center(i).Width

        Tab_center_over(i).Width = Tab_center(i).Width
        Tab_center_over(i).Left = Tab_center(i).Left
        Tab_left_over(i).Left = Tab_left(i).Left
        Tab_right_over(i).Left = Tab_right(i).Left

        TabMouse(i).Width = Tab_left(i).Width + Tab_right(i).Width + Tab_center(i).Width

        Tab_caption(i).ForeColor = TAB_NORMAL

        Tab_caption(i).Visible = True
        If i = 0 Then
            Tab_center(i).Visible = True
            Tab_left(i).Visible = True
            Tab_right(i).Visible = True
            Tab_caption(i).ForeColor = TAB_SELECTED
        End If
        TabMouse(i).Visible = True

        Tab_center(i).ZOrder 0
        Tab_left(i).ZOrder 0
        Tab_right(i).ZOrder 0

        Tab_center_over(i).ZOrder 0
        Tab_left_over(i).ZOrder 0
        Tab_right_over(i).ZOrder 0

        Tab_caption(i).ZOrder 0
        TabMouse(i).ZOrder 0
    Next
End Sub
Private Sub CatsUpdate()
    On Error Resume Next
    ztopo = 360
    Cat_Center_off(0).Top = ztopo
    Cat_Center_on(0).Top = ztopo
    Cat_Left_off(0).Top = ztopo
    Cat_Left_on(0).Top = ztopo
    Cat_Right_off(0).Top = ztopo
    Cat_Right_on(0).Top = ztopo
    CatMouse(0).Top = ztopo
    Cat_Caption(0).Top = 1400

    Dim TotalCatsT As Integer
    Dim CatsIDT(30) As String
    Dim CatsCT(30) As String
    Dim CatsTT(30) As String
    Dim CatsDT(30) As Boolean
    TotalCatsT = 0
    For i = 0 To TotalCats
        If CatsT(i) = TabSelected And TabSelected <> "" And CatsT(i) <> "" Then
            CatsIDT(TotalCatsT) = CatsID(i)
            CatsTT(TotalCatsT) = CatsT(i)
            CatsCT(TotalCatsT) = CatsC(i)
            CatsDT(TotalCatsT) = CatsD(i)
            TotalCatsT = TotalCatsT + 1
        End If
    Next
    For i = 1 To CatMouse.UBound
        Unload Cat_Left_off(i)
        Unload Cat_Left_on(i)
        Unload Cat_Right_off(i)
        Unload Cat_Right_on(i)
        Unload Cat_Center_off(i)
        Unload Cat_Center_on(i)
        Unload Cat_Caption(i)
        Unload CatMouse(i)
        Unload Cat_Dlg(i)
        Unload Cat_Dlg_on(i)
        Unload Cat_Dlg_over(i)
    Next
    For i = 1 To Button_center.UBound
        Unload Button_left(i)
        Unload Button_center(i)
        Unload Button_right(i)
        Unload Button_left_over(i)
        Unload Button_center_over(i)
        Unload Button_right_over(i)
        Unload Button_Caption(i)
        Unload Button_Icon(i)
        Unload Glip_on(i)
        Unload Glip_off(i)
        Unload ButMouse(i)
    Next
    Button_left(0).Visible = False
    Button_center(0).Visible = False
    Button_right(0).Visible = False
    Button_Caption(0).Visible = False
    Button_Icon(0).Visible = False
    ButMouse(0).Visible = False

    Cat_Left_off(0).Visible = False
    Cat_Left_on(0).Visible = False
    Cat_Right_off(0).Visible = False
    Cat_Right_on(0).Visible = False
    Cat_Center_off(0).Visible = False
    Cat_Center_on(0).Visible = False
    Cat_Caption(0).Visible = False
    CatMouse(0).Visible = False
    Cat_Dlg(0).Visible = False
    Cat_Dlg_on(0).Visible = False
    Cat_Dlg_over(0).Visible = False
    For i = 0 To (TotalCatsT - 1)
        If i <> 0 Then
            Load Cat_Left_off(i)
            Load Cat_Left_on(i)
            Load Cat_Right_off(i)
            Load Cat_Right_on(i)
            Load Cat_Center_off(i)
            Load Cat_Center_on(i)
            Load Cat_Caption(i)
            Load CatMouse(i)
            Load Cat_Dlg(i)
            Load Cat_Dlg_on(i)
            Load Cat_Dlg_over(i)
            Cat_Left_off(i).Left = Cat_Right_off(i - 1).Left + Cat_Right_off(i).Width
        Else
            Cat_Left_off(i).Left = 120
        End If
        CatMouse(i).Left = Cat_Left_off(i).Left

        Cat_Caption(i).Caption = CatsCT(i)
        Cat_Caption(i).Tag = CatsIDT(i)

        Cat_Center_off(i).Left = Cat_Left_off(i).Left + Cat_Left_off(i).Width

        BUTSIZE = ButtonsUpdate(CatsIDT(i), Cat_Center_off(i).Left, i + 0)

        If CatsDT(i) = True Then
            Cat_Center_off(i).Width = Cat_Caption(i).Width + Cat_Dlg(i).Width
        Else
            Cat_Center_off(i).Width = Cat_Caption(i).Width
        End If

        If Cat_Center_off(i).Width < BUTSIZE Then
            Cat_Center_off(i).Width = BUTSIZE
            Cat_Caption(i).Left = Cat_Center_off(i).Left + ((Cat_Center_off(i).Width - Cat_Caption(i).Width) / 2)
        Else
            Cat_Caption(i).Left = Cat_Center_off(i).Left
        End If

        Cat_Right_off(i).Left = Cat_Center_off(i).Left + Cat_Center_off(i).Width

        Cat_Center_on(i).Width = Cat_Center_off(i).Width
        Cat_Center_on(i).Left = Cat_Center_off(i).Left
        Cat_Left_on(i).Left = Cat_Left_off(i).Left
        Cat_Right_on(i).Left = Cat_Right_off(i).Left

        CatMouse(i).Width = Cat_Left_off(i).Width + Cat_Right_off(i).Width + Cat_Center_off(i).Width

        Cat_Caption(i).Visible = True
        Cat_Center_off(i).Visible = True
        Cat_Left_off(i).Visible = True
        Cat_Right_off(i).Visible = True
        CatMouse(i).Visible = True

        Cat_Center_off(i).ZOrder 0
        Cat_Left_off(i).ZOrder 0
        Cat_Right_off(i).ZOrder 0

        Cat_Center_on(i).ZOrder 0
        Cat_Left_on(i).ZOrder 0
        Cat_Right_on(i).ZOrder 0

        Cat_Caption(i).ZOrder 0
        CatMouse(i).ZOrder 0

        Cat_Dlg(i).Left = (Cat_Right_off(i).Left - Cat_Dlg(i).Width) + 15
        Cat_Dlg(i).Top = (Cat_Right_off(i).Top + Cat_Right_off(i).Height) - (Cat_Dlg(i).Height + 60)

        Cat_Dlg_on(i).Left = Cat_Dlg(i).Left
        Cat_Dlg_over(i).Left = Cat_Dlg(i).Left

        Cat_Dlg_on(i).Top = Cat_Dlg(i).Top
        Cat_Dlg_over(i).Top = Cat_Dlg(i).Top


        Cat_Dlg_on(i).Visible = False
        Cat_Dlg_over(i).Visible = False

        If CatsDT(i) = True Then
            Cat_Dlg(i).Visible = True
        End If
        Cat_Dlg(i).ZOrder 0
        Cat_Dlg_on(i).ZOrder 0
        Cat_Dlg_over(i).ZOrder 0
    Next
    DoEvents
    For KL = 0 To ButMouse.UBound
        Button_left(KL).Visible = False
        Button_left(KL).ZOrder 0
        Button_right(KL).Visible = False
        Button_right(KL).ZOrder 0
        Button_center(KL).Visible = False
        Button_center(KL).ZOrder 0

        Button_left_over(KL).Visible = False
        Button_left_over(KL).ZOrder 0
        Button_right_over(KL).Visible = False
        Button_right_over(KL).ZOrder 0
        Button_center_over(KL).Visible = False
        Button_center_over(KL).ZOrder 0

        Button_Icon(KL).ZOrder 0
        Button_Caption(KL).ZOrder 0

        Glip_off(KL).ZOrder 0
        Glip_on(KL).ZOrder 0

        ButMouse(KL).ZOrder 0
    Next
End Sub
Private Sub UserControl_Resize()
'On Error Resume Next
    UserControl.Height = Barra2.Height - (26 * 15)
    'UserControl.Width = UserControl.ParentControls.Item(0).ScaleWidth
    'BarraRight.Left = UserControl.Width - BarraRight.Width
End Sub
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl_Resize
    TabsUpdate
    CatsUpdate
End Sub
Private Sub UserControl_InitProperties()
    m_Theme = m_def_Theme
    m_BC = m_def_BC
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Theme = PropBag.ReadProperty("Theme", m_def_Theme)
    m_BC = PropBag.ReadProperty("ButtonCenter", m_def_BC)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Theme", m_Theme, m_def_Theme)
    Call PropBag.WriteProperty("ButtonCenter", m_BC, m_def_BC)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H464646)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &HFFFFFF)
End Sub
Public Function AddTab(zID As String, zCaption As String) As Boolean
    TotalTabs = TotalTabs + 1
    TabID(TotalTabs - 1) = zID
    zCaption = Replace(zCaption, vbNewLine, " ")
    TabC(TotalTabs - 1) = zCaption
    If TabSelected = "" Then
        TabSelected = zID
    End If
End Function
Public Function AddCat(zID As String, zTab As String, zCaption As String, zDlgButton As Boolean) As Boolean
    TotalCats = TotalCats + 1
    CatsID(TotalCats - 1) = zID
    CatsT(TotalCats - 1) = zTab
    zCaption = Replace(zCaption, vbNewLine, " ")
    CatsC(TotalCats - 1) = zCaption
    CatsD(TotalCats - 1) = zDlgButton
End Function
Public Function AddButton(zID As String, zSubCat As String, zCaption As String, zPicture As Integer, Optional zMore As Boolean = False, Optional zToolTip As String) As Boolean
    TotalButton = TotalButton + 1
    TopBuID(TotalButton - 1) = zID
    TopBuS(TotalButton - 1) = zSubCat
    TopBuC(TotalButton - 1) = zCaption
    If zToolTip = "" Or zToolTip = Null Then
        If InStr(zCaption, vbNewLine) Then
            zCaption = Replace(zCaption, vbNewLine, " ")
        End If
        TopBuT(TotalButton - 1) = zCaption
    Else
        zToolTip = Replace(zToolTip, vbNewLine, " ")
        TopBuT(TotalButton - 1) = zToolTip
    End If
    Set TopBuI(TotalButton - 1) = zImg.ListImages.Item(zPicture).Picture
    TopBuG(TotalButton - 1) = zMore
End Function
Private Function ButtonsUpdate(SubCat As String, PosIni As Integer, CatID As Integer) As Integer
    On Error Resume Next
    Dim TotalButtonT As Integer
    Dim TopBuIDT(90) As String
    Dim TopBuST(90) As String
    Dim TopBuCT(90) As String
    Dim TopBuIT(90) As Picture
    Dim TopBuTT(90) As String
    Dim TopBuGT(90) As Boolean
    TotalSize = 0
    TotalButtonT = 0
    For i = 0 To TotalButton
        If TopBuS(i) = SubCat Then
            TopBuIDT(TotalButtonT) = TopBuID(i)
            TopBuST(TotalButtonT) = TopBuS(i)
            TopBuCT(TotalButtonT) = TopBuC(i)
            TopBuTT(TotalButtonT) = TopBuT(i)
            Set TopBuIT(TotalButtonT) = TopBuI(i)
            TopBuGT(TotalButtonT) = TopBuG(i)
            TotalButtonT = TotalButtonT + 1
        End If
    Next
    Button_left(0).Visible = False
    Button_center(0).Visible = False
    Button_right(0).Visible = False
    Button_Caption(0).Visible = True
    Button_Icon(0).Visible = True
    ButMouse(0).Visible = True
    xt = ButMouse.UBound + 1
    For i = xt To (TotalButtonT - 1) + xt
        If i <> 0 Then
            Load Button_left(i)
            Load Button_center(i)
            Load Button_right(i)
            Load Button_left_over(i)
            Load Button_center_over(i)
            Load Button_right_over(i)
            Load Button_Caption(i)
            Load Button_Icon(i)
            Load Glip_on(i)
            Load Glip_off(i)
            Load ButMouse(i)
        End If
        ButMouse(i).Tag = TopBuIDT(i - xt)

        Button_center(i).Tag = CatID

        ButMouse(i).Top = Cat_Left_off(0).Top + 60
        Button_left(i).Top = ButMouse(i).Top
        Button_center(i).Top = ButMouse(i).Top
        Button_right(i).Top = ButMouse(i).Top
        Button_left_over(i).Top = ButMouse(i).Top
        Button_center_over(i).Top = ButMouse(i).Top
        Button_right_over(i).Top = ButMouse(i).Top

        If i = xt Then
            posatu = PosIni
        Else
            posatu = ButMouse(i - 1).Left + ButMouse(i - 1).Width + 30
        End If
        ButMouse(i).Left = posatu
        Button_left(i).Left = ButMouse(i).Left
        Button_left_over(i).Left = Button_left(i).Left
        Button_center(i).Left = Button_left(i).Left + Button_left(i).Width
        Button_center_over(i).Left = Button_center(i).Left

        Button_Caption(i).Caption = TopBuCT(i - xt)

        Set Button_Icon(i) = TopBuIT(i - xt)

        If m_BC = True Then
            esp = Button_center(i).Height - (Button_Icon(i).Height + Button_Caption(i).Height)
            If TopBuGT(i - xt) = True Then
                Button_Icon(i).Top = Button_center(i).Top + ((esp - (Button_Caption(i).Height / 2)) / 2)
            Else
                Button_Icon(i).Top = Button_center(i).Top + ((esp) / 2)
            End If
        Else
            Button_Icon(i).Top = Button_center(i).Top + 90
        End If


        Button_Caption(i).Top = Button_Icon(i).Top + Button_Icon(i).Height

        Glip_off(i).Top = Button_Caption(i).Top + Button_Caption(i).Height + ((Button_Caption(i).Height - Glip_off(i).Height) / 2)
        Glip_on(i).Top = Glip_off(i).Top


        If Button_Caption(i).Width > Button_Icon(i).Width Then
            Button_Caption(i).Left = Button_center(i).Left
            esp2 = (Button_Caption(i).Width - Button_Icon(i).Width) / 2
            Button_Icon(i).Left = Button_Caption(i).Left + esp2
            Area = Button_Caption(i).Width
        Else
            Button_Icon(i).Left = Button_center(i).Left
            esp2 = (Button_Icon(i).Width - Button_Caption(i).Width) / 2
            Button_Caption(i).Left = Button_Icon(i).Left + esp2
            Area = Button_Icon(i).Width
        End If

        Glip_off(i).Left = Button_Caption(i).Left + ((Button_Caption(i).Width - Glip_on(i).Width) / 2)
        Glip_on(i).Left = Glip_off(i).Left

        Button_center(i).Width = Area
        Button_center_over(i).Width = Button_center(i).Width
        Button_right(i).Left = Button_center(i).Left + Button_center(i).Width
        Button_right_over(i).Left = Button_right(i).Left
        ButMouse(i).Width = (Button_right(i).Width + Button_right(i).Width) + Button_center(i).Width

        ButMouse(i).ToolTipText = TopBuTT(i - xt)
        Button_Icon(i).Visible = True
        Button_Caption(i).Visible = True
        ButMouse(i).Visible = True
        If TopBuGT(i - xt) = True Then
            Glip_off(i).Visible = True
            Glip_off(i).ZOrder 0
            Glip_on(i).ZOrder 0
        End If

        TotalSize = TotalSize + ButMouse(i).Width + 30
    Next
    ButtonsUpdate = TotalSize - 30
End Function
Public Property Get Theme() As Integer
Attribute Theme.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Theme = m_Theme
End Property
Public Property Let Theme(ByVal New_Theme As Integer)
    If New_Theme < 0 Or New_Theme > 2 Then New_Theme = 0
    m_Theme = New_Theme
    PropertyChanged "Theme"
    LoadTheme m_Theme
End Property
Public Property Get ButtonCenter() As Variant
    ButtonCenter = m_BC
End Property
Public Property Let ButtonCenter(ByVal New_BC As Variant)
    m_BC = New_BC
    PropertyChanged "ButtonCenter"
End Property
Private Function LoadTheme(iTema)
    Select Case iTema
        Case 0
            Id = "BLACK"
            Cat_Caption(0).ForeColor = &HFFFFFF
            TAB_NORMAL = vbWhite
            TAB_SELECTED = vbBlack
            Button_Caption(0).ForeColor = &H80000008
            UserControl.BackColor = &H464646
            UserControl.ForeColor = &HFFFFFF
        Case 1
            Id = "BLUE"
            Cat_Caption(0).ForeColor = &HB86A3E
            TAB_NORMAL = &H8B4215
            TAB_SELECTED = &H8B4215
            Button_Caption(0).ForeColor = &H8B4215
            UserControl.BackColor = &HDAB08E
            UserControl.ForeColor = &H8B4215
        Case 2
            Id = "SILVER"
            Cat_Caption(0).ForeColor = &H6A625C
            TAB_NORMAL = &H6A625C
            TAB_SELECTED = &H6A625C
            Button_Caption(0).ForeColor = &H6A625C
            UserControl.BackColor = &HDDD4D0
            UserControl.ForeColor = &H6A625C
        Case Else
            Id = "BLACK"
    End Select
    Set Barra2.Picture = LoadResPicture(101, Id)
    Set BarraLeft.Picture = LoadResPicture(102, Id)
    Set BarraRight.Picture = LoadResPicture(103, Id)
    Set Cat_Dlg(0).Picture = LoadResPicture(118, Id)
    Set Cat_Dlg_on(0).Picture = LoadResPicture(119, Id)
    Set Cat_Dlg_over(0).Picture = LoadResPicture(120, Id)
    Set Cat_Left_off(0).Picture = LoadResPicture(121, Id)
    Set Cat_Center_off(0).Picture = LoadResPicture(122, Id)
    Set Cat_Right_off(0).Picture = LoadResPicture(123, Id)
    Set Cat_Left_on(0).Picture = LoadResPicture(124, Id)
    Set Cat_Center_on(0).Picture = LoadResPicture(125, Id)
    Set Cat_Right_on(0).Picture = LoadResPicture(126, Id)
    Set Tab_left(0).Picture = LoadResPicture(127, Id)
    Set Tab_center(0).Picture = LoadResPicture(128, Id)
    Set Tab_right(0).Picture = LoadResPicture(129, Id)
    Set Tab_left_over(0).Picture = LoadResPicture(130, Id)
    Set Tab_center_over(0).Picture = LoadResPicture(131, Id)
    Set Tab_right_over(0).Picture = LoadResPicture(132, Id)
    Set Glip_off(0).Picture = LoadResPicture(133, Id)
    Set Glip_on(0).Picture = LoadResPicture(134, Id)
    Set Button_left_over(0).Picture = LoadResPicture(135, Id)
    Set Button_center_over(0).Picture = LoadResPicture(136, Id)
    Set Button_right_over(0).Picture = LoadResPicture(137, Id)
    Set Button_left(0).Picture = LoadResPicture(138, Id)
    Set Button_center(0).Picture = LoadResPicture(139, Id)
    Set Button_right(0).Picture = LoadResPicture(140, Id)
End Function
Private Property Get TempDir() As String
    Dim sRet As String, c As Long
    Dim lErr As Long
    sRet = String$(MAX_PATH, 0)
    c = GetTempPath(MAX_PATH, sRet)
    lErr = Err.LastDllError
    If c = 0 Then
        Err.Raise 10000 Or lErr, App.EXEName & ".cAniCursor", WinAPIError(lErr)
    End If
    TempDir = Left$(sRet, c)
End Property
Private Property Get TempFileName(Optional ByVal sPrefix As String, Optional ByVal sPathName As String) As String
    Dim lErr As Long
    Dim iPos As Long
    If sPrefix = "" Then sPrefix = ""
    If sPathName = "" Then sPathName = TempDir
    Dim sRet As String
    sRet = String(MAX_PATH, 0)
    GetTempFileName sPathName, sPrefix, 0, sRet
    lErr = Err.LastDllError
    If Not lErr = 0 Then
        Err.Raise 10000 Or lErr, App.EXEName & ".cAniCursor", WinAPIError(lErr)
    End If
    iPos = InStr(sRet, vbNullChar)
    If Not iPos = 0 Then
        TempFileName = Left$(sRet, iPos - 1)
    End If
End Property
Private Function WinAPIError(ByVal lLastDLLError As Long) As String
    Dim sBuff As String
    Dim lCount As Long
    sBuff = String$(256, 0)
    lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
    If lCount Then
        WinAPIError = Left$(sBuff, lCount)
    End If
End Function
Public Property Get LoadBackground() As IPicture
    Dim sFile As String
    Dim b() As Byte
    Dim iFile As Integer
    On Error GoTo errorhandler
    Select Case m_Theme
        Case 0
            b = LoadResData(141, "BLACK")
        Case 1
            b = LoadResData(141, "BLUE")
        Case 2
            b = LoadResData(141, "SILVER")
    End Select
    sFile = TempFileName("LRP")
    iFile = FreeFile
    Open sFile For Binary Access Write Lock Read As #iFile
    Put #iFile, , b
    Close #iFile
    iFile = 0
    Set LoadBackground = LoadPicture(sFile)
    KillFile sFile
    Exit Property
errorhandler:
    Dim lErr As Long, sErr As String
    lErr = Err.Number: sErr = Err.Description
    If Not iFile = 0 Then Close #iFile
    KillFile sFile
    Err.Raise Err.Number, App.EXEName & ".cLoadResPicture", Err.Description
    Exit Property
End Property
Private Property Get LoadResPicture(ByVal Id As Variant, ByVal Format As Variant) As IPicture
    Dim sFile As String
    Dim b() As Byte
    Dim iFile As Integer
    On Error GoTo errorhandler
    b = LoadResData(Id, Format)
    sFile = TempFileName("LRP")
    iFile = FreeFile
    Open sFile For Binary Access Write Lock Read As #iFile
    Put #iFile, , b
    Close #iFile
    iFile = 0
    Set LoadResPicture = LoadPicture(sFile)
    KillFile sFile
    Exit Property
errorhandler:
    Dim lErr As Long, sErr As String
    lErr = Err.Number: sErr = Err.Description
    If Not iFile = 0 Then Close #iFile
    KillFile sFile
    Err.Raise Err.Number, App.EXEName & ".cLoadResPicture", Err.Description
    Exit Property
End Property
Private Sub KillFile(ByVal sFile As String)
    On Error Resume Next
    Kill sFile
End Sub
Public Sub Resize()
    UserControl_Resize
End Sub
Public Property Let ImageList(ByVal zImageList As ImageList)
    Set zImg = zImageList
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property
