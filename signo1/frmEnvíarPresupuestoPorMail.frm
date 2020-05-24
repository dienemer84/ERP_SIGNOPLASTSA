VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmEnviarMail 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enviar E-Mail"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7845
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1560
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   6615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7815
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   315
         Left            =   6360
         TabIndex        =   24
         Top             =   2400
         Width           =   1275
      End
      Begin VB.TextBox txtServert 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   360
         Width           =   4215
      End
      Begin VB.ComboBox cboPriority 
         Height          =   315
         ItemData        =   "frmEnvíarPresupuestoPorMail.frx":0000
         Left            =   6240
         List            =   "frmEnvíarPresupuestoPorMail.frx":0002
         TabIndex        =   10
         Text            =   "cboPriority"
         ToolTipText     =   "Sets the Prioirty of the Mail Message"
         Top             =   1800
         Width           =   1410
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Envíar"
         Default         =   -1  'True
         Height          =   315
         Left            =   6240
         TabIndex        =   8
         Top             =   345
         Width           =   1275
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Nuevo"
         Height          =   315
         Left            =   6300
         TabIndex        =   9
         Top             =   765
         Width           =   1275
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   1680
         Width           =   4200
      End
      Begin VB.TextBox txtFromName 
         Height          =   285
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   945
         Width           =   4200
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   1830
         TabIndex        =   2
         Top             =   1305
         Width           =   4200
      End
      Begin VB.TextBox txtSubject 
         Height          =   285
         Left            =   1860
         TabIndex        =   5
         Top             =   2505
         Width           =   4200
      End
      Begin VB.TextBox txtMsg 
         Height          =   1620
         Left            =   1860
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2865
         Width           =   4200
      End
      Begin VB.TextBox txtAttach 
         Height          =   285
         Left            =   1860
         TabIndex        =   7
         Top             =   4545
         Width           =   4200
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   315
         Left            =   6300
         TabIndex        =   11
         Top             =   4560
         Width           =   1275
      End
      Begin VB.TextBox txtCc 
         Height          =   285
         Left            =   1830
         TabIndex        =   4
         Top             =   2040
         Width           =   4200
      End
      Begin VB.ListBox lstStatus 
         BackColor       =   &H8000000F&
         Height          =   1035
         Left            =   1860
         TabIndex        =   13
         Top             =   4965
         Width           =   4200
      End
      Begin MSComDlg.CommonDialog cmDialog 
         Left            =   600
         Top             =   3705
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblServer 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servidor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   780
         TabIndex        =   23
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destinatario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   22
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label lblFromName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remitente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   585
         TabIndex        =   21
         Top             =   1005
         Width           =   870
      End
      Begin VB.Label lblFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responder A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   435
         TabIndex        =   20
         Top             =   1365
         Width           =   1110
      End
      Begin VB.Label lblSubject 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asunto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   465
         TabIndex        =   19
         Top             =   2505
         Width           =   600
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   405
         TabIndex        =   18
         Top             =   2865
         Width           =   720
      End
      Begin VB.Label lblAttach 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attachment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   17
         Top             =   4605
         Width           =   975
      End
      Begin VB.Label lblCC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1275
         TabIndex        =   16
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Progreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3705
         TabIndex        =   15
         Top             =   6165
         Width           =   780
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   855
         TabIndex        =   14
         Top             =   5025
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmEnviarMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' *****************************************************************************
' Required declaration of the vbSendMail component (withevents is optional)
' You also need a reference to the vbSendMail component in the Project References
' *****************************************************************************
Private WithEvents poSendMail As vbSendMail.clsSendMail
Attribute poSendMail.VB_VarHelpID = -1
' misc local vars
Dim clsPersonal As New ClassPersonal
Dim errores As Boolean
Dim bAuthLogin      As Boolean
Dim bPopLogin       As Boolean
Dim bHtml           As Boolean
Dim MyEncodeType    As ENCODE_METHOD
Dim etPriority      As MAIL_PRIORITY
Dim bReceipt        As Boolean




Private Sub cmdSend_Click()
On Error GoTo err22
    ' *****************************************************************************
    ' This is where all of the Components Properties are set / Methods called
    ' *****************************************************************************

    cmdSend.Enabled = False
    lstStatus.Clear
    Screen.MousePointer = vbHourglass

    With poSendMail

        ' **************************************************************************
        ' Optional properties for sending email, but these should be set first
        ' if you are going to use them
        ' **************************************************************************

        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)

        ' **************************************************************************
        ' Basic properties for sending email
        ' **************************************************************************
        .SMTPHost = Me.txtServert                  ' Required the fist time, optional thereafter
        .From = txtFrom.Text                        ' Required the fist time, optional thereafter
        .FromDisplayName = txtFromName.Text         ' Optional, saved after first use
        .Recipient = txtTo.Text                     ' Required, separate multiple entries with delimiter character
        
        .CcRecipient = txtCc                        ' Optional, separate multiple entries with delimiter character
        
        
        .ReplyToAddress = txtFrom.Text              ' Optional, used when different than 'From' address
        .Subject = txtSubject.Text                  ' Optional
        .Message = txtMsg.Text                      ' Optional
        .Attachment = Trim(txtAttach.Text)          ' Optional, separate multiple entries with delimiter character

        ' **************************************************************************
        ' Additional Optional properties, use as required by your application / environment
        ' **************************************************************************
        .AsHTML = bHtml                             ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = bReceipt                         ' Optional, default = FALSE
        .UseAuthentication = bAuthLogin             ' Optional, default = FALSE
        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        
        
        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised
        
        ' **************************************************************************
        ' Advanced Properties, change only if you have a good reason to do so.
        ' **************************************************************************
        ' .ConnectTimeout = 10                      ' Optional, default = 10
        ' .ConnectRetry = 5                         ' Optional, default = 5
        ' .MessageTimeout = 60                      ' Optional, default = 60
        ' .PersistentSettings = True                ' Optional, default = TRUE
        ' .SMTPPort = 25                            ' Optional, default = 25

        ' **************************************************************************
        ' OK, all of the properties are set, send the email...
        ' **************************************************************************
        ' .Connect                                  ' Optional, use when sending bulk mail
        .Send
        If Not errores Then
         clsPersonal.addMensajes Now, Trim(Me.txtTo), Trim(Me.txtAttach), Trim(Me.txtMsg), Trim(Me.txtSubject), funciones.getUser
        End If
        ' Required
        ' .Disconnect                               ' Optional, use when sending bulk mail
        
        
        txtServert = .SMTPHost                  ' Optional, re-populate the Host in case
                                                    ' MX look up was used to find a host    End With
    End With
    Screen.MousePointer = vbDefault
    cmdSend.Enabled = True
Exit Sub
err22:
MsgBox Err.Description
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

' *****************************************************************************
' The following four Subs capture the Events fired by the vbSendMail component
' *****************************************************************************

Private Sub poSendMail_Progress(lPercentCompete As Long)

    ' vbSendMail 'Progress Event'
    lblProgress = lPercentCompete & "% completo"

End Sub

Private Sub poSendMail_SendFailed(Explanation As String)

    ' vbSendMail 'SendFailed Event
    MsgBox ("Error al enviar un email: " & vbCrLf & Explanation)
    lblProgress = ""
    Screen.MousePointer = vbDefault
    cmdSend.Enabled = True
    errores = True
    
End Sub

Private Sub poSendMail_SendSuccesful()

    ' vbSendMail 'SendSuccesful Event'
    MsgBox "Send Successful!"
    lblProgress = ""
    errores = False

End Sub

Private Sub poSendMail_Status(Status As String)

    ' vbSendMail 'Status Event'
    lstStatus.AddItem Status
    lstStatus.ListIndex = lstStatus.ListCount - 1
    lstStatus.ListIndex = -1

End Sub

Private Sub Form_Load()
Exit Sub
Me.txtServert = funciones.getServerSMTPe
Me.txtServert.Locked = True
Me.txtFrom.Locked = True
Me.txtFromName.Locked = True


Dim classE As New ClassPersonal
Dim nom As String, ape As String, mail As String
classE.ejecutar "select p.apellido, p.nombre,p.email from sp.personal p inner join sp.usuarios u on u.idEmpleado=p.id and u.id=" & funciones.getUser
nom = classE.nombre
ape = classE.apellido
mail = classE.email
Set classE = Nothing
frmEnviarMail.txtFrom = mail
frmEnviarMail.txtFromName = ape & ", " & nom
    
    
    ' *****************************************************************************
    ' Required to activate the vbSendMail component.
    ' *****************************************************************************
    Set poSendMail = New clsSendMail
    cboPriority.AddItem "Normal"
    cboPriority.AddItem "Alta"
    cboPriority.AddItem "Baja"
    cboPriority.ListIndex = 0

    
    
    

    Me.Show

'    RetrieveSavedValues

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' *****************************************************************************
    ' Unload the component before quiting.
    ' *****************************************************************************

    Set poSendMail = Nothing

End Sub

Private Sub RetrieveSavedValues()

    ' *****************************************************************************
    ' Retrieve saved values by reading the components 'Persistent' properties
    ' *****************************************************************************
    poSendMail.PersistentSettings = True
    txtServert.Text = poSendMail.SMTPHost

    txtFrom.Text = poSendMail.From
    txtFromName.Text = poSendMail.FromDisplayName

    'optEncodeType(poSendMail.EncodeType).value = True
    

End Sub


Private Sub cboPriority_Click()

    Select Case cboPriority.ListIndex

        Case 0: etPriority = NORMAL_PRIORITY
        Case 1: etPriority = HIGH_PRIORITY
        Case 2: etPriority = LOW_PRIORITY

    End Select

End Sub

Private Sub cboPriority_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode

        Case 38, 40

        Case Else: KeyCode = 0

    End Select

End Sub

Private Sub cboPriority_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


'Private Sub ckLogin_Click()
'
'    If ckLogin.value = vbChecked Then
'        bAuthLogin = True
'        fraOptions.Height = 3555
'    Else
'        bAuthLogin = False
'        If ckPopLogin.value = vbUnchecked Then fraOptions.Height = 2475
'    End If
'
'End Sub

'Private Sub ckPopLogin_Click()
'
'    If ckPopLogin.value = vbChecked Then
'        bPopLogin = True
'        lblPopServer.Visible = True
'        txtPopServer.Visible = True
'        fraOptions.Height = 3555
'    Else
'        bPopLogin = False
'        lblPopServer.Visible = False
'        txtPopServer.Visible = False
'        If ckLogin.value = vbUnchecked Then fraOptions.Height = 2475
'    End If
'
'End Sub



Private Sub cmdBrowse_Click()

    Dim sFilenames()    As String
    Dim I               As Integer
    
    On Local Error GoTo Err_Cancel
  
    With cmDialog
        .filename = ""
        .CancelError = True
        .Filter = "All Files (*.*)|*.*|HTML Files (*.htm;*.html;*.shtml)|*.htm;*.html;*.shtml|Images (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
        .FilterIndex = 1
        .DialogTitle = "Seleccione adjuntos"
        .MaxFileSize = &H7FFF
        .Flags = &H4 Or &H800 Or &H40000 Or &H200 Or &H80000
        .ShowOpen
        ' get the selected name(s)
        sFilenames = Split(.filename, vbNullChar)
    End With
    
    If UBound(sFilenames) = 0 Then
        If txtAttach.Text = "" Then
            txtAttach.Text = sFilenames(0)
        Else
            txtAttach.Text = txtAttach.Text & ";" & sFilenames(0)
        End If
    ElseIf UBound(sFilenames) > 0 Then
        If Right$(sFilenames(0), 1) <> "\" Then sFilenames(0) = sFilenames(0) & "\"
        For I = 1 To UBound(sFilenames)
            If txtAttach.Text = "" Then
                txtAttach.Text = sFilenames(0) & sFilenames(I)
            Else
                txtAttach.Text = txtAttach.Text & ";" & sFilenames(0) & sFilenames(I)
            End If
        Next
    Else
        Exit Sub
    End If
    
Err_Cancel:

End Sub

Private Sub cmdExit_Click()

Dim frm As Form

For Each frm In Forms
    Unload frm
    Set frm = Nothing
Next

End

End Sub

Private Sub cmdReset_Click()

    ClearTextBoxesOnForm
    lstStatus.Clear
    lblProgress = ""
    RetrieveSavedValues

End Sub

Private Sub AlignControlsLeft(StandardizeWidth As Boolean, base As Object, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com
    On Error Resume Next

    Dim I As Integer
    For I = 0 To UBound(cnts)
        cnts(I).Left = base.Left
        If StandardizeWidth Then cnts(I).Width = base.Width
    Next

End Sub

Private Sub CenterControlsVertical(space As Single, AlignLeft As Boolean, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    Dim sngTotalSpace As Single
    Dim I As Integer
    Dim sngBaseLeft As Single

    Dim sngParentHeight As Single

    sngParentHeight = Me.ScaleHeight

    For I = 0 To UBound(cnts)
        sngTotalSpace = sngTotalSpace + cnts(I).Height
    Next

    sngTotalSpace = sngTotalSpace + (space * (UBound(cnts)))
    cnts(0).Top = (sngParentHeight - sngTotalSpace) / 2

    sngBaseLeft = cnts(0).Left

    For I = 1 To UBound(cnts)
        cnts(I).Top = cnts(I - 1).Top + cnts(I - 1).Height + space
        If AlignLeft Then cnts(I).Left = sngBaseLeft
    Next

End Sub

Private Sub CenterControlHorizontal(child As Object)

    child.Left = (Me.ScaleWidth - child.Width) / 2

End Sub

Public Sub CenterControlsHorizontal(space As Single, AlignTop As Boolean, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    Dim sngTotalSpace As Single
    Dim I As Integer
    Dim sngBaseTop As Single
    Dim sngParentWidth As Single

    sngParentWidth = Me.ScaleWidth

    For I = 0 To UBound(cnts)
        sngTotalSpace = sngTotalSpace + cnts(I).Width
    Next

    sngTotalSpace = sngTotalSpace + (space * (UBound(cnts)))

    cnts(0).Left = (sngParentWidth - sngTotalSpace) / 2
    sngBaseTop = cnts(0).Top

    For I = 1 To UBound(cnts)
        cnts(I).Left = cnts(I - 1).Left + cnts(I - 1).Width + space
        If AlignTop Then cnts(I).Top = sngBaseTop
    Next

End Sub

Public Sub AlignControlsTop(StandardizeHeight As Boolean, base As Object, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    Dim I As Integer
    For I = 0 To UBound(cnts)
        cnts(I).Top = base.Top
        If StandardizeHeight Then cnts(I).Height = base.Height
    Next

End Sub

Public Sub CenterControlRelativeVertical(ctl As Object, RelativeTo As Object)

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    ctl.Top = RelativeTo.Top + ((RelativeTo.Height - ctl.Height) / 2)

End Sub

Public Sub SetHorizontalDistance(distance As Single, StandardizeWidth As Boolean, AlignTop As Boolean, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    Dim I As Integer
    For I = 1 To UBound(cnts)
        If StandardizeWidth Then cnts(I).Width = cnts(I - 1).Width
        cnts(I).Left = cnts(I - 1).Left + cnts(I - 1).Width + distance
        If AlignTop Then cnts(I).Top = cnts(I - 1).Top
    Next

End Sub

Public Sub CenterControlsRelativeHorizontal(RelativeTo As Object, space As Single, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    Dim sngTotalWidth As Single
    Dim I As Integer
    For I = 0 To UBound(cnts)
        sngTotalWidth = sngTotalWidth + cnts(I).Width
        If I < UBound(cnts) Then sngTotalWidth = sngTotalWidth + space
    Next

    cnts(0).Left = RelativeTo.Left + ((RelativeTo.Width - sngTotalWidth) / 2)

    For I = 1 To UBound(cnts)
        cnts(I).Left = cnts(I - 1).Left + cnts(I - 1).Width + space
        cnts(I).Top = cnts(0).Top
    Next

End Sub

Public Sub ClearTextBoxesOnForm()

    ' Snippet Taken From http://www.freevbcode.com

    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        End If
    Next

End Sub


