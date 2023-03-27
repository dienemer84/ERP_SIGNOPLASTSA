Attribute VB_Name = "APIACC"
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Sub ChangeRegionalSettings()

'http://www.vbforums.com/showthread.php?t=379056
'http://msdn.microsoft.com/en-us/library/dd464799%28v=vs.85%29.aspx
'Private Const LOCALE_SDECIMAL = &HE '14
'Private Const LOCALE_STHOUSAND = &HF '15
'Private Const LOCALE_SMONDECIMALSEP = &H16 '22
'Private Const LOCALE_SMONTHOUSANDSEP = &H17 '23


'setea las configuraciones reginales para que el sistema ande ok
    SetRegionalSetting 14, "."    'deberia ir en "," pero se caga todo el sistema, no esta internacionalizado
    SetRegionalSetting 15, "."
    SetRegionalSetting 22, ","
    SetRegionalSetting 23, "."
End Sub

Private Function SetRegionalSetting(ByVal lRegionalSetting As Long, ByVal value As String)
    SetRegionalSetting = SetLocaleInfo(Locale_User_Default, lRegionalSetting, value)
End Function




Public Function LeerIni(lpFileName As String, lpAppName As String, lpKeyName As String, Optional vDefault) As String
'Los parámetros son:
'lpFileName:    La Aplicación (fichero INI)
'lpAppName:     La sección que suele estar entrre corchetes
'lpKeyName:     Clave
'vDefault:      Valor opcional que devolverá
'               si no se encuentra la clave.
'
    Dim lpString As String
    Dim ltmp As Long
    Dim sRetVal As String

    'Si no se especifica el valor por defecto,
    'asignar incialmente una cadena vacía
    If IsMissing(vDefault) Then
        lpString = ""
    Else
        lpString = vDefault
    End If

    sRetVal = String$(255, 0)

    ltmp = GetPrivateProfileString(lpAppName, lpKeyName, lpString, sRetVal, Len(sRetVal), lpFileName)
    If ltmp = 0 Then
        LeerIni = lpString
    Else
        LeerIni = Left(sRetVal, ltmp)
    End If
End Function




Sub GuardarIni(lpFileName As String, lpAppName As String, lpKeyName As String, lpString As String)
'Guarda los datos de configuración
'Los parámetros son los mismos que en LeerIni
'Siendo lpString el valor a guardar
'
    Dim ltmp As Long

    ltmp = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub




