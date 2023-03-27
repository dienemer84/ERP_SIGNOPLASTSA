Attribute VB_Name = "VersionStuff2"
Option Explicit

' Data type to hold version information.
Public Type VersionInformationType
    StructureVersion As String
    FileVersion As String
    ProductVersion As String
    FileFlags As String
    TargetOperatingSystem As String
    FileType As String
    FileSubtype As String
End Type

' API declarations.
Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
    dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
    dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
    dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
    dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
    dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
    dwProductVersionMSl As Integer    '  e.g. = &h0003 = 3
    dwProductVersionMSh As Integer    '  e.g. = &h0010 = .1
    dwProductVersionLSl As Integer    '  e.g. = &h0000 = 0
    dwProductVersionLSh As Integer    '  e.g. = &h0031 = .31
    dwFileFlagsMask As Long        '  = &h3F for version "0.42"
    dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
    dwFileType As Long             '  e.g. VFT_DRIVER
    dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long           '  e.g. 0
    dwFileDateLS As Long           '  e.g. 0
End Type

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)

' dwFileFlags
Private Const VS_FFI_SIGNATURE = &HFEEF04BD
Private Const VS_FFI_structure_versionSION = &H10000
Private Const VS_FFI_file_flagsMASK = &H3F&

' dwFileFlags
Private Const VS_FF_DEBUG = &H1
Private Const VS_FF_PRERELEASE = &H2
Private Const VS_FF_PATCHED = &H4
Private Const VS_FF_PRIVATEBUILD = &H8
Private Const VS_FF_INFOINFERRED = &H10
Private Const VS_FF_SPECIALBUILD = &H20

' dwFileOS
Private Const VOS_UNKNOWN = &H0
Private Const VOS_DOS = &H10000
Private Const VOS_OS216 = &H20000
Private Const VOS_OS232 = &H30000
Private Const VOS_NT = &H40000
Private Const VOS_DOS_WINDOWS16 = &H10001
Private Const VOS_DOS_WINDOWS32 = &H10004
Private Const VOS_OS216_PM16 = &H20002
Private Const VOS_OS232_PM32 = &H30003
Private Const VOS_NT_WINDOWS32 = &H40004

' dwFileType
Private Const VFT_UNKNOWN = &H0
Private Const VFT_APP = &H1
Private Const VFT_DLL = &H2
Private Const VFT_DRV = &H3
Private Const VFT_FONT = &H4
Private Const VFT_VXD = &H5
Private Const VFT_STATIC_LIB = &H7

' dwFileSubtype for drivers
Private Const VFT2_UNKNOWN = &H0
Private Const VFT2_DRV_PRINTER = &H1
Private Const VFT2_DRV_KEYBOARD = &H2
Private Const VFT2_DRV_LANGUAGE = &H3
Private Const VFT2_DRV_DISPLAY = &H4
Private Const VFT2_DRV_MOUSE = &H5
Private Const VFT2_DRV_NETWORK = &H6
Private Const VFT2_DRV_SYSTEM = &H7
Private Const VFT2_DRV_INSTALLABLE = &H8
Private Const VFT2_DRV_SOUND = &H9
Private Const VFT2_DRV_COMM = &HA
' Return version information strings for a file.
Public Function VersionInformation1(ByVal file_name As String) As VersionInformationType
    Dim dummy_handle As Long
    Dim buffer() As Byte
    Dim info_size As Long
    Dim info_address As Long
    Dim fixed_file_info As VS_FIXEDFILEINFO
    Dim fixed_file_info_size As Long
    Dim result As VersionInformationType

    ' Get the version information buffer size.
    info_size = GetFileVersionInfoSize(file_name, dummy_handle)
    If info_size = 0 Then
        MsgBox "No version information available"
        Exit Function
    End If

    ' Load the fixed file information into a buffer.
    ReDim buffer(1 To info_size)
    If GetFileVersionInfo(file_name, 0&, info_size, buffer(1)) = 0 Then
        MsgBox "Error getting version information"
        Exit Function
    End If
    If VerQueryValue(buffer(1), "\", info_address, fixed_file_info_size) = 0 Then
        MsgBox "Error getting fixed file version information"
        Exit Function
    End If

    ' Copy the information from the buffer into a
    ' usable structure.
    MoveMemory fixed_file_info, info_address, Len(fixed_file_info)

    ' Get the version information.
    With fixed_file_info
        ' Structure version.
        result.StructureVersion = _
        Format$(.dwStrucVersionh) & "." & _
                                  Format$(.dwStrucVersionl)

        ' File version number.
        result.FileVersion = _
        Format$(.dwFileVersionMSh) & "." & _
                             Format$(.dwFileVersionMSl) & "." & _
                             Format$(.dwFileVersionLSh) & "." & _
                             Format$(.dwFileVersionLSl)

        ' Product version number.
        result.ProductVersion = _
        Format$(.dwProductVersionMSh) & "." & _
                                Format$(.dwProductVersionMSl) & "." & _
                                Format$(.dwProductVersionLSh) & "." & _
                                Format$(.dwProductVersionLSl)

        ' File attributes.
        result.FileFlags = ""
        If .dwFileFlags And VS_FF_DEBUG Then result.FileFlags = result.FileFlags & " Debug"
        If .dwFileFlags And VS_FF_PRERELEASE Then result.FileFlags = result.FileFlags & " PreRel"
        If .dwFileFlags And VS_FF_PATCHED Then result.FileFlags = result.FileFlags & " Patched"
        If .dwFileFlags And VS_FF_PRIVATEBUILD Then result.FileFlags = result.FileFlags & " Private"
        If .dwFileFlags And VS_FF_INFOINFERRED Then result.FileFlags = result.FileFlags & " Info"
        If .dwFileFlags And VS_FF_SPECIALBUILD Then result.FileFlags = result.FileFlags & " Special"
        If .dwFileFlags And VFT2_UNKNOWN Then result.FileFlags = result.FileFlags + " Unknown"
        If Len(result.FileFlags) > 0 Then result.FileFlags = Mid$(result.FileFlags, 2)

        ' Target operating system.
        Select Case .dwFileOS
        Case VOS_DOS_WINDOWS16
            result.TargetOperatingSystem = "DOS-Win16"
        Case VOS_DOS_WINDOWS32
            result.TargetOperatingSystem = "DOS-Win32"
        Case VOS_OS216_PM16
            result.TargetOperatingSystem = "OS/2-16 PM-16"
        Case VOS_OS232_PM32
            result.TargetOperatingSystem = "OS/2-16 PM-32"
        Case VOS_NT_WINDOWS32
            result.TargetOperatingSystem = "NT-Win32"
        Case Else
            result.TargetOperatingSystem = "Unknown"
        End Select

        ' File type.
        Select Case .dwFileType
        Case VFT_APP
            result.FileType = "App"
        Case VFT_DLL
            result.FileType = "DLL"
        Case VFT_DRV
            result.FileType = "Driver"
            Select Case fixed_file_info.dwFileSubtype
            Case VFT2_DRV_PRINTER
                result.FileSubtype = "Printer drv"
            Case VFT2_DRV_KEYBOARD
                result.FileSubtype = "Keyboard drv"
            Case VFT2_DRV_LANGUAGE
                result.FileSubtype = "Language drv"
            Case VFT2_DRV_DISPLAY
                result.FileSubtype = "Display drv"
            Case VFT2_DRV_MOUSE
                result.FileSubtype = "Mouse drv"
            Case VFT2_DRV_NETWORK
                result.FileSubtype = "Network drv"
            Case VFT2_DRV_SYSTEM
                result.FileSubtype = "System drv"
            Case VFT2_DRV_INSTALLABLE
                result.FileSubtype = "Installable"
            Case VFT2_DRV_SOUND
                result.FileSubtype = "Sound drv"
            Case VFT2_DRV_COMM
                result.FileSubtype = "Comm drv"
            Case VFT2_UNKNOWN
                result.FileSubtype = "Unknown"
            End Select
        Case VFT_FONT
            result.FileType = "Font"
        Case VFT_VXD
            result.FileType = "VxD"
        Case VFT_STATIC_LIB
            result.FileType = "Lib"
        Case Else
            result.FileType = "Unknown"
        End Select
    End With

    VersionInformation1 = result
End Function

