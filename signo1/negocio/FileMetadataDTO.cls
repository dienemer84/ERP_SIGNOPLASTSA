VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileMetadataDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public filename As String
Public DirectoryName As String
Public FileSize As Long    'in bytes

Public Property Get FileSizeInKB() As Double
    FileSizeInKB = Round(Me.FileSize / 1024, 2)
End Property

Public Property Get FileSizeInMB() As Double
    FileSizeInMB = Round((Me.FileSize / 1024) / 1024, 2)
End Property

Public Property Get FullFilePath() As String
    Dim tmp As String
    tmp = Me.DirectoryName
    If Right(Me.DirectoryName, 1) <> "\" Then tmp = tmp & "\"
    tmp = tmp & Me.filename
    FullFilePath = tmp
End Property
