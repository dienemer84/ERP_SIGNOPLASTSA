'########################################################################
'# PROCESO:     dbMakeBackup
'# OBJETO:      Procedimiento para hacer un volcado de una base de datos MySQL,
'#                  a un archivo plano
'# PRECONDICIONES:
'#              Debe existir una variable global llamada 'cnn' del tipo ADODB.Connection
'#                  la cual está apuntando a la base de datos que se quiere respaldar
'# PARAMETROS:
'# (obligatorio)strFileName => Nombre del archivo plano donde se desea dejar el volcado.
'# (opcional)   IncludeCreateDB => Incluye en el volcado los comandos
'#                  DROP DATABASE y el CREATE DATABASE
'# (opcional)   IncludeStructure => Incluye en el volcado la creación de las estructuras de
'#                  tablas de la base de datos.
'# (opcional)   IncludeData => Incluye en el volcado los datos de cada tabla.
'#
'# RETORNA:     n/a
'#
'# AUTOR:       Williams Castillo - will@eduven.com
'# FECHA ULTIMA MODIFICACION: 24/01/2006
'#
'# POR HACER:   Está desarrollado para versiones anteriores a la 5.0
'#              Falta incluir los stored procedures y triggers para hacerla
'#                  100% compatible con MySQL 5.0 y posteriores.
'########################################################################
Public Sub dbMakeBackup(ByVal strFileName As String, Optional IncludeCreateDB As Boolean = True, Optional IncludeStructure As Boolean = True, Optional IncludeData As Boolean = True)
Dim rss As ADODB.Recordset
Dim rssAux As ADODB.Recordset

Dim x As Long, I As Integer

Dim strTableName As String
Dim strCurLine As String
Dim strBuffer As String
Dim strDBName As String
    
On Error Resume Next
    
    x = FreeFile
    Open strFileName For Output As x
    
    Print #x, ""
    Print #x, "#"
    
    Print #x, "# Respaldo creado por: "; App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
    strDBName = VBA.mID(cnn.ConnectionString, InStr(cnn.ConnectionString, "DATABASE=") + 9)
    strDBName = left(strDBName, InStr(strDBName, ";") - 1)
    Print #x, "# Base datos: " & strDBName
    
    Set rss = New ADODB.Recordset
    Set rssAux = New ADODB.Recordset
    
    'Print #X, "# Fecha/Hora: " & Format(Now, "DD/MM/YYYY HH:MM:SS")
    rss.Open "SHOW VARIABLES LIKE 'version';", cnn
    If Not rss.EOF Then
        Print #x, "# DBMS: MySQL v" & rss.Fields(1)
    End If
    rss.Close
    
    Print #x, "#"
    If IncludeData Then
        Print #x, ""
        Print #x, "SET FOREIGN_KEY_CHECKS=0;"
    End If
        Print #x, ""
    
    If IncludeCreateDB Then
        Print #x, "DROP DATABASE IF EXISTS `" & strDBName & "`;"
        Print #x, "CREATE DATABASE `" & strDBName & "`;"
    End If
    Print #x, "USE `" & strDBName & "`;"
    
    strTableName = ""


    With rss
        .Open "SHOW TABLE STATUS", cnn

        Do While Not .EOF
            strTableName = .Fields.Item("Name").Value
            
            If IncludeStructure Then
                With rssAux
                    
                    .Open "SHOW CREATE TABLE " & strTableName, cnn
                    Print #x, ""
                    Print #x, "#"
                    Print #x, "# Estructura de la tabla " & strTableName & ""
                    Print #x, "#"
                    
                    If Not IncludeCreateDB Then
                        Print #x, "DROP TABLE IF EXISTS `" & strTableName & "`;"
                    End If
                    Do While Not .EOF
                        Print #x, .Fields.Item(1).Value & ";"
                        
                        .MoveNext
                    Loop
                    .Close
                    
                End With
            End If
                
            If IncludeData Then
                With rssAux
                    .Open "SELECT * FROM " & strTableName & "", cnn
                    Print #x, ""
                    Print #x, "#"
                    Print #x, "# Datos de la tabla " & strTableName & ""
                    Print #x, "#"
'                    Print #X, "LOCK TABLES `" & strTableName & "` write;"
    
                    If Not .EOF Then
                        Print #x, "INSERT INTO `" & strTableName & "` VALUES "
                        
                        Do While Not .EOF
                        
                            strCurLine = ""
                            For I = 0 To .Fields.Count - 1
                                If IsNull(.Fields.Item(I).Value) Then
                                    If strCurLine <> "" Then
                                        strCurLine = strCurLine & ", "
                                    End If
                                    strCurLine = strCurLine & "Null"
                                Else
                                    strBuffer = .Fields.Item(I).Value
                                    
                                    If .Fields.Item(I).Type = 131 Then
                                        strBuffer = Replace(Format(strBuffer, "0.00"), ",", ".")
                                    End If
                                    
                                    strBuffer = Replace(strBuffer, "\", "\\")
                                    strBuffer = Replace(strBuffer, "'", "\'")
                                    strBuffer = Replace(strBuffer, Chr(10), "")
                                    strBuffer = Replace(strBuffer, Chr(13), "\r\n")
                                    
                                    If strCurLine <> "" Then
                                        strCurLine = strCurLine & ", "
                                    End If
                                    strCurLine = strCurLine & "'" & strBuffer & "'"
                                End If
                            Next
                            .MoveNext
                            
                            strCurLine = "(" & strCurLine & ")"
                            If .EOF Then
                                Print #x, strCurLine & ";"
                            Else
                                Print #x, strCurLine & ","
                            End If
                        Loop
                        
                    End If
'                    Print #X, "UNLOCK TABLES;"
                    
                    .Close
                End With
                Print #x, "#--------------------------------------------"
            End If
            
            .MoveNext
        Loop
        
        Print #x, ""
        Print #x, "SET FOREIGN_KEY_CHECKS=1;"
        Print #x, ""
'        Print #X, "# Fin del Respaldo: " & Format(Now, "DD/MM/YYYY HH:MM:SS")
        
        .Close
    End With
    
    Close #x
End Sub




'########################################################################
'# PROCESO:     dbRestoreBackup
'# OBJETO:      Procedimiento para recuperar un volcado de una base de datos MySQL
'#
'# PRECONDICIONES:
'#              Debe existir una variable global llamada 'cnn' del tipo ADODB.Connection
'#                  la cual está apuntando a la base de datos que se quiere respaldar
'# PARAMETROS:
'# (obligatorio)strFileName => Nombre del archivo plano donde se reside el volcado.
'# RETORNA:     n/a
'#
'# AUTOR:       Williams Castillo - will@eduven.com
'# FECHA ULTIMA MODIFICACION: 24/01/2006
'#
'# POR HACER:   Está desarrollado para versiones anteriores a la 5.0
'#              Falta incluir los stored procedures y triggers para hacerla
'#                  100% compatible con MySQL 5.0 y posteriores.
'########################################################################
Public Sub dbRestoreBackup(ByVal strFileName As String)
Dim TotalBytes As Long, CurrentBytes As Long
Dim x As Integer, strCurLine As String, strAux As String
Dim blnPassLines As Boolean
Dim blnAnalizeIt As Boolean
    
    x = FreeFile
    
    On Error GoTo ErrorsDrv
'    Call dbBeginTX
    cnn.BeginTrans
    
    Open strFileName For Input As #x
    TotalBytes = LOF(x)
    
    blnPassLines = False
    Do While Not EOF(x)
        Line Input #x, strCurLine
        CurrentBytes = CurrentBytes + LenB(strCurLine)
        
'        #If IS_A_PLUGGIN = 0 Then
'            Call UpdateProgressBar(TotalBytes, CurrentBytes)
'            Call MyDoEvents
'        #End If
        
        blnAnalizeIt = True
        strCurLine = Trim(strCurLine)
        If Not blnPassLines Then
            If left(strCurLine, 1) = "#" Then
                blnAnalizeIt = False
            ElseIf left(strCurLine, 2) = "/*" Then
                blnAnalizeIt = False
                blnPassLines = True
            End If
        ElseIf right(Trim(strCurLine), 2) = "*/" Then
            blnPassLines = False
            blnAnalizeIt = False
        End If
         
        If blnAnalizeIt And strCurLine <> "" Then
        
            While mID(strCurLine, Len(strCurLine), 1) <> ";"
                strAux = strCurLine
                Line Input #x, strCurLine
                CurrentBytes = CurrentBytes + LenB(strCurLine)
                strCurLine = Trim(strCurLine)
                
'                #If IS_A_PLUGGIN = 0 Then
'                    Call UpdateProgressBar(TotalBytes, CurrentBytes)
'                    Call MyDoEvents
'                #End If
                
                strCurLine = strAux & strCurLine
            Wend
            
'            Call dbExecuteSQL(strCurLine)
            cnn.Execute strCurLine
            
        End If
        
'        #If IS_A_PLUGGIN = 0 Then
'            Call MyDoEvents
'        #End If
    Loop
    
    Close #x
'    #If IS_A_PLUGGIN = 0 Then
'        Call UpdateProgressBar(TotalBytes, TotalBytes)
'    #End If
    
'    Call dbCommitTX
    cnn.CommitTrans
    
    Exit Sub
    
ErrorsDrv:
'    Call dbRollbackTX
    cnn.RollbackTrans
    
    MsgBox "ERROR:" & Err.Number & vbNewLine & Err.Description & vbNewLine, vbCritical
    Err.Clear
    
    Exit Sub
    Resume  'Debugging purpouses...
End Sub