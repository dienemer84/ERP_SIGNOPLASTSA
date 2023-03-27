Attribute VB_Name = "QRHelper"
Enum TQRCodeEncoding
    ceALPHA
    ceBYTE
    ceNUMERIC
    ceKANJI
    ceAUTO
End Enum

Enum TQRCodeECLevel
    LEVEL_L
    LEVEL_M
    LEVEL_Q
    LEVEL_H
End Enum

Private Declare Sub FullQRCode Lib "QRCodeLib.dll" _
                               (ByVal autoConfigurate As Boolean, _
                                ByVal AutoFit As Boolean, _
                                ByVal backColor As Long, _
                                ByVal barColor As Long, _
                                ByVal texto As String, _
                                ByVal correctionLevel As TQRCodeECLevel, _
                                ByVal encoding As TQRCodeEncoding, _
                                ByVal marginpixels As Integer, _
                                ByVal moduleWidth As Integer, _
                                ByVal Height As Integer, _
                                ByVal Width As Integer, _
                                ByVal filename As String)
Private Declare Sub FastQRCode Lib "QRCodeLib.dll" _
                               (ByVal texto As String, _
                                ByVal filename As String)
Private Declare Function QRCodeLibVer Lib "QRCodeLib.dll" () As String

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" _
                                    (ByVal lpFileName As String) _
                                    As Long


Public Sub generar(F As Factura)
    Dim url As String
    url = "https://www.afip.gob.ar/fe/qr/?p="
    Dim file As String
    file = App.path & "\" & F.Id & ".bmp"
    Dim oFile As New Scripting.FileSystemObject
    FileExists = oFile.FileExists(file)
    If FileExists Then
        DeleteFile file
    End If

    Dim T As String
    T = "{'ver':1,'fecha':'%fecha','cuit':30657604972,'ptoVta':%pv,'tipoCmp':%tipo,'nroCmp':%nroComp,'importe':%importe,'moneda':%moneda,'ctz':%tc,'tipoDocRec':%tipoDocReceptor,'nroDocRec':%DocReceptor,'tipoCodAut':'E','codAut':%cae}"
    T = Replace$(T, "%fecha", F.FechaEmision)
    T = Replace$(T, "%cae", F.CAE)
    T = Replace$(T, "%pv", F.Tipo.PuntoVenta.PuntoVenta)
    T = Replace$(T, "%nroComp", F.numero)

    Dim tipoafip As String


    If F.esCredito Then
        If F.TipoDocumento = tipoDocumentoContable.Factura Then

            If F.Tipo.TipoFactura.Id = 1 Then    'a
                tipoafip = "201"
            End If
            If F.Tipo.TipoFactura.Id = 2 Then    'b
                tipoafip = "206"
            End If

        End If


        If F.TipoDocumento = tipoDocumentoContable.notaCredito Then
            If F.Tipo.TipoFactura.Id = 1 Then    'a
                tipoafip = "203"
            End If
            If F.Tipo.TipoFactura.Id = 2 Then    'b
                tipoafip = "208"
            End If
        End If

        If F.TipoDocumento = tipoDocumentoContable.notaDebito Then
            If F.Tipo.TipoFactura.Id = 1 Then    'a
                tipoafip = "202"
            End If
            If F.Tipo.TipoFactura.Id = 2 Then    'b
                tipoafip = "207"
            End If
        End If

    Else
        'NO ES CREDITO
        If F.TipoDocumento = tipoDocumentoContable.Factura Then

            If F.Tipo.TipoFactura.Id = 1 Then    'a
                tipoafip = "001"
            End If
            If F.Tipo.TipoFactura.Id = 2 Then    'b
                tipoafip = "006"
            End If

        End If


        If F.TipoDocumento = tipoDocumentoContable.notaCredito Then
            If F.Tipo.TipoFactura.Id = 1 Then    'a
                tipoafip = "003"
            End If
            If F.Tipo.TipoFactura.Id = 2 Then    'b
                tipoafip = "008"
            End If
        End If

        If F.TipoDocumento = tipoDocumentoContable.notaDebito Then
            If F.Tipo.TipoFactura.Id = 1 Then    'a
                tipoafip = "002"
            End If
            If F.Tipo.TipoFactura.Id = 2 Then    'b
                tipoafip = "007"
            End If
        End If

    End If


    T = Replace$(T, "%tipo", tipoafip)    'va un mapeo de signo a afip

    If F.moneda.Id = 0 Then
        T = Replace$(T, "%moneda", "PES")    'va un mapeo de signo a afip
    End If
    If F.moneda.Id = 1 Then
        T = Replace$(T, "%moneda", "DOL")    'va un mapeo de signo a afip
    End If
    If F.moneda.Id = 2 Then
        T = Replace$(T, "%moneda", "060")    'va un mapeo de signo a afip
    End If
    T = Replace$(T, "%importe", F.TotalEstatico.Total)
    T = Replace$(T, "%tc", F.CambioAPatron)
    T = Replace$(T, "%tipoDocReceptor", 80)    'revisar si va siempre 80
    T = Replace$(T, "%DocReceptor", F.cliente.Cuit)
    Dim t64 As String

    t64 = Base64EncodeString(T)
    FastQRCode url & t64, file
End Sub

