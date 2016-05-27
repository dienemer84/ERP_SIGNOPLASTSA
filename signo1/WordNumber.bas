Attribute VB_Name = "WordNumber"
'---------------------------------------------------------------------------------------
' Module    : NumWord - WordNum
' DateTime  : 2006-04-12
' Author    : Rafael Arriagada (raff@pow4ever.com)
' Purpose   : Conversor de números a palabras y viceversa
' Comments  : Soporta números gigantes, 2 decimales y separador de miles (configuración regional).
'             Al interior de WordNumber.bas se incluye función SUM para sumar números en
'             formato String, de este modo se elimina el famoso problema de la notación
'             científica que utiliza VB al manipular números demasiado grandes, tal vez les sea útil.
'             Se incluye también la función IsNumber, la cual permite validar si un String
'             de largo ilimitado es número o no, ya que la función IsNumeric de VB no soporta
'             números gigantes.
'---------------------------------------------------------------------------------------
Option Explicit
Const sN1 = "1", sN2 = "2", sN3 = "3", sN4 = "4", sN5 = "5", sN6 = "6", sN7 = "7", sN8 = "8", sN9 = "9", sN0 = "0"

'---------------------------------------------------------------------------------------
' Procedure : NumWord
' DateTime  : 4/12/2006 18:29
' Author    : Rafael Arriagada (raff@pow4ever.com)
' Purpose   : Convierte cualquier número dentro de una frase a palabras, conservando el contexto
'---------------------------------------------------------------------------------------
'
Public Function NumWord(Data As String, Optional sSing As String = "peso", Optional sPlur As String = "pesos", Optional sSing2 As String = "centavo", Optional sPlur2 As String = "centavos", Optional sSing3 As String = "centécimo", Optional sPlur3 As String = "centécimos") As String
       Dim aC() As String, X As Long, NE As String, ND As String, CD As String * 1, CM As String * 1, NEG As Boolean
       Const sMenos As String = "menos ", sCon As String = " con ", sCR As String = "1,0", sComa As String = ",", sPunto As String = ".", sSMenos As String = "-"
       'Determina la configuración regional de este equipo
       If sCR + 1 = 2 Then
          CD = sComa: CM = sPunto
       Else
          CD = sPunto: CM = sComa
       End If
       aC = Split(Replace(Data, vbCrLf, Space(1)), Space(1))
       'Procesa parte por parte la frase
       For X = 0 To UBound(aC)
           'Si la parte es un número, la convierte en palabras
           If IsNumber(aC(X), CD, CM) Then
              'Elimina los separadores de miles si los hay (no recomendado usarlos)
              If InStr(1, aC(X), CM) Then aC(X) = Replace(aC(X), CM, vbNullString)
              If InStr(1, aC(X), CD) Then
                 NE = Left(aC(X), InStr(1, aC(X), CD, vbBinaryCompare) - 1)
                 ND = Left(Mid(aC(X), InStr(1, aC(X), CD, vbBinaryCompare) + 1), 2)
              Else
                 NE = aC(X)
              End If
              NEG = Left(NE, 1) = sMenos
              If NEG Then NE = Mid(NE, 2)
              If Len(ND) > 0 Then
                 aC(X) = NumWordInternal(NE) & Space(1) & IIf(Val(NE) = 1, sSing, sPlur)
                 If Left(ND, 1) = sN0 Then
                    aC(X) = aC(X) & sCon & NumWordInternal(ND * 10) & Space(1) & IIf(Val(ND) * 10 = 1, sSing3, sPlur3)
                 Else
                    aC(X) = aC(X) & sCon & NumWordInternal(ND) & Space(1) & IIf(Val(ND) = 1, sSing2, sPlur2)
                 End If
              Else
                 If sSing = vbNullString Then
                    aC(X) = NumWordInternal(NE)
                 Else
                    aC(X) = NumWordInternal(NE) & Space(1) & IIf(Trim(NE) = "1", sSing, sPlur)
                 End If
              End If
           If NEG Then aC(X) = sSMenos & aC(X)
         End If
       Next
       NumWord = Join(aC, Space(1))
End Function
'---------------------------------------------------------------------------------------
' Procedure : NumWordInternal
' DateTime  : 4/12/2006 18:29
' Author    : Rafael Arriagada (raff@pow4ever.com)
' Purpose   : Convierte números en palabras, uso interno ya que aquí no se manejan
'             problemáticas de alto nivel (formato, decimales, signo)
'---------------------------------------------------------------------------------------
'
Private Function NumWordInternal(ByVal N As String) As String
        Dim NB() As String, X As Long, C As Long, Y As Long, aC As String, CN1 As String, CN2 As String
        Const sMil As String = "mil ", sCero As String = "cero", sUN As String = "un", sDesconosido As String = "[...]", sComa As String = ","
        Const sNNames As String = "millón,millones,billón,billones,trillón,trillones,cuatrillón,cuatrillones,quintillón,quintillones,sextillón,sextillones,septillón,septillones,octillón,octillones,nonillón,nonillones,decillón,decillones,undecillón,undecillones,duodecillón,duodecillones,tredecillón,tredecillones,cuatordecillón,cuatordecillones,quindecillón,quindecillones,sexdecillón,sexdecillones,septendecillón,septendecillones,octodecillón,octodecillones,novendecillón,novendecillones,vigintillón,vigintillones"
        NB = Split(sNNames, sComa)
        C = Len(N)
        If C Mod 6 > 0 Then N = String(6 - C Mod 6, sN0) & N
        C = Len(N) / 3 - 1
        Y = C - 2
        For X = 0 To C
            If X Mod 2 > 0 Then
               CN2 = Mid(N, X * 3 + 1, 3)
               If Val(CN1) > 0 Then
                  If Y > 0 Then
                     If UBound(NB) < Y Then
                        If aC = sUN Then
                           aC = sMil & IIf(Val(CN2) > 0, Centena(CN2) & Space(1), vbNullString) & sDesconosido
                        Else
                           aC = aC & Space(1) & sMil & IIf(Val(CN2) > 0, Centena(CN2) & Space(1), vbNullString) & sDesconosido
                        End If
                     Else
                        If aC = sUN Then
                           aC = sMil & IIf(Val(CN2) > 0, Centena(CN2) & Space(1), vbNullString) & IIf(CN2 = sN1, NB(Y - 1), NB(Y))
                        Else
                           aC = aC & Space(1) & sMil & IIf(Val(CN2) > 0, Centena(CN2) & Space(1), vbNullString) & IIf(CN2 = sN1, NB(Y - 1), NB(Y))
                        End If
                     End If
                  Else
                     If aC = sUN Then
                        aC = sMil & Centena(CN2)
                     Else
                        aC = aC & Space(1) & Trim(sMil) & IIf(Centena(CN2) = vbNullString, vbNullString, Space(1) & Centena(CN2))
                     End If
                  End If
               ElseIf Val(CN1) + Val(CN2) > 0 Then
                  If Y > 0 Then
                     If UBound(NB) < Y Then
                        aC = aC & IIf(Len(aC) > 0 And Val(CN2) > 0, Space(1), vbNullString) & Centena(CN2) & Space(1) & sDesconosido
                     Else
                        aC = aC & IIf(Len(aC) > 0 And Val(CN2) > 0, Space(1), vbNullString) & Centena(CN2) & Space(1) & IIf(Val(CN2) = 1, NB(Y - 1), NB(Y))
                     End If
                  Else
                     aC = aC & IIf(Len(aC) > 0 And Val(CN2) > 0, Space(1), vbNullString) & Centena(CN2)
                  End If
               End If
               Y = Y - 2
            Else
               CN1 = Mid(N, X * 3 + 1, 3)
               If Val(CN1) > 0 Then aC = aC & IIf(Len(aC) > 0 And Val(CN1) > 0, Space(1), vbNullString) & Centena(CN1)
            End If
        Next
        NumWordInternal = IIf(aC = vbNullString, sCero, aC)
End Function
'---------------------------------------------------------------------------------------
' Procedure : Centena
' DateTime  : 4/12/2006 18:30
' Author    : Rafael Arriagada (raff@pow4ever.com)
' Purpose   : Convierte números de 3 dígitos en palabras
'---------------------------------------------------------------------------------------
'
Public Function Centena(N As String) As String
        Const sCE As String = ",cien,doscientos,trecientos,cuatrocientos,quinientos,seiscientos,setecientos,ochocientos,novecientos"
        Const sDE As String = ",dieci,veinti,treinta,cuarenta,cincuenta,sesenta,setenta,ochenta,noventa"
        Const sUN As String = ",un,dos,tres,cuatro,cinco,seis,siete,ocho,nueve", sComa As String = ",", sTo As String = "to "
        Const sES As String = "diez,once,doce,trece,catorce,quince,dieciséis,dós,trés,veinte", sY As String = " y "
        Dim D1 As String, D2 As String, D3 As String, B1 As Boolean, B2 As Boolean
        Dim aCE() As String, aDE() As String, aUN() As String, aES() As String
        'Asigna los dígitos de la Centena y recuerda si son mayores que cero
        D3 = Left(N, 1)
        D2 = Mid(N, 2, 1): B2 = Val(D2) > 0
        D1 = Right(N, 1): B1 = Val(D1) > 0
        aCE = Split(sCE, sComa)
        aDE = Split(sDE, sComa)
        aUN = Split(sUN, sComa)
        aES = Split(sES, sComa)
        
        'Procesa las unidades
        Centena = aUN(D1)
        
        'Procesa las decenas
        Select Case D2
               Case sN1
                    'Maneja lógica del [%&$&%$] que puso nombres ilógicos a algunos números.
                    Select Case D1
                           Case sN0: Centena = aES(0)
                           Case sN1: Centena = aES(1) 'dieciuno
                           Case sN2: Centena = aES(2) 'diecidos
                           Case sN3: Centena = aES(3) 'diecitres
                           Case sN4: Centena = aES(4) 'diecicuatro
                           Case sN5: Centena = aES(5) 'diecicinco
                           Case sN6: Centena = aES(6)
                           Case Else
                                Centena = aDE(D2) & Centena
                    End Select
               Case sN2
                    If B1 Then
                       If D1 = sN2 Then Centena = aES(7)
                       If D1 = sN3 Then Centena = aES(8)
                       Centena = aDE(D2) & Centena
                    Else
                       Centena = aES(9)
                    End If
               Case Is <> sN0
                    Centena = aDE(D2) & IIf(B1, sY & Centena, vbNullString)
        End Select
        'Procesa las Centenas
        If D3 = sN1 Then
           Centena = aCE(1) & IIf(B1 Or B2, sTo & Centena, vbNullString)
        ElseIf D3 <> sN0 Then
           Centena = aCE(D3) & IIf(B1 Or B2, Space(1) & Centena, vbNullString)
        End If
End Function
'---------------------------------------------------------------------------------------
' Procedure : WordNum
' DateTime  : 4/12/2006 18:30
' Author    : Rafael Arriagada (raff@pow4ever.com)
' Purpose   : Convierte un número escrito en un número real, ej: veintiuno -> 21
'---------------------------------------------------------------------------------------
'
Public Function WordNum(ByVal sNumero As String) As String
       Dim X As Long, aC() As String, S1 As String, S2 As String, S3 As String, LT As Boolean, RI As Boolean, LRI As Boolean
       'Limpieza de doble espacios si los hay
       Do While InStr(1, sNumero, Space(2)) > 0
          sNumero = Replace(sNumero, Space(2), Space(1))
       Loop
       
       'Corrección Cultural :)
       If InStr(1, sNumero, "c", vbTextCompare) Then sNumero = Replace(sNumero, "c", vbNullString)
       If InStr(1, sNumero, "z", vbTextCompare) Then sNumero = Replace(sNumero, "z", vbNullString)
       If InStr(1, sNumero, "s", vbTextCompare) Then sNumero = Replace(sNumero, "s", vbNullString)
       If InStr(1, sNumero, "b", vbTextCompare) Then sNumero = Replace(sNumero, "b", "v")
       If InStr(1, sNumero, "ó", vbTextCompare) Then sNumero = Replace(sNumero, "ó", vbNullString)
       If InStr(1, sNumero, "o", vbTextCompare) Then sNumero = Replace(sNumero, "o", vbNullString)
       If InStr(1, sNumero, "é", vbTextCompare) Then sNumero = Replace(sNumero, "é", "e")
       If InStr(1, sNumero, "ll", vbTextCompare) Then sNumero = Replace(sNumero, "ll", "y")
       If InStr(1, sNumero, "qu", vbTextCompare) Then sNumero = Replace(sNumero, "qu", "q")
       If InStr(1, sNumero, "k", vbTextCompare) Then sNumero = Replace(sNumero, "k", vbNullString)
       
       aC = Split(LCase(sNumero), Space(1))
       For X = 0 To UBound(aC)
           If Left(aC(X), 4) = "diei" Then
              S3 = sN1 & WordNum(Mid(aC(X), 5))
           ElseIf Left(aC(X), 6) = "veinti" Then
              S3 = sN2 & WordNum(Mid(aC(X), 7))
           Else
              Select Case aC(X)
                     Case "er", vbNullString: S3 = sN0
                     Case "un": S3 = sN1
                     Case "d": S3 = sN2
                     Case "tre": S3 = sN3
                     Case "uatr": S3 = sN4
                     Case "in": S3 = sN5
                     Case "ei": S3 = sN6
                     Case "iete": S3 = sN7
                     Case "h": S3 = sN8
                     Case "nueve": S3 = sN9
                     Case "die": S3 = sN1 & sN0
                     Case "ne": S3 = sN1 & sN1
                     Case "de": S3 = sN1 & sN2
                     Case "tree": S3 = sN1 & sN3
                     Case "atre": S3 = sN1 & sN4
                     Case "qine": S3 = sN1 & sN5
                     Case "veinte": S3 = sN2 & sN0
                     Case "treinta": S3 = sN3 & sN0
                     Case "uarenta": S3 = sN4 & sN0
                     Case "inuenta": S3 = sN5 & sN0
                     Case "eenta": S3 = sN6 & sN0
                     Case "etenta": S3 = sN7 & sN0
                     Case "henta": S3 = sN8 & sN0
                     Case "nventa": S3 = sN9 & sN0
                     Case "ien", "ient": S3 = sN1 & sN0 & sN0
                     Case "dient": S3 = sN2 & sN0 & sN0
                     Case "treient": S3 = sN3 & sN0 & sN0
                     Case "uatrient": S3 = sN4 & sN0 & sN0
                     Case "qinient": S3 = sN5 & sN0 & sN0
                     Case "eiient": S3 = sN6 & sN0 & sN0
                     Case "eteient": S3 = sN7 & sN0 & sN0
                     Case "hient": S3 = sN8 & sN0 & sN0
                     Case "nveient": S3 = sN9 & sN0 & sN0
                     Case "mil": S3 = sN1 & sN0 & sN0 & sN0
                     Case "miyn", "miyne": RI = True: S3 = sN1 & String(6, sN0)
                     Case "viyn", "viyne": RI = True: S3 = sN1 & String(12, sN0)
                     Case "triyn", "triyne": RI = True: S3 = sN1 & String(18, sN0)
                     Case "uatriyn", "uatriyne": RI = True: S3 = sN1 & String(24, sN0)
                     Case "qintiyn", "qintiyne": RI = True: S3 = sN1 & String(30, sN0)
                     Case "extiyn", "extiyne": RI = True: S3 = sN1 & String(36, sN0)
                     Case "eptiyn", "eptiyne": RI = True: S3 = sN1 & String(42, sN0)
                     Case "tiyn", "tiyne": RI = True: S3 = sN1 & String(48, sN0)
                     Case "nniyn", "nniyne": RI = True: S3 = sN1 & String(54, sN0)
                     Case "deiyn", "deiyne": RI = True: S3 = sN1 & String(60, sN0)
                     Case "undeiyn", "undeiyne": RI = True: S3 = sN1 & String(66, sN0)
                     Case "dudeiyn", "dudeiyne": RI = True: S3 = sN1 & String(72, sN0)
                     Case "tredeiyn", "tredeiyne": RI = True: S3 = sN1 & String(78, sN0)
                     Case "uatrdeiyn", "uatrdeiyne": RI = True: S3 = sN1 & String(84, sN0)
                     Case "qindeiyn", "qindeiyne": RI = True: S3 = sN1 & String(90, sN0)
                     Case "exdeiyn", "exdeiyne": RI = True: S3 = sN1 & String(96, sN0)
                     Case "eptendeiyn", "eptendeiyne": RI = True: S3 = sN1 & String(102, sN0)
                     Case "tdeiyn", "tdeiyne": RI = True: S3 = sN1 & String(108, sN0)
                     Case "nvendeiyn", "nvendeiyne": RI = True: S3 = sN1 & String(114, sN0)
                     Case "vigintiyn", "vigintiyne": RI = True: S3 = sN1 & String(120, sN0)
                     Case Else: LT = True
              End Select
           End If
           If Not LT Then
              If Not RI Then
                 If Len(S2) > Len(S3) Then
                    S3 = Mid(S2, 1, Len(S2) - Len(S3)) & S3
                 ElseIf Len(S2) > 0 Then
                    S3 = S2 & Mid(S3, 2)
                 End If
                 S2 = S3
              Else
                 If Len(S1) > 0 Then
                    S2 = S2 & Mid(S3, 2)
                    S1 = Left(S1, Len(S1) - Len(S2)) & S2
                 Else
                    S1 = S2
                    If Len(S2) > 0 Then S1 = S2 & Mid(S3, 2)
                 End If
                 S2 = vbNullString
                 RI = False
              End If
           Else
              LT = False
              LRI = True
           End If
       Next
       If Len(S2) > 0 Then
          If Len(S1) > 0 Then
             S1 = Mid(S1, 1, Len(S1) - Len(S2)) & S2
          Else
             S1 = S2
          End If
       End If
       WordNum = S1
End Function
'---------------------------------------------------------------------------------------
' Procedure : IsNumber
' DateTime  : 4/12/2006 18:31
' Author    : Rafael Arriagada (raff@pow4ever.com)
' Purpose   : Retorna True cuando el dato indicado es 100% numérico
'---------------------------------------------------------------------------------------
'
Public Function IsNumber(ByVal N As String, Optional ByVal CD As String = ".", Optional ByVal CM As String = ",") As Boolean
       Dim sDec As String
       N = Replace(N, CM, vbNullString)
       If Len(N) - Len(Replace(N, CD, vbNullString)) > 1 Then Exit Function
       If InStr(1, N, CD, vbBinaryCompare) > 0 Then
          sDec = Mid(N, InStr(1, N, CD, vbBinaryCompare) + 1)
          N = Left(N, InStr(1, N, CD, vbBinaryCompare) - 1)
          IsNumber = IsNumber(N, CD, CM) And IsNumber(sDec, CD, CM)
          Exit Function
       End If
       IsNumber = True
       Do While IsNumber
          If Len(N) > 65 Then
             IsNumber = IsNumeric(Left(N, 64)) And IsNumber
             N = Mid(N, 65)
          Else
             IsNumber = IsNumeric(N) And IsNumber
             Exit Function
          End If
       Loop
End Function
'---------------------------------------------------------------------------------------
' Procedure : Sum
' DateTime  : 4/12/2006 18:32
' Author    : Rafael Arriagada (raff@pow4ever.com)
' Purpose   : Realiza una suma entre variables String - Evita notación científica de VB
'---------------------------------------------------------------------------------------
'
Function Sum(ByVal N As String, ByVal AD As String) As String
         Dim R As String, X As Long, Y As Long, Z As Long, S As Boolean
         X = Len(N)
         Z = Len(AD)
         If X < Z Then R = N: N = AD: AD = R: X = Z: Z = Len(AD)
         Do While True
            R = Mid(N, X, 1)
            Y = Val(R) + Val(Mid(AD, Z, 1)) + Abs(S)
            If Y > 9 Then
               N = Left(N, X - 1) & Abs(10 - Y) & Mid(N, X + 1)
               X = X - 1: Z = Z - 1
               If Z < 1 Then Z = 1: AD = sN0
               S = True
            Else
               N = Left(N, X - 1) & CStr(Y) & Mid(N, X + 1)
               X = X - 1: Z = Z - 1: S = False
               If Z < 1 Then Exit Do
            End If
            If X < 1 Then
               N = sN1 & N
               Exit Do
            End If
         Loop
         Sum = N
End Function
