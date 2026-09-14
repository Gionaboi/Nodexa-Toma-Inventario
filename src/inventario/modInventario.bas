Option Explicit

' ===== Columnas de tblAurora (orden fijado en la consulta Power Query) =====
Public Const A_SERIAL As Long = 1
Public Const A_CODIGO As Long = 2
Public Const A_DESC As Long = 3
Public Const A_COND As Long = 4
Public Const A_ESTADO As Long = 5
Public Const A_BODEGA As Long = 6
Public Const A_SUC As Long = 7
Public Const A_NV As Long = 8
Public Const A_PICK As Long = 9

' ===== Resultado de cada escaneo =====
Public Const E_ENCONTRADO As String = "ENCONTRADO"
Public Const E_DUPLICADO As String = "DUPLICADO"
Public Const E_OTRA As String = "OTRA SUCURSAL"
Public Const E_NOEXISTE As String = "NO EXISTE"
Public Const E_SALIDA As String = "EN SALIDA"
Public Const E_EXCLUIDO As String = "BODEGA EXCLUIDA"
Public Const E_SKU As String = "CÓDIGO SKU"
' ===== Estados del resultado final =====
Public Const E_FALTANTE As String = "FALTANTE"
Public Const E_AJUSTADO As String = "AJUSTADO"
Public Const E_SOB_OTRA As String = "SOBRANTE - OTRA SUCURSAL"
Public Const E_SOB_NOEXISTE As String = "SOBRANTE - NO EXISTE"

' ===== Hoja ESCANEOS =====
Public Const L_N As Long = 1
Public Const L_FECHA As Long = 2
Public Const L_SERIAL As Long = 3
Public Const L_RESULT As Long = 4
Public Const L_DETALLE As Long = 5
Public Const L_ORIGEN As Long = 6
Public Const L_ANULADO As Long = 7

' ===== Celdas de CONFIG =====
Public Const CFG_SUCURSAL As String = "C3"
Public Const CFG_INICIO As String = "C5"
Public Const CFG_AURORA As String = "C6"
Public Const CFG_AURORA_FECHA As String = "C7"
Public Const CFG_ACTUALIZADO As String = "C8"
Public Const CFG_FILA_BODEGAS As Long = 12

' ===== Hoja RESULTADO =====
Public Const RES_FILA As Long = 8
Public Const RES_COLS As Long = 10

' ===== Índices en memoria (se reconstruyen con CargarIndices) =====
Public arrA As Variant           ' datos Aurora
Public dAurora As Object         ' serial -> fila en arrA (todas las sucursales)
Public dSKU As Object            ' código SKU -> True
Public dEsperado As Object       ' serial -> fila en arrA (lo que se debe encontrar)
Public dEscaneado As Object      ' serial -> fila en ESCANEOS (escaneos válidos)
Public nEncontrados As Long, nSobrantes As Long, nDuplicados As Long, nSalida As Long
' Totales del último GenerarResultado
Public rEsp As Long, rEnc As Long, rFal As Long, rSob As Long, rAj As Long, rSal As Long
Private sSucursal As String

Public Function Cfg(celda As String) As Range
    Set Cfg = Hj("CONFIG").Range(celda)
End Function

Public Function Sucursal() As String
    Sucursal = Trim(CStr(Cfg(CFG_SUCURSAL).Value))
End Function

' ---------------------------------------------------------------------------
' Índices
' ---------------------------------------------------------------------------

Public Function CargarIndices() As Boolean
    Set dAurora = NuevoDic(): Set dSKU = NuevoDic(): Set dEsperado = NuevoDic(): Set dEscaneado = NuevoDic()
    nEncontrados = 0: nSobrantes = 0: nDuplicados = 0: nSalida = 0
    sSucursal = Sucursal()

    Dim lo As ListObject: Set lo = Hj("Datos").ListObjects(1)
    If lo.ListRows.Count = 0 Then Exit Function
    arrA = lo.DataBodyRange.Value

    Dim dInc As Object: Set dInc = BodegasIncluidas()
    Dim i As Long, s As String
    For i = 1 To UBound(arrA, 1)
        s = CStr(arrA(i, A_SERIAL))
        If Len(s) > 0 Then
            If Not dAurora.Exists(s) Then dAurora.Add s, i
            If Len(arrA(i, A_CODIGO)) > 0 Then dSKU(UCase$(CStr(arrA(i, A_CODIGO)))) = True
            If EsEsperado(i, dInc) Then dEsperado(s) = i
        End If
    Next
    If dAurora.Count = 0 Then Exit Function

    Dim ws As Worksheet: Set ws = Hj("ESCANEOS")
    Dim ult As Long: ult = UltimaFila(ws)
    If ult >= 2 Then
        Dim L As Variant: L = ws.Range(ws.Cells(1, 1), ws.Cells(ult, L_ANULADO)).Value
        For i = 2 To ult
            If CStr(L(i, L_ANULADO)) <> "SÍ" Then Contar CStr(L(i, L_SERIAL)), CStr(L(i, L_RESULT)), i, 1
        Next
    End If
    CargarIndices = True
End Function

Private Function EsEsperado(i As Long, dInc As Object) As Boolean
    If StrComp(CStr(arrA(i, A_SUC)), sSucursal, vbTextCompare) <> 0 Then Exit Function
    If TieneSalida(i) Then Exit Function
    Dim b As String: b = CStr(arrA(i, A_BODEGA))
    If dInc.Exists(b) Then
        If Not dInc(b) Then Exit Function
    End If
    EsEsperado = True
End Function

' Con Nota de Venta o Picking asignado: va de salida, no se cuenta
Public Function TieneSalida(i As Long) As Boolean
    TieneSalida = Not (EsNo(arrA(i, A_NV)) And EsNo(arrA(i, A_PICK)))
End Function

Private Function EsNo(v As Variant) As Boolean
    Dim s As String: s = UCase$(Trim$(CStr(v)))
    EsNo = (s = "" Or s = "NO")
End Function

Private Function BodegasIncluidas() As Object
    Dim d As Object: Set d = NuevoDic(True)
    Dim ws As Worksheet: Set ws = Hj("CONFIG")
    Dim f As Long, b As String
    For f = CFG_FILA_BODEGAS To UltimaFila(ws, 2)
        b = Trim$(CStr(ws.Cells(f, 2).Value))
        If Len(b) > 0 Then d(b) = (UCase$(Left$(Trim$(CStr(ws.Cells(f, 3).Value)), 1)) <> "N")
    Next
    Set BodegasIncluidas = d
End Function

' Agrega a CONFIG las bodegas de la sucursal que aún no están (con "Sí")
Public Function RegistrarBodegas() As String
    Dim ws As Worksheet: Set ws = Hj("CONFIG")
    Dim dExiste As Object: Set dExiste = NuevoDic(True)
    Dim ult As Long: ult = UltimaFila(ws, 2)
    Dim f As Long
    For f = CFG_FILA_BODEGAS To ult
        If Len(ws.Cells(f, 2).Value) > 0 Then dExiste(CStr(ws.Cells(f, 2).Value)) = True
    Next
    If ult < CFG_FILA_BODEGAS - 1 Then ult = CFG_FILA_BODEGAS - 1

    Dim i As Long, b As String, nuevas As String
    For i = 1 To UBound(arrA, 1)
        If StrComp(CStr(arrA(i, A_SUC)), sSucursal, vbTextCompare) = 0 Then
            b = CStr(arrA(i, A_BODEGA))
            If Len(b) > 0 And Not dExiste.Exists(b) Then
                dExiste(b) = True
                ult = ult + 1
                ws.Cells(ult, 2).Value = b
                ws.Cells(ult, 3).Value = "Sí"
                If nuevas <> "" Then nuevas = nuevas & ", "
                nuevas = nuevas & b
            End If
        End If
    Next
    RegistrarBodegas = nuevas
End Function

' Solo para la construcción del archivo
Public Sub InicializarConfig()
    If CargarIndices() Then RegistrarBodegas
End Sub

' Estado actual de un serial según Aurora (sin considerar duplicados)
Public Function Clasificar(serial As String) As String
    If dEsperado.Exists(serial) Then
        Clasificar = E_ENCONTRADO
    ElseIf Not dAurora.Exists(serial) Then
        If dSKU.Exists(serial) Then Clasificar = E_SKU Else Clasificar = E_NOEXISTE
    Else
        Dim i As Long: i = dAurora(serial)
        If StrComp(CStr(arrA(i, A_SUC)), sSucursal, vbTextCompare) <> 0 Then
            Clasificar = E_OTRA
        ElseIf TieneSalida(i) Then
            Clasificar = E_SALIDA
        Else
            Clasificar = E_EXCLUIDO
        End If
    End If
End Function

' Suma (signo = 1) o resta (signo = -1) un escaneo a los contadores
Private Sub Contar(serial As String, resultado As String, fila As Long, signo As Long)
    If resultado = E_DUPLICADO Then
        nDuplicados = nDuplicados + signo
        Exit Sub
    End If
    If resultado = E_SKU Then Exit Sub
    If signo > 0 Then
        dEscaneado(serial) = fila
    ElseIf dEscaneado.Exists(serial) Then
        dEscaneado.Remove serial
    End If
    Select Case Clasificar(serial)
        Case E_ENCONTRADO: nEncontrados = nEncontrados + signo
        Case E_OTRA, E_NOEXISTE, E_SKU: nSobrantes = nSobrantes + signo
        Case E_SALIDA: nSalida = nSalida + signo
    End Select
End Sub

' ---------------------------------------------------------------------------
' Escaneo
' ---------------------------------------------------------------------------

Public Function ProcesarEscaneo(ByVal codigo As String, origen As String, ByRef detalle As String) As String
    Dim s As String: s = NormalizarSerial(codigo)
    If s = "" Then Exit Function
    Dim r As String
    If dEscaneado.Exists(s) Then
        r = E_DUPLICADO
        detalle = "Ya fue escaneado el " & Format(Hj("ESCANEOS").Cells(dEscaneado(s), L_FECHA).Value, "dd-mm-yyyy hh:nn") & _
                  ". No se cuenta dos veces. " & Descripcion(s)
    Else
        r = Clasificar(s)
        detalle = DetalleSerial(s, r)
    End If
    If r = E_SKU Then
        detalle = "Es un código de producto (SKU), no un serial. No se registró: escanee el serial del equipo."
        ProcesarEscaneo = r
        Exit Function
    End If
    Dim fila As Long: fila = RegistrarLog(s, r, detalle, origen)
    Contar s, r, fila, 1
    ProcesarEscaneo = r
End Function

Private Function RegistrarLog(s As String, r As String, detalle As String, origen As String) As Long
    Dim ws As Worksheet: Set ws = Hj("ESCANEOS")
    Dim f As Long: f = UltimaFila(ws) + 1
    If IsEmpty(Cfg(CFG_INICIO).Value) Then Cfg(CFG_INICIO).Value = Now
    ws.Range(ws.Cells(f, 1), ws.Cells(f, L_ANULADO)).Value = Array(f - 1, Now, s, r, detalle, origen, "")
    RegistrarLog = f
End Function

' Anula (no borra) el último escaneo vigente. Devuelve el serial anulado.
Public Function DeshacerUltimo() As String
    Dim ws As Worksheet: Set ws = Hj("ESCANEOS")
    Dim f As Long
    For f = UltimaFila(ws) To 2 Step -1
        If CStr(ws.Cells(f, L_ANULADO).Value) <> "SÍ" Then Exit For
    Next
    If f < 2 Then Exit Function
    Dim s As String: s = CStr(ws.Cells(f, L_SERIAL).Value)
    ws.Cells(f, L_ANULADO).Value = "SÍ"
    Contar s, CStr(ws.Cells(f, L_RESULT).Value), f, -1
    DeshacerUltimo = s
End Function

Private Function Descripcion(s As String) As String
    If dAurora.Exists(s) Then Descripcion = CStr(arrA(dAurora(s), A_DESC))
End Function

Private Function DetalleSerial(s As String, r As String) As String
    If Not dAurora.Exists(s) Then
        DetalleSerial = "El serial no aparece en el reporte Aurora (ninguna sucursal)."
        Exit Function
    End If
    Dim i As Long: i = dAurora(s)
    Select Case r
        Case E_OTRA
            DetalleSerial = "Aurora lo registra en " & arrA(i, A_SUC) & " (" & arrA(i, A_BODEGA) & "). " & arrA(i, A_DESC)
        Case E_SALIDA
            DetalleSerial = "Tiene Nota de Venta " & arrA(i, A_NV) & " / Picking " & arrA(i, A_PICK) & ": va de salida, no se cuenta. " & arrA(i, A_DESC)
        Case E_EXCLUIDO
            DetalleSerial = "La bodega '" & arrA(i, A_BODEGA) & "' está marcada para NO contarse (hoja CONFIG). " & arrA(i, A_DESC)
        Case Else
            DetalleSerial = arrA(i, A_DESC) & "  |  " & arrA(i, A_BODEGA) & "  |  " & arrA(i, A_COND)
    End Select
End Function

' ---------------------------------------------------------------------------
' Botones
' ---------------------------------------------------------------------------

Public Sub Escanear()
    If Not CargarIndices() Then
        MsgBox "No hay datos de Aurora cargados." & vbLf & vbLf & "Presione primero ACTUALIZAR AURORA.", vbExclamation, "Escanear"
        Exit Sub
    End If
    frmEscaneo.Show
    ActualizarInicio
    GuardarSilencioso
End Sub

Public Sub ActualizarAurora()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim base As String: base = RutaLocal()
    If base = "" Then
        MsgBox "No se pudo determinar la carpeta de este archivo." & vbLf & "Guárdelo en una carpeta del equipo (Escritorio o Documentos).", vbCritical
        Exit Sub
    End If
    Dim carpeta As String: carpeta = base & "\Aurora"
    If Not fso.FolderExists(carpeta) Then
        MsgBox "No existe la carpeta 'Aurora' junto a este archivo:" & vbLf & carpeta & vbLf & vbLf & _
               "Créela y guarde dentro el reporte descargado de Aurora.", vbExclamation
        Exit Sub
    End If

    ' Se usa el .xlsx más reciente de la carpeta, sin importar su nombre
    Dim f As Object, ultimo As Object
    For Each f In fso.GetFolder(carpeta).Files
        If LCase(fso.GetExtensionName(f.Name)) = "xlsx" And Left(f.Name, 2) <> "~$" Then
            If ultimo Is Nothing Then
                Set ultimo = f
            ElseIf f.DateLastModified > ultimo.DateLastModified Then
                Set ultimo = f
            End If
        End If
    Next
    If ultimo Is Nothing Then
        MsgBox "La carpeta Aurora no tiene ningún archivo .xlsx." & vbLf & vbLf & _
               "Descargue el reporte de productos almacenados desde Aurora y guárdelo en:" & vbLf & carpeta, vbExclamation
        Exit Sub
    End If

    Dim q As WorkbookQuery: Set q = ThisWorkbook.Queries("Aurora")
    q.Formula = CambiarRuta(q.Formula, ultimo.Path)
    Application.StatusBar = "Leyendo reporte Aurora (unos segundos)..."
    Application.Cursor = xlWait
    On Error GoTo Fallo
    Application.DisplayAlerts = False
    Hj("Datos").ListObjects(1).QueryTable.Refresh BackgroundQuery:=False
    Application.DisplayAlerts = True
    On Error GoTo 0
    Application.StatusBar = False
    Application.Cursor = xlDefault

    Cfg(CFG_AURORA).Value = ultimo.Name
    Cfg(CFG_AURORA_FECHA).Value = ultimo.DateLastModified
    Cfg(CFG_ACTUALIZADO).Value = Now
    If Not CargarIndices() Then
        MsgBox "El reporte se leyó, pero no trae seriales.", vbExclamation
        Exit Sub
    End If
    Dim nuevas As String: nuevas = RegistrarBodegas()
    ActualizarInicio
    GuardarSilencioso

    Dim msg As String
    msg = "Reporte cargado: " & ultimo.Name & vbLf & vbLf & _
          "Seriales en el reporte (todas las sucursales): " & Format(dAurora.Count, "#,##0") & vbLf & _
          "Seriales esperados en " & sSucursal & ": " & Format(dEsperado.Count, "#,##0")
    If dEsperado.Count = 0 Then msg = msg & vbLf & vbLf & "ATENCIÓN: no hay seriales de " & sSucursal & ". ¿Descargó el reporte correcto?"
    If nuevas <> "" Then msg = msg & vbLf & vbLf & "Bodegas nuevas agregadas en CONFIG (se cuentan): " & nuevas
    If Not IsEmpty(Cfg(CFG_INICIO).Value) Then msg = msg & vbLf & vbLf & "Los escaneos de la toma en curso se mantienen."
    MsgBox msg, vbInformation, "Aurora actualizado"
    Exit Sub

Fallo:
    Dim e As String: e = Err.Description
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.Cursor = xlDefault
    MsgBox "No se pudo leer el reporte Aurora:" & vbLf & e & vbLf & vbLf & _
           "- Si el reporte está abierto en Excel, ciérrelo y repita." & vbLf & _
           "- Si el mensaje menciona columnas, el formato del reporte cambió: avise para ajustar el archivo.", _
           vbCritical, "Error al leer Aurora"
End Sub

Public Sub VerResultado()
    If GenerarResultado() Then
        Hj("RESULTADO").Activate
        Hj("RESULTADO").Range("A1").Select
    End If
End Sub

Public Sub IrInicio()
    Hj("INICIO").Activate
    Hj("INICIO").Range("A1").Select
End Sub

Public Sub AjusteManual()
    Dim s As String
    s = NormalizarSerial(InputBox("Escanee o escriba el serial que quiere ajustar:", "Ajuste manual"))
    If s = "" Then Exit Sub
    Dim m As String
    m = Trim(InputBox("Motivo del ajuste para " & s & ":" & vbLf & vbLf & _
        "Ejemplos: 'Salió a cliente durante la toma', 'Llegó hoy desde otra sucursal'.", "Ajuste manual"))
    If m = "" Then Exit Sub
    Dim ws As Worksheet: Set ws = Hj("AJUSTES")
    Dim f As Long: f = UltimaFila(ws) + 1
    ws.Range(ws.Cells(f, 1), ws.Cells(f, 3)).Value = Array(Now, s, m)
    GuardarSilencioso
    MsgBox "Ajuste registrado para " & s & "." & vbLf & vbLf & _
           "Si ese serial es faltante o sobrante, en el resultado aparecerá como AJUSTADO con su motivo." & vbLf & _
           "Para quitar un ajuste, borre su fila en la hoja AJUSTES.", vbInformation, "Ajuste manual"
End Sub

Public Sub BotonExportar()
    If Not GenerarResultado() Then Exit Sub
    Dim ruta As String: ruta = ExportarReporte()
    If ruta <> "" Then
        MsgBox "Reporte exportado:" & vbLf & ruta & vbLf & vbLf & _
               "Es un Excel normal (sin macros): puede abrirse y trabajarse en cualquier PC.", vbInformation, "Exportar"
    End If
End Sub

Public Sub FinalizarToma()
    If Not GenerarResultado() Then Exit Sub
    If dEscaneado.Count = 0 Then
        MsgBox "La toma actual no tiene escaneos. No hay nada que finalizar.", vbInformation, "Finalizar toma"
        Exit Sub
    End If
    Dim msg As String
    msg = "Resumen de la toma:" & vbLf & vbLf & _
          "   Esperados en sistema:  " & rEsp & vbLf & _
          "   Encontrados:  " & rEnc & vbLf & _
          "   Faltantes:  " & rFal & vbLf & _
          "   Sobrantes:  " & rSob & vbLf & _
          "   Ajustados:  " & rAj & vbLf & vbLf & _
          "Al finalizar:" & vbLf & _
          "   - se guarda en el historial (hoja HISTORICO)," & vbLf & _
          "   - se crea el reporte en la carpeta Reportes," & vbLf & _
          "   - se limpia la toma para empezar una nueva." & vbLf & vbLf & _
          "¿Finalizar la toma?"
    ' Botón por defecto = No: un Enter accidental de la pistola no cierra la toma
    If MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2, "Finalizar toma") <> vbYes Then Exit Sub

    Dim ruta As String: ruta = ExportarReporte()
    If ruta = "" Then
        MsgBox "No se pudo crear el reporte. La toma NO se finalizó.", vbCritical, "Finalizar toma"
        Exit Sub
    End If
    ArchivarHistorial ruta
    LimpiarToma
    CargarIndices
    ActualizarInicio
    GuardarSilencioso
    Hj("INICIO").Activate
    MsgBox "Toma finalizada y guardada en el historial." & vbLf & vbLf & "Reporte: " & ruta, vbInformation, "Finalizar toma"
End Sub

' ---------------------------------------------------------------------------
' Resultado, reporte e historial
' ---------------------------------------------------------------------------

Public Function GenerarResultado() As Boolean
    If Not CargarIndices() Then
        MsgBox "No hay datos de Aurora cargados." & vbLf & vbLf & "Presione ACTUALIZAR AURORA en la hoja INICIO.", vbExclamation
        Exit Function
    End If
    Dim dAj As Object: Set dAj = LeerAjustes()
    Dim wsL As Worksheet: Set wsL = Hj("ESCANEOS")
    Dim ultL As Long: ultL = UltimaFila(wsL)
    Dim L As Variant
    If ultL >= 2 Then L = wsL.Range(wsL.Cells(1, 1), wsL.Cells(ultL, L_ANULADO)).Value

    rEsp = dEsperado.Count: rEnc = 0: rFal = 0: rSob = 0: rAj = 0: rSal = 0
    Dim out() As Variant
    ReDim out(1 To dEsperado.Count + dEscaneado.Count + 1, 1 To RES_COLS + 1)
    Dim k As Long, s As Variant, est As String, fecha As Variant, aj As String, ia As Long

    For Each s In dEsperado.Keys
        fecha = Empty: aj = ""
        If dAj.Exists(s) Then aj = dAj(s)
        If dEscaneado.Exists(s) Then
            est = E_ENCONTRADO: fecha = L(dEscaneado(s), L_FECHA): rEnc = rEnc + 1
        ElseIf aj <> "" Then
            est = E_AJUSTADO: rAj = rAj + 1
        Else
            est = E_FALTANTE: rFal = rFal + 1
        End If
        k = k + 1
        LlenarFila out, k, est, CStr(s), CLng(dEsperado(s)), fecha, aj
    Next

    For Each s In dEscaneado.Keys
        If Not dEsperado.Exists(s) Then
            aj = ""
            If dAj.Exists(s) Then aj = dAj(s)
            est = Clasificar(CStr(s))
            Select Case est
                Case E_OTRA, E_NOEXISTE, E_SKU
                    If aj <> "" Then
                        est = E_AJUSTADO: rAj = rAj + 1
                    Else
                        If est = E_OTRA Then est = E_SOB_OTRA Else est = E_SOB_NOEXISTE
                        rSob = rSob + 1
                    End If
                Case E_SALIDA
                    rSal = rSal + 1
            End Select
            ia = 0
            If dAurora.Exists(s) Then ia = dAurora(s)
            k = k + 1
            LlenarFila out, k, est, CStr(s), ia, L(dEscaneado(s), L_FECHA), aj
        End If
    Next

    Dim ws As Worksheet: Set ws = Hj("RESULTADO")
    Application.ScreenUpdating = False
    LimpiarResultado
    If k > 0 Then
        With ws.Cells(RES_FILA, 1).Resize(k, RES_COLS + 1)
            .Value = out
            .Sort Key1:=ws.Cells(RES_FILA, RES_COLS + 1), Order1:=xlAscending, _
                  Key2:=ws.Cells(RES_FILA, 2), Order2:=xlAscending, Header:=xlNo
        End With
        ws.Cells(RES_FILA, RES_COLS + 1).Resize(k, 1).ClearContents
    End If
    ws.Cells(RES_FILA - 1, 1).Resize(k + 1, RES_COLS).AutoFilter

    Dim avance As Double
    If rEsp > 0 Then avance = rEnc / rEsp
    ws.Range("A5:G5").Value = Array(rEsp, rEnc, rFal, rSob, rAj, rSal, avance)
    ws.Range("A2").Value = "Actualizado: " & Format(Now, "dd-mm-yyyy hh:nn") & _
        "     |     Toma desde: " & TextoInicio() & _
        "     |     Aurora: " & Cfg(CFG_AURORA).Value
    Application.ScreenUpdating = True
    GenerarResultado = True
End Function

Private Sub LlenarFila(out() As Variant, k As Long, est As String, s As String, ia As Long, fecha As Variant, aj As String)
    out(k, 1) = est
    out(k, 2) = s
    If ia > 0 Then
        out(k, 3) = arrA(ia, A_CODIGO)
        out(k, 4) = arrA(ia, A_DESC)
        out(k, 5) = arrA(ia, A_BODEGA)
        out(k, 6) = arrA(ia, A_COND)
        out(k, 7) = arrA(ia, A_ESTADO)
        out(k, 8) = arrA(ia, A_SUC)
    End If
    out(k, 9) = fecha
    out(k, 10) = aj
    out(k, 11) = Prioridad(est)
End Sub

Private Function Prioridad(est As String) As Long
    Select Case est
        Case E_FALTANTE: Prioridad = 1
        Case E_SOB_NOEXISTE: Prioridad = 2
        Case E_SOB_OTRA: Prioridad = 3
        Case E_AJUSTADO: Prioridad = 4
        Case E_SALIDA: Prioridad = 5
        Case E_EXCLUIDO: Prioridad = 6
        Case Else: Prioridad = 7
    End Select
End Function

Private Function EsDiferencia(est As String) As Boolean
    EsDiferencia = (est = E_FALTANTE Or est = E_SOB_NOEXISTE Or est = E_SOB_OTRA Or est = E_AJUSTADO)
End Function

Private Sub LimpiarResultado()
    Dim ws As Worksheet: Set ws = Hj("RESULTADO")
    If ws.FilterMode Then ws.ShowAllData
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    Dim ult As Long: ult = UltimaFila(ws, 2)
    If ult >= RES_FILA Then ws.Range(ws.Cells(RES_FILA, 1), ws.Cells(ult, RES_COLS + 1)).ClearContents
End Sub

Private Function LeerAjustes() As Object
    Dim d As Object: Set d = NuevoDic()
    Dim ws As Worksheet: Set ws = Hj("AJUSTES")
    Dim f As Long, s As String
    For f = 2 To UltimaFila(ws, 2)
        s = NormalizarSerial(CStr(ws.Cells(f, 2).Value))
        If s <> "" Then d(s) = CStr(ws.Cells(f, 3).Value)
    Next
    Set LeerAjustes = d
End Function

Private Function TextoInicio() As String
    If IsEmpty(Cfg(CFG_INICIO).Value) Then
        TextoInicio = "sin escaneos"
    Else
        TextoInicio = Format(Cfg(CFG_INICIO).Value, "dd-mm-yyyy hh:nn")
    End If
End Function

' Crea el reporte .xlsx (sin macros) en la carpeta Reportes. Devuelve la ruta.
Public Function ExportarReporte() As String
    Dim base As String: base = RutaLocal()
    If base = "" Then
        MsgBox "No se pudo determinar la carpeta de este archivo.", vbCritical
        Exit Function
    End If
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim carpeta As String: carpeta = base & "\Reportes"
    If Not fso.FolderExists(carpeta) Then fso.CreateFolder carpeta
    Dim ruta As String
    ruta = carpeta & "\Reporte_Toma_" & Replace(sSucursal, " ", "") & "_" & Format(Now, "yyyy-mm-dd_hhnn") & ".xlsx"
    If fso.FileExists(ruta) Then ruta = Replace(ruta, ".xlsx", Format(Now, "ss") & ".xlsx")

    Dim wsR As Worksheet: Set wsR = Hj("RESULTADO")
    Dim n As Long: n = UltimaFila(wsR, 2) - RES_FILA + 1
    Dim enc As Variant: enc = wsR.Cells(RES_FILA - 1, 1).Resize(1, RES_COLS).Value
    Dim datos As Variant, dif() As Variant, m As Long, i As Long, j As Long

    Dim wb As Workbook
    On Error GoTo Fallo
    Application.ScreenUpdating = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Dim wsRes As Worksheet: Set wsRes = wb.Worksheets(1)
    wsRes.Name = "Resumen"
    Dim wsDif As Worksheet: Set wsDif = wb.Worksheets.Add(After:=wsRes)
    wsDif.Name = "Diferencias"
    Dim wsDet As Worksheet: Set wsDet = wb.Worksheets.Add(After:=wsDif)
    wsDet.Name = "Detalle completo"

    PrepararHojaReporte wsDif, enc
    PrepararHojaReporte wsDet, enc
    If n > 0 Then
        datos = wsR.Cells(RES_FILA, 1).Resize(n, RES_COLS).Value
        wsDet.Range("A2").Resize(n, RES_COLS).Value = datos
        ReDim dif(1 To n, 1 To RES_COLS)
        For i = 1 To n
            If EsDiferencia(CStr(datos(i, 1))) Then
                m = m + 1
                For j = 1 To RES_COLS
                    dif(m, j) = datos(i, j)
                Next
            End If
        Next
        If m > 0 Then wsDif.Range("A2").Resize(m, RES_COLS).Value = dif
    End If
    wsDif.Range("A1").Resize(m + 1, RES_COLS).AutoFilter
    wsDet.Range("A1").Resize(n + 1, RES_COLS).AutoFilter

    With wsRes
        .Columns("A").ColumnWidth = 30
        .Columns("B").ColumnWidth = 95
        .Range("A1").Value = "REPORTE DE TOMA DE INVENTARIO - " & UCase(sSucursal)
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A2").Value = "Generado: " & Format(Now, "dd-mm-yyyy hh:nn")
        .Range("A3").Value = "Toma iniciada: " & TextoInicio()
        .Range("A4").Value = "Reporte Aurora usado: " & Cfg(CFG_AURORA).Value & "  (archivo del " & Format(Cfg(CFG_AURORA_FECHA).Value, "dd-mm-yyyy hh:nn") & ")"

        .Range("A6").Value = "TOTALES"
        EscribirPares wsRes, 7, Array( _
            Array("Esperados en sistema", rEsp), Array("Encontrados", rEnc), Array("Faltantes", rFal), _
            Array("Sobrantes", rSob), Array("Ajustados", rAj), Array("En salida (NV / Picking)", rSal))
        .Range("B7:B12").HorizontalAlignment = xlLeft

        .Range("A14").Value = "QUÉ SIGNIFICA CADA ESTADO"
        EscribirPares wsRes, 15, Array( _
            Array(E_FALTANTE, "Está en el sistema (Aurora) en " & sSucursal & ", pero no se encontró físicamente."), _
            Array(E_SOB_NOEXISTE, "Se encontró físicamente, pero el serial no existe en Aurora."), _
            Array(E_SOB_OTRA, "Se encontró físicamente, pero Aurora lo tiene en otra sucursal (ver columna Sucursal Aurora)."), _
            Array(E_AJUSTADO, "Era faltante o sobrante y se justificó manualmente (ver columna Ajuste / comentario)."), _
            Array(E_SALIDA, "Tiene Nota de Venta o Picking asignado: va de salida y no se cuenta."), _
            Array(E_ENCONTRADO, "El serial está en el sistema y se encontró físicamente. Todo en orden."))

        .Range("A22").Value = "Hoja 'Diferencias': solo lo que requiere revisión (faltantes, sobrantes y ajustados)."
        .Range("A23").Value = "Hoja 'Detalle completo': todos los seriales de la toma."
        .Range("A6,A14").Font.Bold = True
        .Range("A7:A12,A15:A20").Font.Bold = True
        .Range("A22:A23").Font.Italic = True
    End With
    wsRes.Activate

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=ruta, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ThisWorkbook.Activate
    ExportarReporte = ruta
    Exit Function

Fallo:
    Dim e As String: e = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "No se pudo crear el reporte:" & vbLf & e, vbCritical, "Exportar"
    ExportarReporte = ""
End Function

Private Sub EscribirPares(ws As Worksheet, filaIni As Long, pares As Variant)
    Dim i As Long
    For i = 0 To UBound(pares)
        ws.Cells(filaIni + i, 1).Value = pares(i)(0)
        ws.Cells(filaIni + i, 2).Value = pares(i)(1)
    Next
End Sub

Private Sub PrepararHojaReporte(ws As Worksheet, enc As Variant)
    ws.Range("B:C").NumberFormat = "@"
    ws.Range("I:I").NumberFormat = "dd-mm-yyyy hh:mm"
    With ws.Range("A1").Resize(1, RES_COLS)
        .Value = enc
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(64, 64, 64)
    End With
    Dim anchos As Variant: anchos = Array(26, 20, 15, 55, 28, 15, 15, 15, 17, 45)
    Dim i As Long
    For i = 0 To RES_COLS - 1
        ws.Columns(i + 1).ColumnWidth = anchos(i)
    Next
End Sub

Private Sub ArchivarHistorial(ruta As String)
    Dim wsH As Worksheet: Set wsH = Hj("HISTORICO")
    Dim wsD As Worksheet: Set wsD = Hj("HIST_DETALLE")
    Dim nToma As Long: nToma = Application.WorksheetFunction.Max(wsH.Range("A:A")) + 1
    Dim cierre As Date: cierre = Now

    Dim fH As Long: fH = UltimaFila(wsH) + 1
    wsH.Cells(fH, 1).Resize(1, 13).Value = Array(nToma, Cfg(CFG_INICIO).Value, cierre, Cfg(CFG_AURORA).Value, _
        rEsp, rEnc, rFal, rSob, rAj, rSal, nDuplicados, dEscaneado.Count, Mid(ruta, InStrRev(ruta, "\") + 1))

    Dim wsR As Worksheet: Set wsR = Hj("RESULTADO")
    Dim n As Long: n = UltimaFila(wsR, 2) - RES_FILA + 1
    If n <= 0 Then Exit Sub
    Dim datos As Variant: datos = wsR.Cells(RES_FILA, 1).Resize(n, RES_COLS).Value
    Dim out() As Variant: ReDim out(1 To n, 1 To RES_COLS + 2)
    Dim i As Long, j As Long
    For i = 1 To n
        out(i, 1) = nToma
        out(i, 2) = cierre
        For j = 1 To RES_COLS
            out(i, j + 2) = datos(i, j)
        Next
    Next
    wsD.Cells(UltimaFila(wsD) + 1, 1).Resize(n, RES_COLS + 2).Value = out
End Sub

Private Sub LimpiarToma()
    Dim ws As Worksheet, ult As Long
    Set ws = Hj("ESCANEOS")
    ult = UltimaFila(ws)
    If ult >= 2 Then ws.Range(ws.Cells(2, 1), ws.Cells(ult, L_ANULADO)).ClearContents
    Set ws = Hj("AJUSTES")
    ult = UltimaFila(ws, 2)
    If ult >= 2 Then ws.Range(ws.Cells(2, 1), ws.Cells(ult, 3)).ClearContents
    Cfg(CFG_INICIO).ClearContents
    LimpiarResultado
    Hj("RESULTADO").Range("A5:G5").ClearContents
    Hj("RESULTADO").Range("A2").Value = "Presione ACTUALIZAR para ver el resultado de la toma en curso."
End Sub

' ---------------------------------------------------------------------------
' Pantalla INICIO
' ---------------------------------------------------------------------------

Public Sub ActualizarInicio()
    Dim ws As Worksheet: Set ws = Hj("INICIO")
    Dim cargado As Boolean
    If Not dAurora Is Nothing Then cargado = (dAurora.Count > 0)

    If IsEmpty(Cfg(CFG_INICIO).Value) Then
        ws.Range("C9").Value = "Sin escaneos aún (la toma comienza con el primer escaneo)"
    Else
        ' Se escribe la fecha como valor (no como texto): Excel invertiría día y mes
        ws.Range("C9").NumberFormat = "dd-mm-yyyy hh:mm"
        ws.Range("C9").Value = Cfg(CFG_INICIO).Value
    End If

    If Not cargado Then
        ws.Range("C10").Value = "NO CARGADO: presione ACTUALIZAR AURORA"
        ws.Range("C10").Font.Color = RGB(192, 0, 0)
        ws.Range("C11:C16").ClearContents
        Exit Sub
    End If
    ws.Range("C10").Value = Cfg(CFG_AURORA).Value & "   (archivo del " & Format(Cfg(CFG_AURORA_FECHA).Value, "dd-mm-yyyy hh:nn") & ")"
    ws.Range("C10").Font.Color = vbBlack
    ws.Range("C11").Value = dEsperado.Count
    ws.Range("C12").Value = nEncontrados
    ws.Range("C13").Value = dEsperado.Count - nEncontrados
    ws.Range("C14").Value = nSobrantes
    ws.Range("C15").Value = nDuplicados
    If dEsperado.Count > 0 Then ws.Range("C16").Value = nEncontrados / dEsperado.Count Else ws.Range("C16").Value = 0
End Sub

' ---------------------------------------------------------------------------
' Aviso de macros (el archivo se guarda con el aviso rojo; si las macros
' corren, al abrir se cambia a verde)
' ---------------------------------------------------------------------------

Public Sub BannerBloqueado()
    Dim ws As Worksheet: Set ws = Hj("INICIO")
    If ws.Range("B3").Interior.Color = RGB(192, 0, 0) Then Exit Sub
    ws.Range("B3").Value = "MACROS BLOQUEADAS: si ve este mensaje en rojo, el archivo NO funciona. Siga estos pasos:"
    ws.Range("B3").Interior.Color = RGB(192, 0, 0)
    ws.Range("B4").Value = "1) Si arriba aparece una barra amarilla con 'Habilitar edición' o 'Habilitar contenido', presiónela."
    ws.Range("B5").Value = "2) Si aparece una barra roja 'Microsoft bloqueó la ejecución de macros': cierre Excel, clic derecho sobre el archivo > Propiedades > marque 'Desbloquear' > Aceptar, y ábralo de nuevo."
    ws.Range("B6").Value = "3) Si no aparece ninguna barra o no existe 'Desbloquear': las macros están bloqueadas por política de la empresa. Envíe una captura de esta pantalla a TI."
    ws.Rows("4:6").Hidden = False
End Sub

Public Sub BannerOK()
    Dim ws As Worksheet: Set ws = Hj("INICIO")
    If ws.Range("B3").Interior.Color = RGB(0, 128, 0) Then Exit Sub
    ws.Range("B3").Value = "Archivo listo. Presione ESCANEAR para continuar la toma."
    ws.Range("B3").Interior.Color = RGB(0, 128, 0)
    ws.Range("B4:F6").ClearContents
    ws.Rows("4:6").Hidden = True
End Sub

Public Sub AlAbrir()
    BannerOK
    Hj("INICIO").Activate
    CargarIndices
    ActualizarInicio
    ThisWorkbook.Saved = True
End Sub
