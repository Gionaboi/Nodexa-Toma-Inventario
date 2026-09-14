Option Explicit

' Filas de la tabla de resultados en la hoja "Prueba"
Public Const F_MACROS As Long = 12
Public Const F_EQUIPO As Long = 13
Public Const F_UBICACION As Long = 14
Public Const F_AURORA As Long = 15
Public Const F_ESCRITURA As Long = 16
Public Const F_ESCANER As Long = 17
Public Const F_FINAL As Long = 19

Public Const R_OK As String = "OK"
Public Const R_ADV As String = "ADVERTENCIA"
Public Const R_ERR As String = "ERROR"
Public Const R_NP As String = "NO PROBADO"

Private Function Hoja() As Worksheet
    Set Hoja = ThisWorkbook.Worksheets("Prueba")
End Function

' ---------------------------------------------------------------------------
' Estado del archivo
' ---------------------------------------------------------------------------

' Estado con el que se guarda el archivo: si las macros no corren, esto es lo
' que el usuario ve (banner rojo con instrucciones). Nunca queda guardado un
' "OK" de otro equipo, para evitar falsos positivos.
Public Sub EstadoInicial()
    Dim ws As Worksheet: Set ws = Hoja()
    With ws.Range("B3")
        .Value = "MACROS BLOQUEADAS: si ve este mensaje en rojo, las macros NO se están ejecutando. Siga estos pasos:"
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = vbWhite
    End With
    ws.Range("B4").Value = "1) Si arriba aparece una barra amarilla con 'Habilitar edición' o 'Habilitar contenido', presiónela."
    ws.Range("B5").Value = "2) Si aparece una barra roja 'Microsoft bloqueó la ejecución de macros': cierre Excel, clic derecho sobre el archivo > Propiedades > marque 'Desbloquear' > Aceptar, y ábralo de nuevo."
    ws.Range("B6").Value = "3) Si no aparece ninguna barra, o no existe la opción 'Desbloquear': las macros están bloqueadas por política de la empresa. Envíe una captura de esta pantalla a TI (código MAC-01)."
    ws.Range("B4:B6").Font.Color = RGB(192, 0, 0)

    Dim f As Long
    For f = F_EQUIPO To F_ESCANER
        Escribir f, "SIN PROBAR", "", ""
    Next
    Escribir F_MACROS, R_ERR, "Las macros no se ejecutaron al abrir el archivo.", "Siga los pasos en rojo de arriba."
    FinalPendiente
End Sub

' Se ejecuta al abrir: si esto corre, las macros funcionan.
Public Sub AlAbrir()
    Dim ws As Worksheet: Set ws = Hoja()
    ws.Activate
    With ws.Range("B3")
        .Value = "MACROS FUNCIONANDO. Presione el botón INICIAR PRUEBA."
        .Interior.Color = RGB(0, 128, 0)
        .Font.Color = vbWhite
    End With
    ws.Range("B4:E6").ClearContents
    Dim f As Long
    For f = F_EQUIPO To F_ESCANER
        Escribir f, "SIN PROBAR", "", ""
    Next
    Escribir F_MACROS, R_OK, "Las macros se ejecutan correctamente.", ""
    FinalPendiente
    ThisWorkbook.Saved = True
End Sub

Private Sub FinalPendiente()
    With Hoja().Cells(F_FINAL, 3)
        .Value = "PENDIENTE"
        .Interior.Color = RGB(237, 237, 237)
        .Font.Color = RGB(89, 89, 89)
    End With
    Hoja().Cells(F_FINAL, 4).Value = ""
End Sub

Public Sub Escribir(fila As Long, resultado As String, detalle As String, accion As String)
    With Hoja()
        .Cells(fila, 3).Value = resultado
        .Cells(fila, 4).Value = detalle
        .Cells(fila, 5).Value = accion
        Select Case resultado
            Case R_OK: Pintar .Cells(fila, 3), RGB(198, 239, 206), RGB(0, 97, 0)
            Case R_ADV: Pintar .Cells(fila, 3), RGB(255, 235, 156), RGB(156, 87, 0)
            Case R_ERR: Pintar .Cells(fila, 3), RGB(255, 199, 206), RGB(156, 0, 6)
            Case Else: Pintar .Cells(fila, 3), RGB(237, 237, 237), RGB(89, 89, 89)
        End Select
        .Rows(fila).AutoFit
    End With
End Sub

Private Sub Pintar(c As Range, fondo As Long, letra As Long)
    c.Interior.Color = fondo
    c.Font.Color = letra
End Sub

' ---------------------------------------------------------------------------
' Botones
' ---------------------------------------------------------------------------

Public Sub IniciarPrueba()
    Dim f As Long
    For f = F_EQUIPO To F_ESCANER
        Escribir f, "PROBANDO...", "", ""
    Next
    FinalPendiente
    DoEvents

    ProbarEquipo
    ProbarUbicacion
    Application.StatusBar = "Leyendo reporte Aurora con Power Query (puede tardar unos segundos)..."
    ProbarAurora
    Application.StatusBar = "Probando guardado de archivos..."
    ProbarEscritura
    Application.StatusBar = False

    PruebaEscaner
End Sub

Public Sub PruebaEscaner()
    frmEscaner.Show
    EvaluarFinal
End Sub

Public Sub EvaluarFinal()
    Dim ws As Worksheet: Set ws = Hoja()
    Dim f As Long, r As String, hayErr As Boolean, hayAdv As Boolean
    For f = F_MACROS To F_ESCANER
        r = ws.Cells(f, 3).Value
        If r = R_ERR Then
            hayErr = True
        ElseIf r <> R_OK Then
            hayAdv = True
        End If
    Next

    Dim firma As String
    firma = "Equipo: " & Environ("COMPUTERNAME") & " · Usuario: " & Environ("USERNAME") & " · " & Format(Now, "dd-mm-yyyy hh:nn")
    With ws.Cells(F_FINAL, 3)
        If hayErr Then
            .Value = "NO APTO"
            Pintar ws.Cells(F_FINAL, 3), RGB(192, 0, 0), vbWhite
            ws.Cells(F_FINAL, 4).Value = "Revise las filas en ERROR y siga la columna 'Qué hacer'. " & firma
        ElseIf hayAdv Then
            .Value = "APTO CON ADVERTENCIAS"
            Pintar ws.Cells(F_FINAL, 3), RGB(255, 192, 0), vbBlack
            ws.Cells(F_FINAL, 4).Value = "Funciona, pero revise las filas en amarillo o gris. " & firma
        Else
            .Value = "APTO"
            Pintar ws.Cells(F_FINAL, 3), RGB(0, 128, 0), vbWhite
            ws.Cells(F_FINAL, 4).Value = "Este equipo puede usar el archivo de inventario. " & firma
        End If
    End With
    ThisWorkbook.Saved = True
End Sub

' ---------------------------------------------------------------------------
' Pruebas
' ---------------------------------------------------------------------------

Private Sub ProbarEquipo()
    Dim bits As String
    #If Win64 Then
        bits = "64 bits"
    #Else
        bits = "32 bits"
    #End If
    Dim det As String
    det = "Excel " & Application.Version & " (compilación " & Application.Build & ", " & bits & ") · " & Application.OperatingSystem
    If Val(Application.Version) < 16 Then
        Escribir F_EQUIPO, R_ERR, det, "Se requiere Excel 2016 o superior (Microsoft 365 recomendado). Solicitar a TI."
    Else
        Escribir F_EQUIPO, R_OK, det, ""
    End If
End Sub

Private Sub ProbarUbicacion()
    Dim p As String: p = ThisWorkbook.Path
    Dim rl As String: rl = RutaLocal()
    Dim lp As String: lp = LCase(rl)
    If rl = "" Then
        Escribir F_UBICACION, R_ERR, "No se pudo determinar la carpeta local del archivo: " & p, _
            "Copie la carpeta Prueba_Equipo completa al Escritorio o Documentos y ábralo desde ahí (código FS-02)."
    ElseIf InStr(lp, "\temp\") > 0 Or InStr(lp, "content.outlook") > 0 Or InStr(lp, "inetcache") > 0 Then
        Escribir F_UBICACION, R_ERR, "El archivo se abrió directo desde el correo, un .zip o una carpeta temporal: " & rl, _
            "Extraiga/guarde la carpeta Prueba_Equipo completa en el Escritorio o Documentos y ábralo desde ahí."
    ElseIf ThisWorkbook.ReadOnly Then
        Escribir F_UBICACION, R_ADV, "El archivo está abierto en modo solo lectura: " & rl, _
            "Cierre otras copias abiertas del archivo o revise que la carpeta permita escribir."
    ElseIf LCase(Left(p, 4)) = "http" Then
        Escribir F_UBICACION, R_OK, "OneDrive sincronizado. Carpeta local: " & rl, ""
    Else
        Escribir F_UBICACION, R_OK, "Carpeta local: " & rl, ""
    End If
End Sub

Private Sub ProbarAurora()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim base As String: base = RutaLocal()
    If base = "" Then
        Escribir F_AURORA, R_NP, "No se probó porque no se pudo determinar la carpeta del archivo.", "Resuelva primero la fila 'Ubicación del archivo'."
        Exit Sub
    End If

    Dim carpeta As String: carpeta = base & "\Aurora"
    If Not fso.FolderExists(carpeta) Then
        Escribir F_AURORA, R_ERR, "No existe la carpeta 'Aurora' junto a este archivo.", _
            "Cree una carpeta llamada Aurora al lado de este archivo y copie dentro el reporte descargado de Aurora."
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
        Escribir F_AURORA, R_ERR, "La carpeta Aurora no contiene ningún archivo .xlsx.", _
            "Descargue el reporte de productos almacenados desde Aurora y guárdelo en la carpeta Aurora."
        Exit Sub
    End If

    Dim q As WorkbookQuery: Set q = ThisWorkbook.Queries("Aurora")
    q.Formula = CambiarRuta(q.Formula, ultimo.Path)

    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("Datos").ListObjects(1)
    Dim t As Single: t = Timer
    On Error GoTo Fallo
    Application.DisplayAlerts = False
    lo.QueryTable.Refresh BackgroundQuery:=False
    Application.DisplayAlerts = True
    On Error GoTo 0

    Dim total As Long, lc As Long
    If Not lo.DataBodyRange Is Nothing Then
        total = lo.ListRows.Count
        lc = Application.WorksheetFunction.CountIf(lo.ListColumns("Sucursal").DataBodyRange, "Casa Matriz")
    End If
    Dim det As String
    det = "Leído '" & ultimo.Name & "' en " & Format(Timer - t, "0.0") & " s: " & _
          Format(total, "#,##0") & " seriales, " & Format(lc, "#,##0") & " de Casa Matriz."
    If lc = 0 Then
        Escribir F_AURORA, R_ADV, det, "El reporte no trae seriales de Casa Matriz. Verifique que descargó el reporte de productos almacenados de todas las sucursales."
    Else
        Escribir F_AURORA, R_OK, det, ""
    End If
    Exit Sub

Fallo:
    Dim msg As String: msg = Err.Description
    Application.DisplayAlerts = True
    Escribir F_AURORA, R_ERR, "Power Query no pudo leer el reporte: " & msg, _
        "Si el mensaje menciona columnas, el formato del reporte Aurora cambió: avise para ajustar el archivo. " & _
        "Si el reporte está abierto en Excel, ciérrelo y repita. En otro caso envíe captura (código PQ-01)."
End Sub

Private Sub ProbarEscritura()
    Dim base As String: base = RutaLocal()
    If base = "" Then
        Escribir F_ESCRITURA, R_NP, "No se probó porque no se pudo determinar la carpeta del archivo.", "Resuelva primero la fila 'Ubicación del archivo'."
        Exit Sub
    End If

    ' Nombre único: nunca puede coincidir con (ni borrar) un archivo del usuario
    Dim ruta As String: ruta = base & "\_prueba_escritura_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    If Len(Dir(ruta)) > 0 Then Err.Raise vbObjectError + 2, , "Ya existe un archivo de prueba con ese nombre."
    Dim wb As Workbook
    On Error GoTo Fallo
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Range("A1").Value = "Prueba de escritura " & Now
    wb.SaveAs Filename:=ruta, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Set wb = Nothing
    If Len(Dir(ruta)) = 0 Then Err.Raise vbObjectError + 1, , "El archivo no quedó guardado."
    Kill ruta
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Escribir F_ESCRITURA, R_OK, "Se creó y se borró un archivo temporal de prueba en la carpeta: puede guardar archivos (así se exportará el reporte).", ""
    Exit Sub

Fallo:
    Dim msg As String: msg = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Escribir F_ESCRITURA, R_ERR, "No se pudo guardar un archivo en la carpeta: " & msg, _
        "Mueva la carpeta al Escritorio o Documentos. Si es una carpeta de red o SharePoint, pida a TI permiso de escritura (código FS-01)."
End Sub

' ---------------------------------------------------------------------------
' Utilidades
' ---------------------------------------------------------------------------

' ThisWorkbook.Path devuelve una URL (https://...) cuando el archivo está en
' OneDrive. Power Query y el sistema de archivos necesitan la ruta local.
Public Function RutaLocal() As String
    Dim p As String: p = ThisWorkbook.Path
    If LCase(Left(p, 4)) <> "http" Then
        RutaLocal = p
        Exit Function
    End If

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim partes() As String: partes = Split(DecodificarURL(p), "/")
    Dim raices As Variant, r As Variant, i As Long, j As Long, resto As String
    raices = Array(Environ("OneDriveCommercial"), Environ("OneDriveConsumer"), Environ("OneDrive"))
    For Each r In raices
        If Len(r) > 0 Then
            ' Del sufijo más largo al más corto: el primero que exista es la carpeta
            For i = 3 To UBound(partes)
                resto = ""
                For j = i To UBound(partes)
                    resto = resto & "\" & partes(j)
                Next
                If fso.FolderExists(r & resto) Then
                    RutaLocal = r & resto
                    Exit Function
                End If
            Next
        End If
    Next
    RutaLocal = ""
End Function

Private Function DecodificarURL(s As String) As String
    Dim i As Long, c As String, res As String, b1 As Long, b2 As Long, b3 As Long
    On Error GoTo Fin
    i = 1
    Do While i <= Len(s)
        c = Mid(s, i, 1)
        If c = "%" And i + 2 <= Len(s) Then
            b1 = CLng("&H" & Mid(s, i + 1, 2))
            If b1 < &H80 Then
                res = res & Chr(b1): i = i + 3
            ElseIf b1 >= &HE0 And i + 8 <= Len(s) Then
                b2 = CLng("&H" & Mid(s, i + 4, 2)): b3 = CLng("&H" & Mid(s, i + 7, 2))
                res = res & ChrW((b1 And &HF) * 4096 + (b2 And &H3F) * 64 + (b3 And &H3F)): i = i + 9
            ElseIf b1 >= &HC0 And i + 5 <= Len(s) Then
                b2 = CLng("&H" & Mid(s, i + 4, 2))
                res = res & ChrW((b1 And &H1F) * 64 + (b2 And &H3F)): i = i + 6
            Else
                res = res & c: i = i + 1
            End If
        Else
            res = res & c: i = i + 1
        End If
    Loop
    DecodificarURL = res
    Exit Function
Fin:
    DecodificarURL = s
End Function

' Reemplaza la ruta dentro de File.Contents("...") en la fórmula M
Private Function CambiarRuta(formula As String, ruta As String) As String
    Dim marca As String: marca = "File.Contents("""
    Dim ini As Long, fin As Long
    ini = InStr(formula, marca) + Len(marca)
    fin = InStr(ini, formula, """)")
    CambiarRuta = Left(formula, ini - 1) & Replace(ruta, """", """""") & Mid(formula, fin)
End Function
