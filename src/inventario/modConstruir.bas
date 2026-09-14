Option Explicit

' Solo se usa al generar el archivo; el script de construcción lo elimina después.

Public Sub Construir()
    ConstruirInicio
    ConstruirResultado
    ConstruirTabla "ESCANEOS", Array("N°", "Fecha y hora", "Serial", "Resultado", "Detalle", "Origen", "Anulado"), _
        Array(7, 18, 22, 26, 80, 10, 9), "C:C", "B:B"
    FormatoEstados Hj("ESCANEOS").Range("D2:D30000")
    ConstruirTabla "AJUSTES", Array("Fecha", "Serial", "Motivo"), Array(18, 22, 70), "B:B", "A:A"
    Hj("AJUSTES").Range("E1").Value = "Para quitar un ajuste, borre su fila completa."
    ConstruirTabla "HISTORICO", Array("N° toma", "Inicio", "Cierre", "Reporte Aurora", "Esperados", "Encontrados", _
        "Faltantes", "Sobrantes", "Ajustados", "En salida", "Duplicados", "Seriales escaneados", "Archivo de reporte"), _
        Array(9, 17, 17, 45, 11, 12, 10, 10, 10, 10, 11, 12, 45), "", "B:C"
    ConstruirTabla "HIST_DETALLE", Array("N° toma", "Cierre", "Estado", "Serial", "SKU", "Descripción", "Bodega", _
        "Condición", "Estado Aurora", "Sucursal Aurora", "Escaneado", "Ajuste / comentario"), _
        Array(9, 17, 26, 22, 15, 55, 28, 15, 15, 15, 17, 45), "D:E", "B:B,K:K"
    FormatoEstados Hj("HIST_DETALLE").Range("C2:C300000")
    ConstruirConfig
    Hj("Datos").Visible = xlSheetHidden
    BannerBloqueado
    Hj("INICIO").Activate
    Hj("INICIO").Range("A1").Select
End Sub

Private Sub Base(ws As Worksheet)
    ws.Activate
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 11
    ws.Cells.VerticalAlignment = xlCenter
End Sub

Private Sub ConstruirInicio()
    Dim ws As Worksheet: Set ws = Hj("INICIO")
    Base ws
    ActiveWindow.DisplayGridlines = False
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 26
    ws.Columns("C").ColumnWidth = 60
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E:F").ColumnWidth = 20

    With ws.Range("B1")
        .Value = "TOMA DE INVENTARIO"
        .Font.Size = 20
        .Font.Bold = True
    End With
    ws.Range("B2").Value = "Escanee serial por serial. Todo se guarda solo y puede continuar otro día."
    ws.Range("B2").Font.Color = RGB(89, 89, 89)
    ws.Rows(1).RowHeight = 32

    With ws.Range("B3:F3")
        .Merge
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = vbWhite
        .IndentLevel = 1
    End With
    ws.Rows(3).RowHeight = 30
    Dim f As Long
    For f = 4 To 6
        ws.Range("B" & f & ":F" & f).Merge
        ws.Range("B" & f).WrapText = True
        ws.Range("B" & f).Font.Color = RGB(192, 0, 0)
        ws.Rows(f).RowHeight = 30
    Next

    With ws.Range("B8")
        .Value = "ESTADO DE LA TOMA"
        .Font.Bold = True
        .Font.Size = 12
    End With
    ws.Range("B9:B16").Value = Application.Transpose(Array("Toma en curso desde", "Reporte Aurora", _
        "Esperados (sistema)", "Encontrados", "Faltan por encontrar", "Sobrantes", "Duplicados", "Avance"))
    ws.Range("B9:B16").Font.Color = RGB(89, 89, 89)
    With ws.Range("C9:C16")
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlLeft
    End With
    ws.Range("C16").NumberFormat = "0%"
    With ws.Range("B9:C16").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = RGB(217, 217, 217)
    End With
    For f = 8 To 16
        ws.Rows(f).RowHeight = 24
    Next

    Boton ws, "Escanear", "ESCANEAR", ws.Range("E8:F10"), RGB(0, 140, 60), 20
    Boton ws, "ActualizarAurora", "Actualizar Aurora", ws.Range("E11:F11"), RGB(0, 112, 192), 12
    Boton ws, "VerResultado", "Ver resultado", ws.Range("E12:F12"), RGB(0, 112, 192), 12
    Boton ws, "AjusteManual", "Ajuste manual", ws.Range("E13:F13"), RGB(110, 110, 110), 12
    Boton ws, "FinalizarToma", "Finalizar toma", ws.Range("E15:F16"), RGB(192, 0, 0), 14

    With ws.Range("B18")
        .Value = "CÓMO SE USA"
        .Font.Bold = True
        .Font.Size = 12
    End With
    Dim pasos As Variant
    pasos = Array( _
        "1. Descargue desde Aurora el reporte de productos almacenados y guárdelo en la carpeta Aurora (junto a este archivo), reemplazando el anterior.", _
        "2. Presione ACTUALIZAR AURORA. Hágalo al empezar una toma o cuando descargue un reporte nuevo (los escaneos ya hechos se mantienen).", _
        "3. Presione ESCANEAR y escanee serial por serial. Puede cerrar la ventana y seguir otro día.", _
        "4. Si un equipo entró o salió durante la toma, use AJUSTE MANUAL.", _
        "5. Al terminar, presione FINALIZAR TOMA: se guarda en el historial y se crea el reporte en la carpeta Reportes.")
    For f = 0 To UBound(pasos)
        ws.Range("B" & 19 + f & ":F" & 19 + f).Merge
        ws.Range("B" & 19 + f).Value = pasos(f)
        ws.Range("B" & 19 + f).WrapText = True
        ws.Rows(19 + f).RowHeight = 30
    Next
    With ws.Range("B25")
        .Value = "Otras hojas: RESULTADO (detalle de la toma) · ESCANEOS (registro de cada escaneo) · AJUSTES · HISTORICO (tomas anteriores) · CONFIG (sucursal y bodegas)."
        .Font.Italic = True
        .Font.Color = RGB(89, 89, 89)
    End With
End Sub

Private Sub ConstruirResultado()
    Dim ws As Worksheet: Set ws = Hj("RESULTADO")
    Base ws
    Dim anchos As Variant: anchos = Array(26, 20, 15, 55, 28, 15, 15, 15, 17, 45, 4)
    Dim i As Long
    For i = 0 To UBound(anchos)
        ws.Columns(i + 1).ColumnWidth = anchos(i)
    Next
    ws.Range("B:C").NumberFormat = "@"
    ws.Range("I:I").NumberFormat = "dd-mm-yyyy hh:mm"

    With ws.Range("A1")
        .Value = "RESULTADO DE LA TOMA EN CURSO"
        .Font.Size = 16
        .Font.Bold = True
    End With
    ws.Range("A2").Value = "Presione ACTUALIZAR para ver el resultado de la toma en curso."
    ws.Range("A2").Font.Color = RGB(89, 89, 89)
    ws.Rows(1).RowHeight = 24
    ws.Rows(2).RowHeight = 24

    With ws.Range("A4:G4")
        .Value = Array("Esperados", "Encontrados", "Faltantes", "Sobrantes", "Ajustados", "En salida", "Avance")
        .Font.Color = RGB(89, 89, 89)
        .Interior.Color = RGB(242, 242, 242)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("A5:G5")
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("G5").NumberFormat = "0%"

    With ws.Range("A7:J7")
        .Value = Array("Estado", "Serial", "SKU", "Descripción", "Bodega", "Condición", "Estado Aurora", _
                       "Sucursal Aurora", "Escaneado", "Ajuste / comentario")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(64, 64, 64)
    End With
    ws.Range("A8").Select
    ActiveWindow.FreezePanes = True
    FormatoEstados ws.Range("A8:A30000")

    Boton ws, "VerResultado", "Actualizar", ws.Range("H1:H2"), RGB(0, 112, 192), 11
    Boton ws, "BotonExportar", "Exportar a Excel", ws.Range("I1:I2"), RGB(0, 140, 60), 11
    Boton ws, "IrInicio", "Volver al inicio", ws.Range("J1:J2"), RGB(110, 110, 110), 11
End Sub

Private Sub ConstruirTabla(nombre As String, encabezados As Variant, anchos As Variant, colsTexto As String, colsFecha As String)
    Dim ws As Worksheet: Set ws = Hj(nombre)
    Base ws
    Dim n As Long: n = UBound(encabezados) + 1
    Dim i As Long
    For i = 0 To n - 1
        ws.Columns(i + 1).ColumnWidth = anchos(i)
    Next
    If colsTexto <> "" Then ws.Range(colsTexto).NumberFormat = "@"
    If colsFecha <> "" Then ws.Range(colsFecha).NumberFormat = "dd-mm-yyyy hh:mm"
    With ws.Range("A1").Resize(1, n)
        .Value = encabezados
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(64, 64, 64)
    End With
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub ConstruirConfig()
    Dim ws As Worksheet: Set ws = Hj("CONFIG")
    Base ws
    ActiveWindow.DisplayGridlines = False
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 45

    With ws.Range("B1")
        .Value = "CONFIGURACIÓN"
        .Font.Size = 16
        .Font.Bold = True
    End With
    ws.Range("B3").Value = "Sucursal que se inventaría"
    ws.Range("C3").Value = "Casa Matriz"
    ws.Range("C3").Interior.Color = RGB(255, 242, 204)
    ws.Range("D3").Value = "Debe escribirse igual que en la columna Sucursal de Aurora."
    ws.Range("D3").Font.Color = RGB(89, 89, 89)

    ws.Range("B5:B8").Value = Application.Transpose(Array("Toma en curso desde", "Reporte Aurora cargado", _
        "Fecha del archivo Aurora", "Última lectura de Aurora"))
    ws.Range("C5,C7,C8").NumberFormat = "dd-mm-yyyy hh:mm"
    ws.Range("C5:C8").HorizontalAlignment = xlLeft
    ws.Range("B5:B8").Font.Color = RGB(89, 89, 89)

    ws.Range("B10").Value = "Bodegas de la sucursal: escriba No en las que NO se deben contar. Las bodegas nuevas se agregan solas con Sí."
    ws.Range("B10").Font.Bold = True
    With ws.Range("B11:C11")
        .Value = Array("Bodega", "¿Contar?")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(64, 64, 64)
    End With
    With ws.Range("C" & CFG_FILA_BODEGAS & ":C300").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Sí,No"
    End With
End Sub

' Colores por estado (formato condicional)
Private Sub FormatoEstados(rng As Range)
    CF rng, "FALTANTE", RGB(255, 199, 206), RGB(156, 0, 6)
    CF rng, "SOBRANTE - NO EXISTE", RGB(255, 220, 170), RGB(150, 70, 0)
    CF rng, "SOBRANTE - OTRA SUCURSAL", RGB(255, 220, 170), RGB(150, 70, 0)
    CF rng, "NO EXISTE", RGB(255, 220, 170), RGB(150, 70, 0)
    CF rng, "OTRA SUCURSAL", RGB(255, 220, 170), RGB(150, 70, 0)
    CF rng, "DUPLICADO", RGB(255, 235, 156), RGB(156, 87, 0)
    CF rng, "AJUSTADO", RGB(221, 235, 247), RGB(31, 78, 121)
    CF rng, "EN SALIDA", RGB(228, 223, 236), RGB(80, 60, 120)
    CF rng, "BODEGA EXCLUIDA", RGB(237, 237, 237), RGB(89, 89, 89)
    CF rng, "ENCONTRADO", RGB(198, 239, 206), RGB(0, 97, 0)
End Sub

Private Sub CF(rng As Range, texto As String, fondo As Long, letra As Long)
    Dim fc As FormatCondition
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & texto & """")
    fc.Interior.Color = fondo
    fc.Font.Color = letra
    fc.Font.Bold = True
End Sub

Private Sub Boton(ws As Worksheet, macro As String, texto As String, rng As Range, color As Long, tam As Single)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, rng.Left + 2, rng.Top + 2, rng.Width - 4, rng.Height - 4)
    s.OnAction = macro
    s.Fill.ForeColor.RGB = color
    s.Line.Visible = msoFalse
    s.Placement = xlFreeFloating
    With s.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = texto
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Size = tam
        .TextRange.Font.Fill.ForeColor.RGB = vbWhite
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub

' Controles de la ventana de escaneo
Public Sub CrearControles()
    Dim d As Object: Set d = ThisWorkbook.VBProject.VBComponents("frmEscaneo").Designer
    Dim c As Object
    Set c = Ctl(d, "Forms.Label.1", "lblTitulo", 12, 8, 520, 18, 11)
    c.Caption = "Escanee serial por serial. Cada escaneo se guarda solo."
    c.Font.Bold = True
    Set c = Ctl(d, "Forms.TextBox.1", "txtScan", 12, 30, 520, 36, 20)
    c.EnterKeyBehavior = False
    c.TabKeyBehavior = True
    Set c = Ctl(d, "Forms.Label.1", "lblEstado", 12, 74, 520, 44, 24)
    c.Font.Bold = True
    c.TextAlign = 2
    c.BackStyle = 1
    Set c = Ctl(d, "Forms.Label.1", "lblSerial", 12, 124, 520, 24, 14)
    c.Font.Bold = True
    Set c = Ctl(d, "Forms.Label.1", "lblDetalle", 12, 150, 520, 54, 11)
    c.WordWrap = True
    Set c = Ctl(d, "Forms.Label.1", "lblContador", 12, 210, 520, 42, 13)
    c.Font.Bold = True
    c.WordWrap = True
    c.BackColor = RGB(242, 242, 242)
    Set c = Ctl(d, "Forms.CommandButton.1", "btnDeshacer", 12, 262, 170, 32, 11)
    c.Caption = "Deshacer último"
    c.TakeFocusOnClick = False
    Set c = Ctl(d, "Forms.CommandButton.1", "btnCerrar", 362, 262, 170, 32, 11)
    c.Caption = "Cerrar"
    c.TakeFocusOnClick = False
    Set c = Ctl(d, "Forms.Label.1", "lblAyuda", 12, 302, 520, 44, 9)
    c.Caption = "Etiqueta dañada: escriba el serial a mano y presione Enter." & vbLf & _
                "Mientras esta ventana esté abierta, todo lo que lea la pistola entra al inventario. " & _
                "Para usar la pistola en otra cosa (por ejemplo, buscar en Google), cierre esta ventana primero."
    c.WordWrap = True
    c.ForeColor = RGB(89, 89, 89)
End Sub

Private Function Ctl(d As Object, tipo As String, nombre As String, l As Single, t As Single, w As Single, h As Single, tam As Single) As Object
    Dim c As Object
    Set c = d.Controls.Add(tipo, nombre)
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.Font.Size = tam
    Set Ctl = c
End Function
