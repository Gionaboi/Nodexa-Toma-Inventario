Option Explicit

' Solo se usa al generar el archivo; el script de construcción lo elimina después.
Public Sub Construir()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Prueba")
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 11
    ws.Cells.VerticalAlignment = xlCenter

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 24
    ws.Columns("D").ColumnWidth = 70
    ws.Columns("E").ColumnWidth = 60
    ws.Range("D:E").WrapText = True

    With ws.Range("B1")
        .Value = "PRUEBA DE EQUIPO · TOMA DE INVENTARIO"
        .Font.Size = 18
        .Font.Bold = True
    End With
    ws.Range("B2").Value = "Comprueba si este computador puede usar el archivo de inventario. Tarda unos 2 minutos y no modifica ninguna configuración del equipo."
    ws.Range("B2").Font.Color = RGB(89, 89, 89)

    With ws.Range("B3:E3")
        .Merge
        .Font.Size = 13
        .Font.Bold = True
        .IndentLevel = 1
    End With
    ws.Rows(3).RowHeight = 30
    Dim f As Long
    For f = 4 To 6
        ws.Range("B" & f & ":E" & f).Merge
        ws.Range("B" & f).WrapText = True
        ws.Rows(f).RowHeight = 20
    Next

    ws.Rows(8).RowHeight = 22
    ws.Rows(9).RowHeight = 22
    Boton ws, "IniciarPrueba", ChrW(&H25B6) & "  INICIAR PRUEBA", ws.Range("B8:C9"), RGB(0, 112, 192)
    Boton ws, "PruebaEscaner", "Repetir prueba de pistola", ws.Range("D8:D9"), RGB(89, 89, 89)

    With ws.Range("B11:E11")
        .Value = Array("Prueba", "Resultado", "Detalle", "Qué hacer")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(64, 64, 64)
    End With
    ws.Range("B12:B17").Value = Application.Transpose(Array( _
        "1. Macros (VBA)", "2. Excel y Windows", "3. Ubicación del archivo", _
        "4. Lectura Aurora (Power Query)", "5. Guardar archivos", "6. Pistola lectora"))
    ws.Range("B12:B17").Font.Bold = True
    ws.Range("C12:C17").HorizontalAlignment = xlCenter
    ws.Range("C12:C17").Font.Bold = True
    With ws.Range("B11:E17").Borders
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
    End With

    ws.Range("B19").Value = "RESULTADO FINAL"
    ws.Range("B19:C19").Font.Size = 14
    ws.Range("B19:C19").Font.Bold = True
    ws.Range("C19").HorizontalAlignment = xlCenter
    ws.Rows(19).RowHeight = 30

    ws.Range("B21").Value = "Si el resultado no es APTO, envíe una captura de esta pantalla. El archivo no necesita guardarse."
    ws.Range("B21").Font.Italic = True
    ws.Range("B21").Font.Color = RGB(89, 89, 89)

    ThisWorkbook.Worksheets("Datos").Visible = xlSheetHidden
    modPrueba.EstadoInicial
    ws.Range("A1").Select
End Sub

' Crea los controles de la ventana de prueba de pistola
Public Sub CrearControles()
    Dim d As Object: Set d = ThisWorkbook.VBProject.VBComponents("frmEscaner").Designer
    Dim c As Object
    Set c = Ctl(d, "Forms.Label.1", "lblTitulo", 12, 10, 420, 20)
    c.Caption = "Escanee cualquier etiqueta (idealmente el serial de un equipo)."
    c.Font.Bold = True
    Set c = Ctl(d, "Forms.TextBox.1", "txtScan", 12, 34, 420, 32)
    c.Font.Size = 18
    c.EnterKeyBehavior = False
    c.TabKeyBehavior = True
    Set c = Ctl(d, "Forms.Label.1", "lblResultado", 12, 76, 420, 140)
    c.WordWrap = True
    Set c = Ctl(d, "Forms.CommandButton.1", "btnSi", 12, 226, 130, 32)
    c.Caption = "Sí, es idéntico"
    Set c = Ctl(d, "Forms.CommandButton.1", "btnNo", 152, 226, 130, 32)
    c.Caption = "No, es distinto"
    Set c = Ctl(d, "Forms.CommandButton.1", "btnOmitir", 302, 226, 130, 32)
    c.Caption = "Cerrar / Omitir"
End Sub

Private Function Ctl(d As Object, tipo As String, nombre As String, l As Single, t As Single, w As Single, h As Single) As Object
    Dim c As Object
    Set c = d.Controls.Add(tipo, nombre)
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.Font.Size = 11
    Set Ctl = c
End Function

Private Sub Boton(ws As Worksheet, macro As String, texto As String, rng As Range, color As Long)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, rng.Left, rng.Top, rng.Width, rng.Height)
    s.OnAction = macro
    s.Fill.ForeColor.RGB = color
    s.Line.Visible = msoFalse
    With s.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = texto
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Size = 13
        .TextRange.Font.Fill.ForeColor.RGB = vbWhite
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub
