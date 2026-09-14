let
    Origen = Excel.Workbook(File.Contents("{{ARCHIVO}}"), true),
    Hojas = Table.SelectRows(Origen, each [Kind] = "Sheet"),
    Datos = Hojas{0}[Data],
    Requeridas = {"Serial", "Código", "Descripción", "Condición", "Estado", "Bodega", "Sucursal", "Nota de Venta", "Picking"},
    Faltan = List.Difference(Requeridas, Table.ColumnNames(Datos)),
    Validado = if List.IsEmpty(Faltan) then Datos
        else error Error.Record("Formato Aurora", "El reporte Aurora no trae las columnas: " & Text.Combine(Faltan, ", ")),
    Seleccion = Table.SelectColumns(Validado, Requeridas),
    Texto = Table.TransformColumnTypes(Seleccion, List.Transform(Requeridas, each {_, type text})),
    Limpio = Table.TransformColumns(Texto, {{"Serial", each Text.Upper(Text.Trim(_)), type text}}),
    SinVacios = Table.SelectRows(Limpio, each [Serial] <> null and [Serial] <> "")
in
    SinVacios
