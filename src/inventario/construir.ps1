# Genera demo\Toma_Inventario_Demo.xlsm a partir de los fuentes de esta carpeta.
#
# Requisitos:
#   - Excel de escritorio (Windows).
#   - Excel > Opciones > Centro de confianza > Configuración de macros >
#     "Confiar en el acceso al modelo de objetos de proyectos de VBA" activado
#     (solo para generar el archivo; desactívelo después).
#   - El reporte demo ya generado:  python tools\generar_demo.py
$ErrorActionPreference = 'Stop'
$src = Split-Path -Parent $MyInvocation.MyCommand.Path
$repo = Split-Path -Parent (Split-Path -Parent $src)
$outDir = Join-Path $repo 'demo'
$salida = Join-Path $outDir 'Toma_Inventario_Demo.xlsm'

function Leer($nombre) { [IO.File]::ReadAllText((Join-Path $src $nombre), [Text.Encoding]::UTF8) }

$aurora = Get-ChildItem (Join-Path $outDir 'Aurora') -Filter '*.xlsx' -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $aurora) { throw 'Falta el reporte en demo\Aurora. Ejecute primero: python tools\generar_demo.py' }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add(-4167)
    $wb.Worksheets.Item(1).Name = 'INICIO'
    foreach ($n in 'RESULTADO', 'ESCANEOS', 'AJUSTES', 'HISTORICO', 'HIST_DETALLE', 'CONFIG', 'Datos') {
        $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
        $ws.Name = $n
    }
    $vbp = $wb.VBProject

    foreach ($mod in 'modUtil', 'modInventario', 'modTimer', 'modConstruir') {
        $m = $vbp.VBComponents.Add(1); $m.Name = $mod
        $m.CodeModule.AddFromString((Leer "$mod.bas"))
    }
    $vbp.VBComponents.Item($wb.CodeName).CodeModule.AddFromString((Leer 'ThisWorkbook.txt'))

    # Ventana de escaneo
    $f = $vbp.VBComponents.Add(3); $f.Name = 'frmEscaneo'
    $f.Properties.Item('Caption').Value = 'Escaneo de inventario'
    $f.Properties.Item('Width').Value = 556
    $f.Properties.Item('Height').Value = 380
    $xl.Run('modConstruir.CrearControles')
    $f.CodeModule.AddFromString((Leer 'frmEscaneo.txt'))

    # Consulta Power Query + tabla de destino (hoja oculta "Datos")
    $wd = $wb.Worksheets.Item('Datos')
    $mcode = (Leer 'Aurora.m').Replace('{{ARCHIVO}}', $aurora.FullName)
    $null = $wb.Queries.Add('Aurora', $mcode)
    $conn = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Aurora;Extended Properties=""'
    $lo = $wd.ListObjects.Add(0, $conn, [Type]::Missing, 1, $wd.Range('A1'))
    $lo.Name = 'tblAurora'
    $lo.QueryTable.CommandType = 2
    $lo.QueryTable.CommandText = 'SELECT * FROM [Aurora]'
    $lo.QueryTable.BackgroundQuery = $false
    $null = $lo.QueryTable.Refresh($false)
    Write-Output "Filas leidas del reporte: $($lo.ListRows.Count)"

    $xl.Run('modConstruir.Construir')
    $xl.Run('modInventario.InicializarConfig')   # llena la lista de bodegas en CONFIG
    $vbp.VBComponents.Remove($vbp.VBComponents.Item('modConstruir'))

    # Se entrega sin datos cargados: el usuario los carga con ACTUALIZAR AURORA
    $null = $lo.DataBodyRange.Delete()
    $wb.Worksheets.Item('INICIO').Activate()

    $xl.EnableEvents = $false
    $tmp = Join-Path $env:TEMP 'Toma_Inventario_Demo.xlsm'
    if (Test-Path $tmp) { Remove-Item $tmp }
    $wb.SaveAs($tmp, 52)
    $wb.Close($false)
    Copy-Item $tmp $salida -Force
    Write-Output "Generado: $salida"
}
finally {
    $xl.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
