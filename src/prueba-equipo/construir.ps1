# Genera demo\Prueba_Equipo.xlsm a partir de los fuentes de esta carpeta.
# Mismos requisitos que src\inventario\construir.ps1.
$ErrorActionPreference = 'Stop'
$src = Split-Path -Parent $MyInvocation.MyCommand.Path
$repo = Split-Path -Parent (Split-Path -Parent $src)
$outDir = Join-Path $repo 'demo'
$salida = Join-Path $outDir 'Prueba_Equipo.xlsm'

function Leer($nombre) { [IO.File]::ReadAllText((Join-Path $src $nombre), [Text.Encoding]::UTF8) }

$aurora = Get-ChildItem (Join-Path $outDir 'Aurora') -Filter '*.xlsx' -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $aurora) { throw 'Falta el reporte en demo\Aurora. Ejecute primero: python tools\generar_demo.py' }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add(-4167)
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'Prueba'
    $wd = $wb.Worksheets.Add([Type]::Missing, $ws); $wd.Name = 'Datos'
    $vbp = $wb.VBProject

    $m = $vbp.VBComponents.Add(1); $m.Name = 'modPrueba'
    $m.CodeModule.AddFromString((Leer 'modPrueba.bas'))
    $mt = $vbp.VBComponents.Add(1); $mt.Name = 'modTimer'
    $mt.CodeModule.AddFromString((Leer 'modTimer.bas'))
    $mc = $vbp.VBComponents.Add(1); $mc.Name = 'modConstruir'
    $mc.CodeModule.AddFromString((Leer 'modConstruir.bas'))
    $vbp.VBComponents.Item($wb.CodeName).CodeModule.AddFromString((Leer 'ThisWorkbook.txt'))

    $f = $vbp.VBComponents.Add(3); $f.Name = 'frmEscaner'
    $f.Properties.Item('Caption').Value = 'Prueba de pistola lectora'
    $f.Properties.Item('Width').Value = 460
    $f.Properties.Item('Height').Value = 300
    $xl.Run('modConstruir.CrearControles')
    $f.CodeModule.AddFromString((Leer 'frmEscaner.txt'))

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
    $vbp.VBComponents.Remove($vbp.VBComponents.Item('modConstruir'))
    $null = $lo.DataBodyRange.Delete()

    $xl.EnableEvents = $false
    $tmp = Join-Path $env:TEMP 'Prueba_Equipo.xlsm'
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
