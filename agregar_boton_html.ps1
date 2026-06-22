# Anade un boton en la hoja BUSCADOR que abre el HTML a pantalla completa (hipervinculo, sin macros).
$ErrorActionPreference = "Stop"
$path = "\\192.168.4.225\shared\GECI_GROUP_BI\00_PRE_PRODUCCION\Proyectos\maestro_contratos.xlsx"
$html = "\\192.168.4.225\shared\GECI_GROUP_BI\00_PRE_PRODUCCION\Proyectos\maestro_contratos.html"

$p = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($p) { $p | Stop-Process -Force; Start-Sleep -Seconds 2 }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($path)
  $b = $wb.Sheets.Item("BUSCADOR")

  # quitar boton previo si existe
  foreach ($sh in @($b.Shapes)) { if ($sh.Name -eq "btnAbrirHTML") { $sh.Delete() } }

  $anchor = $b.Range("E3")
  # msoShapeRoundedRectangle = 5
  $shp = $b.Shapes.AddShape(5, $anchor.Left + 4, $anchor.Top, 240, 26)
  $shp.Name = "btnAbrirHTML"
  $shp.TextFrame2.TextRange.Text = "ABRIR EXPLORADOR (HTML) a pantalla completa"
  $shp.TextFrame2.TextRange.Font.Size = 10
  $shp.TextFrame2.TextRange.Font.Bold = $true
  $shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
  $shp.Fill.ForeColor.RGB = 0x9C6E0C     # azul/teal oscuro (BGR)
  $shp.Line.Visible = $false
  $shp.TextFrame2.VerticalAnchor = 3     # msoAnchorMiddle

  # hipervinculo al HTML (se abre en el navegador por defecto)
  $b.Hyperlinks.Add($shp, $html) | Out-Null

  $wb.Save()
  $wb.Close($true)
  Write-Output "OK: boton 'ABRIR EXPLORADOR (HTML)' anadido a la hoja BUSCADOR."
}
finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
