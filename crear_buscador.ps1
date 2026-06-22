# Crea la hoja BUSCADOR dentro de maestro_contratos.xlsx usando Excel COM.
$ErrorActionPreference = "Stop"
$path = "\\192.168.4.225\shared\GECI_GROUP_BI\00_PRE_PRODUCCION\Proyectos\maestro_contratos.xlsx"
$backup = "\\192.168.4.225\shared\GECI_GROUP_BI\00_PRE_PRODUCCION\Proyectos\maestro_contratos_backup_pre_buscador.xlsx"

# 0) Cerrar cualquier Excel abierto (descarta cambios sin guardar, p.ej. la 'S' suelta)
$p = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($p) { $p | Stop-Process -Force; Start-Sleep -Seconds 2 }

# 1) Copia de seguridad
Copy-Item -Path $path -Destination $backup -Force
Write-Output "Backup -> $backup"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($path)
  $S = "'Maestro Contratos'"
  $data = $wb.Sheets.Item("Maestro Contratos")

  # Restaurar cabecera A1 por si quedo dañada
  if ([string]$data.Range("A1").Value2 -ne "Codigo dimension") { $data.Range("A1").Value2 = "Codigo dimension" }

  # Ultima fila de datos (columna D = Proyecto GE)
  $L = $data.Cells($data.Rows.Count, 4).End(-4162).Row
  Write-Output "Ultima fila de datos: $L"

  # Refs de columnas
  function ColRef($c) { return $S + '!$' + $c + '$2:$' + $c + '$' + $L }
  $dD=ColRef 'D'; $dE=ColRef 'E'; $dF=ColRef 'F'; $dG=ColRef 'G'; $dH=ColRef 'H'
  $dC=ColRef 'C'; $dJ=ColRef 'J'; $dI=ColRef 'I'
  $dN=ColRef 'N'; $dP=ColRef 'P'; $dT=ColRef 'T'; $dV=ColRef 'V'; $dX=ColRef 'X'; $dZ=ColRef 'Z'
  $dAll = $S + '!$A$2:$AB$' + $L

  # Borrar hojas previas si existen
  foreach ($nm in @("BUSCADOR","_listas")) {
    foreach ($sh in @($wb.Sheets)) { if ($sh.Name -eq $nm) { $sh.Delete() } }
  }

  # Hoja oculta con la lista unica de codigos (la calcula Excel, sin volcado masivo)
  $lst = $wb.Sheets.Add()
  $lst.Name = "_listas"
  $lst.Range("A1").Value2 = "ProyectoGE (lista)"
  $lst.Range("A2").Formula2 = "=SORT(UNIQUE(" + $dD + "))"
  $wb.Application.CalculateFull()
  $lst.Visible = 0   # xlSheetHidden
  # nombre que apunta al rango derramado para la validacion
  try { $wb.Names.Item("lstProy").Delete() } catch {}
  $wb.Names.Add("lstProy", "=_listas!`$A`$2#") | Out-Null

  # codigo inicial = primer proyecto de los datos
  $firstCode = [string]$data.Range("D2").Value2
  Write-Output "Codigo inicial: $firstCode"

  # Hoja BUSCADOR
  $b = $wb.Sheets.Add()
  $b.Name = "BUSCADOR"
  $b.Move($wb.Sheets.Item(1))   # ponerla la primera

  # --- Titulo
  $b.Range("A1").Value2 = "BUSCADOR DE PROYECTOS  -  Maestro de Contratos"
  $b.Range("A1:K1").Merge()
  $b.Range("A1").Font.Size = 16
  $b.Range("A1").Font.Bold = $true
  $b.Range("A1").Interior.Color = 0x6E4A0C
  $b.Range("A1").Font.Color = 0xFFFFFF
  $b.Range("A1").RowHeight = 26

  # --- Seleccion + ficha
  $labels = @{ "A3"="Proyecto GE:"; "A4"="Descripcion:"; "A5"="Cliente:"; "A6"="Pais:"; "A7"="Empresa(s):"; "A8"="Project Manager:"; "A9"="Divisa(s):" }
  foreach ($k in $labels.Keys) { $b.Range($k).Value2 = $labels[$k]; $b.Range($k).Font.Bold = $true }

  $b.Range("C3").Value2 = $firstCode   # valor inicial
  $b.Range("C4").Formula2 = '=IFERROR(XLOOKUP($C$3,'+$dD+','+$dE+'),"")'
  $b.Range("C5").Formula2 = '=IFERROR(XLOOKUP($C$3,'+$dD+','+$dG+')&"  ("&XLOOKUP($C$3,'+$dD+','+$dF+')&")","")'
  $b.Range("C6").Formula2 = '=IFERROR(XLOOKUP($C$3,'+$dD+','+$dH+'),"")'
  $b.Range("C7").Formula2 = '=TEXTJOIN(", ",TRUE,UNIQUE(FILTER('+$dC+','+$dD+'=$C$3,"")))'
  $b.Range("C8").Formula2 = '=TEXTJOIN(", ",TRUE,UNIQUE(FILTER('+$dJ+',('+$dD+'=$C$3)*('+$dJ+'<>""),"")))'
  $b.Range("C9").Formula2 = '=TEXTJOIN(", ",TRUE,UNIQUE(FILTER('+$dI+','+$dD+'=$C$3,"")))'

  # caja de seleccion con validacion (lista de codigos)
  $dv = $b.Range("C3").Validation
  try { $dv.Delete() } catch {}
  $dv.Add(3, 1, 1, "=lstProy") | Out-Null
  $dv.IgnoreBlank = $true
  $dv.InCellDropdown = $true
  $b.Range("C3").Interior.Color = 0xFFF2CC
  $b.Range("C3:K3").Font.Bold = $true

  # --- Resumen por divisa
  $b.Range("A11").Value2 = "RESUMEN POR DIVISA  (importes netos, sin IVA)"
  $b.Range("A11").Font.Bold = $true; $b.Range("A11").Font.Color = 0x7C3AED
  $rh = @("Divisa","Contrato","Facturado","% Ejec","Pdte facturar","Cobrado","Pdte cobro s/fact","Pdte cobro s/contrato")
  for ($i=0; $i -lt $rh.Count; $i++) { $cell = $b.Cells(12, 1+$i); $cell.Value2 = $rh[$i]; $cell.Font.Bold = $true; $cell.Interior.Color = 0xE2EFDA }
  $b.Range("A13").Formula2 = '=IFERROR(UNIQUE(FILTER('+$dI+','+$dD+'=$C$3)),"-")'
  $b.Range("B13").Formula2 = '=IFERROR(SUMIFS('+$dN+','+$dD+',$C$3,'+$dI+',$A$13#),0)'
  $b.Range("C13").Formula2 = '=IFERROR(SUMIFS('+$dP+','+$dD+',$C$3,'+$dI+',$A$13#),0)'
  $b.Range("D13").Formula2 = '=IFERROR($C$13#/$B$13#,"")'
  $b.Range("E13").Formula2 = '=IFERROR(SUMIFS('+$dT+','+$dD+',$C$3,'+$dI+',$A$13#),0)'
  $b.Range("F13").Formula2 = '=IFERROR(SUMIFS('+$dV+','+$dD+',$C$3,'+$dI+',$A$13#),0)'
  $b.Range("G13").Formula2 = '=IFERROR(SUMIFS('+$dX+','+$dD+',$C$3,'+$dI+',$A$13#),0)'
  $b.Range("H13").Formula2 = '=IFERROR(SUMIFS('+$dZ+','+$dD+',$C$3,'+$dI+',$A$13#),0)'

  # --- Detalle por origen
  $b.Range("A17").Value2 = "DETALLE POR ORIGEN  (importes netos)"
  $b.Range("A17").Font.Bold = $true; $b.Range("A17").Font.Color = 0x155E75
  $dh = @("Empresa","Origen","Estado","Divisa","Contrato","Facturado","% Ejec","Pdte facturar","Cobrado","Pdte cobro s/fact","Pdte cobro s/contrato")
  for ($i=0; $i -lt $dh.Count; $i++) { $cell = $b.Cells(18, 1+$i); $cell.Value2 = $dh[$i]; $cell.Font.Bold = $true; $cell.Interior.Color = 0xD9E1F2 }
  $b.Range("A19").Formula2 = '=IFERROR(CHOOSECOLS(FILTER('+$dAll+','+$dD+'=$C$3),3,12,13,9,14,16,18,20,22,24,26),"(sin datos para este proyecto)")'

  # --- Buscador por texto (a la derecha)
  $b.Range("M2").Value2 = "Buscar por texto (codigo / descripcion / cliente):"
  $b.Range("M2").Font.Bold = $true
  $b.Range("M3").Interior.Color = 0xFFF2CC
  $b.Range("M3:O3").Merge()
  $b.Range("M5").Value2 = "Codigo"; $b.Range("N5").Value2 = "Descripcion"; $b.Range("O5").Value2 = "Cliente"
  $b.Range("M5:O5").Font.Bold = $true; $b.Range("M5:O5").Interior.Color = 0xE2EFDA
  $b.Range("M6").Formula2 = '=IF($M$3="","(escribe texto en la celda amarilla)",IFERROR(UNIQUE(FILTER(CHOOSECOLS('+$dAll+',4,5,7),(ISNUMBER(SEARCH(LOWER($M$3),LOWER('+$dD+')))+ISNUMBER(SEARCH(LOWER($M$3),LOWER('+$dE+')))+ISNUMBER(SEARCH(LOWER($M$3),LOWER('+$dG+'))))>0)),"(sin coincidencias)"))'

  # --- Formatos numericos
  $b.Range("B13:C13").NumberFormat = "#,##0"
  $b.Range("E13:H13").NumberFormat = "#,##0"
  $b.Range("D13").NumberFormat = "0.0%"
  $b.Range("E19:F60").NumberFormat = "#,##0"
  $b.Range("H19:K60").NumberFormat = "#,##0"
  $b.Range("G19:G60").NumberFormat = "0.0%"

  # Anchos de columna
  $widths = @{1=22; 2=24; 3=20; 4=8; 5=14; 6=14; 7=9; 8=14; 9=14; 10=16; 11=18; 13=12; 14=40; 15=34}
  foreach ($k in $widths.Keys) { $b.Columns.Item($k).ColumnWidth = $widths[$k] }
  $b.Range("C4:C9").Font.Size = 11

  # Nota
  $b.Range("A62").Value2 = "Nota: importes en la divisa de cada contrato, sin consolidar entre origenes (igual que la hoja Maestro Contratos). Existe ademas la version bruta (con IVA) en la hoja origen."
  $b.Range("A62").Font.Italic = $true; $b.Range("A62").Font.Color = 0x808080

  $b.Range("C3").Select() | Out-Null
  $wb.Application.CalculateFull()
  $wb.Save()
  $wb.Close($true)
  Write-Output "OK: hoja BUSCADOR creada y guardada."
}
finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
