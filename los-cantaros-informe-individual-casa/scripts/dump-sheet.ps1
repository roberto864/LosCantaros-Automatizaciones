# dump-sheet.ps1 — vuelca una hoja de un .xlsx a TSV, sin abrir Excel.
#
# Lee el XML interno del archivo, asi que no modifica nada: es seguro usarlo
# sobre las planillas de produccion (que tienen graficos, tablas dinamicas y
# segmentaciones que openpyxl destruiria).
#
# Uso:
#   .\dump-sheet.ps1 -Path "...\Cartola Mensual Gastos Comunes Año 2026.xlsx" -Sheet "Cartola" -Out c2026.tsv
#   .\dump-sheet.ps1 -Path "...\archivo.xlsx" -List          # solo lista las hojas
#
# -Sheet acepta el NOMBRE de la hoja (o parte de el, sin distinguir mayusculas).
# No asumas que la hoja principal es sheet1.xml: en la Cartola 2023 es sheet2.xml.
# Este script resuelve el nombre contra workbook.xml + rels, que es lo correcto.
#
# Salida TSV: primera columna = numero de fila de Excel, luego las columnas A..T.
# Asi los numeros de fila del TSV calzan con los de la planilla y puedes citarlos.
param(
  [Parameter(Mandatory=$true)][string]$Path,
  [string]$Sheet = "",
  [string]$Out = "",
  [switch]$List,
  [int]$MaxCol = 20
)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# OneDrive y Excel dejan las planillas bloqueadas para lectura exclusiva bastante
# seguido. Copiar a temp siempre funciona y evita tocar el original.
$src = $Path
try {
  $probe = [System.IO.File]::Open($Path, 'Open', 'Read', 'Read'); $probe.Close()
} catch {
  $src = Join-Path ([System.IO.Path]::GetTempPath()) ("lc_" + [Guid]::NewGuid().ToString("N") + ".xlsx")
  Copy-Item -LiteralPath $Path -Destination $src -Force
  Write-Host "(archivo bloqueado; se trabajo sobre una copia temporal)"
}

$zip = [System.IO.Compression.ZipFile]::OpenRead($src)
function ReadEntry([string]$name){
  $e = $zip.Entries | Where-Object { $_.FullName -eq $name }
  if(-not $e){ return $null }
  $sr = New-Object System.IO.StreamReader($e.Open())
  $t = $sr.ReadToEnd(); $sr.Close(); return $t
}

# --- resolver nombre de hoja -> worksheets/sheetN.xml -------------------------
$wbXml = New-Object System.Xml.XmlDocument; $wbXml.LoadXml((ReadEntry "xl/workbook.xml"))
$relXml = New-Object System.Xml.XmlDocument; $relXml.LoadXml((ReadEntry "xl/_rels/workbook.xml.rels"))
$relMap = @{}
foreach($r in $relXml.DocumentElement.ChildNodes){ $relMap[$r.GetAttribute("Id")] = $r.GetAttribute("Target") }

$sheets = @()
$ns = New-Object System.Xml.XmlNamespaceManager($wbXml.NameTable)
$ns.AddNamespace("d","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
$ns.AddNamespace("r","http://schemas.openxmlformats.org/officeDocument/2006/relationships")
foreach($sh in $wbXml.SelectNodes("//d:sheets/d:sheet",$ns)){
  $rid = $sh.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships")
  $target = $relMap[$rid] -replace '^/xl/','' -replace '^xl/',''
  $sheets += [pscustomobject]@{ Name = $sh.GetAttribute("name"); Part = "xl/$target" }
}

if($List -or $Sheet -eq ""){
  $zip.Dispose()
  if($src -ne $Path){ Remove-Item $src -Force }
  $sheets | ForEach-Object { Write-Output ("{0}`t{1}" -f $_.Name, $_.Part) }
  return
}

$match = $sheets | Where-Object { $_.Name -eq $Sheet }
if(-not $match){ $match = $sheets | Where-Object { $_.Name -like "*$Sheet*" } }
if(-not $match){
  $zip.Dispose(); if($src -ne $Path){ Remove-Item $src -Force }
  throw "No hay hoja que calce con '$Sheet'. Hojas disponibles: " + ($sheets.Name -join ", ")
}
$match = @($match)[0]

# --- shared strings -----------------------------------------------------------
$ss = New-Object System.Collections.ArrayList
$sstText = ReadEntry "xl/sharedStrings.xml"
if($sstText){
  $sx = New-Object System.Xml.XmlDocument; $sx.LoadXml($sstText)
  foreach($si in $sx.DocumentElement.ChildNodes){ [void]$ss.Add($si.InnerText) }
}

$wsText = ReadEntry $match.Part
$zip.Dispose()
if($src -ne $Path){ Remove-Item $src -Force }

# --- volcar filas -------------------------------------------------------------
$doc = New-Object System.Xml.XmlDocument; $doc.LoadXml($wsText)
$nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$nsm.AddNamespace("d","http://schemas.openxmlformats.org/spreadsheetml/2006/main")

$cols = @()
for($i=1; $i -le $MaxCol; $i++){
  $n=""; $k=$i; while($k -gt 0){ $r=($k-1)%26; $n=[char](65+$r)+$n; $k=[int](($k-$r)/26) }
  $cols += $n
}

$sb = New-Object System.Text.StringBuilder
foreach($row in $doc.SelectNodes("//d:sheetData/d:row",$nsm)){
  $cells = @{}
  foreach($c in $row.SelectNodes("d:c",$nsm)){
    $col = ($c.GetAttribute("r") -replace '[0-9]','')
    $t = $c.GetAttribute("t")
    $v = ""
    if($t -eq "inlineStr"){ $v = $c.InnerText }
    else {
      $vn = $c.SelectSingleNode("d:v",$nsm)
      if($vn){ $v = $vn.InnerText; if($t -eq "s"){ $v = $ss[[int]$v] } }
    }
    # Un TSV con tabs o saltos incrustados se desalinea al parsearlo despues.
    $v = $v -replace "`t"," " -replace "`r`n"," " -replace "`n"," "
    $cells[$col] = $v
  }
  $line = $row.GetAttribute("r")
  foreach($cc in $cols){ $line += "`t" + $(if($cells.ContainsKey($cc)){ $cells[$cc] } else { "" }) }
  [void]$sb.AppendLine($line)
}

if($Out -eq ""){ Write-Output $sb.ToString() }
else {
  [System.IO.File]::WriteAllText($Out, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
  Write-Output ("Hoja '" + $match.Name + "' -> " + $Out)
}
