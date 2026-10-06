# write-xlsx.ps1 — construye un .xlsx nuevo y autocontenido desde un TSV.
# No abre ni toca planillas existentes: escribe las partes XML y las comprime.
# Uso: .\write-xlsx.ps1 -Tsv datos.tsv -Out informe.xlsx [-SheetName "Casa 25"] [-Title "..."]
param(
  [Parameter(Mandatory=$true)][string]$Tsv,
  [Parameter(Mandatory=$true)][string]$Out,
  [string]$SheetName = "Detalle",
  [string]$Title = ""
)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Esc([string]$s){ [System.Security.SecurityElement]::Escape($s) }
function ColName([int]$i){  # 1 -> A
  $n=""; while($i -gt 0){ $r=($i-1)%26; $n=[char](65+$r)+$n; $i=[int](($i-$r)/26) }; return $n
}

$lines = [System.IO.File]::ReadAllLines($Tsv, [System.Text.Encoding]::UTF8)
$lines = $lines | Where-Object { $_ -ne $null }
if($lines.Count -eq 0){ throw "TSV vacio: $Tsv" }

$rowsXml = New-Object System.Text.StringBuilder
$rowIdx = 0
$maxCols = 0

if($Title -ne ""){
  $rowIdx++
  [void]$rowsXml.Append("<row r=`"$rowIdx`" ht=`"20`" customHeight=`"1`"><c r=`"A$rowIdx`" s=`"3`" t=`"inlineStr`"><is><t>$(Esc $Title)</t></is></c></row>")
  $rowIdx++   # fila en blanco
}

$headerDone = $false
foreach($line in $lines){
  $rowIdx++
  $cells = $line -split "`t"
  if($cells.Count -gt $maxCols){ $maxCols = $cells.Count }
  $style = 0
  if(-not $headerDone){ $style = 1; $headerDone = $true }
  [void]$rowsXml.Append("<row r=`"$rowIdx`">")
  for($c=0; $c -lt $cells.Count; $c++){
    $v = $cells[$c]
    if($v -eq ""){ continue }
    $ref = (ColName ($c+1)) + $rowIdx
    if($style -eq 0 -and $v -match '^-?\d+$'){
      [void]$rowsXml.Append("<c r=`"$ref`" s=`"2`"><v>$v</v></c>")
    } else {
      [void]$rowsXml.Append("<c r=`"$ref`" s=`"$style`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(Esc $v)</t></is></c>")
    }
  }
  [void]$rowsXml.Append("</row>")
}

# Anchos de columna razonables: la primera angosta (fecha), la segunda ancha (descripcion), el resto medio.
$colsXml = "<cols>"
for($i=1; $i -le [Math]::Max($maxCols,1); $i++){
  $w = 16
  if($i -eq 2){ $w = 42 }
  if($i -ge 3 -and $i -le 5){ $w = 13 }
  if($i -ge 6){ $w = 46 }
  $colsXml += "<col min=`"$i`" max=`"$i`" width=`"$w`" customWidth=`"1`"/>"
}
$colsXml += "</cols>"

$freeze = if($Title -ne ""){ "A4" } else { "A2" }
$sheet = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="$(if($Title -ne ''){3}else{1})" topLeftCell="$freeze" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>$colsXml<sheetData>$($rowsXml.ToString())</sheetData></worksheet>
"@

$styles = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0"/></numFmts>
<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="14"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF14232C"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" vertical="top"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"@

$workbook = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="$(Esc $SheetName)" sheetId="1" r:id="rId1"/></sheets></workbook>
"@

$wbRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
"@

$rootRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
"@

$contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>
"@

# El paquete OPC exige rutas con "/" — CreateFromDirectory las escribe con "\" en
# .NET Framework y Excel rechaza el archivo, asi que armamos las entradas a mano.
$parts = [ordered]@{
  "[Content_Types].xml"        = $contentTypes.Trim()
  "_rels/.rels"                = $rootRels.Trim()
  "xl/workbook.xml"            = $workbook.Trim()
  "xl/_rels/workbook.xml.rels" = $wbRels.Trim()
  "xl/styles.xml"              = $styles.Trim()
  "xl/worksheets/sheet1.xml"   = $sheet.Trim()
}
if(Test-Path $Out){ Remove-Item $Out -Force }
$enc = New-Object System.Text.UTF8Encoding($false)
$fs = [System.IO.File]::Open($Out, [System.IO.FileMode]::Create)
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
foreach($k in $parts.Keys){
  $entry = $zip.CreateEntry($k, [System.IO.Compression.CompressionLevel]::Optimal)
  $st = $entry.Open()
  $bytes = $enc.GetBytes($parts[$k])
  $st.Write($bytes, 0, $bytes.Length)
  $st.Close()
}
$zip.Dispose(); $fs.Close()
Write-Output "XLSX escrito: $Out ($((Get-Item $Out).Length) bytes, $rowIdx filas)"
