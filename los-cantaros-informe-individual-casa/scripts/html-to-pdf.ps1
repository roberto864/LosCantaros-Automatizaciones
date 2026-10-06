# html-to-pdf.ps1 — imprime un HTML local a PDF con Edge o Chrome headless.
#
# Uso: .\html-to-pdf.ps1 -Html informe.html -Out informe.pdf
#
# Detalles que importan y que cuestan encontrar:
#  * hace falta un -user-data-dir propio, si no Edge se cuelga o no escribe nada
#    cuando ya hay una ventana abierta con el perfil del usuario;
#  * --virtual-time-budget le da tiempo a bajar las fuentes de Google antes de
#    imprimir, si no salen con la tipografia de respaldo;
#  * el HTML debe traer un @media print que fuerce el tema claro; el PDF hereda
#    prefers-color-scheme del sistema y en modo oscuro sale ilegible.
param(
  [Parameter(Mandatory=$true)][string]$Html,
  [Parameter(Mandatory=$true)][string]$Out,
  [int]$WaitMs = 8000
)
$ErrorActionPreference = "Stop"

$browser = $null
$candidates = @(
  "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
  "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
  "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
  "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
  "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
)
foreach($c in $candidates){ if($c -and (Test-Path $c)){ $browser = $c; break } }
if(-not $browser){ throw "No se encontro Edge ni Chrome para generar el PDF." }

$htmlFull = (Resolve-Path $Html).Path
$url = "file:///" + ($htmlFull -replace '\\','/')
$profile = Join-Path ([System.IO.Path]::GetTempPath()) ("pdfprof_" + [Guid]::NewGuid().ToString("N"))

if(Test-Path $Out){ Remove-Item $Out -Force }
# Ruta absoluta: el navegador no hereda el directorio de trabajo de PowerShell.
$outFull = if([System.IO.Path]::IsPathRooted($Out)){ $Out }
           else { [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $Out)) }

# Las rutas con espacios ("Informe Pagos Casa 25.pdf") hay que entrecomillarlas
# a mano. Start-Process -ArgumentList las parte en dos y el navegador cree que
# le pasaron varias URLs: muere con "Multiple targets are not supported".
# ProcessStartInfo.Arguments con comillas explicitas es la forma que aguanta.
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $browser
$psi.Arguments = @(
  "--headless=new","--disable-gpu","--no-sandbox",
  "--user-data-dir=`"$profile`"",
  "--virtual-time-budget=$WaitMs",
  "--no-pdf-header-footer",
  "--print-to-pdf=`"$outFull`"",
  "`"$url`""
) -join " "
$psi.UseShellExecute = $false
$psi.CreateNoWindow = $true
$p = [System.Diagnostics.Process]::Start($psi)
$p.WaitForExit()
Remove-Item $profile -Recurse -Force -ErrorAction SilentlyContinue

if(-not (Test-Path $outFull)){ throw "El navegador no genero el PDF (exit $($p.ExitCode))." }
$Out = $outFull
Write-Output ("PDF generado: $Out (" + [int]((Get-Item $Out).Length/1024) + " KB)")
