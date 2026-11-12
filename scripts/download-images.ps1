<#
  Descarga los retratos de los campeones desde MCOCHUB (mcochub.insaneskull.com),
  los reduce a 128 px y los guarda en img\. Genera data\images.js con el
  mapa  nombre -> archivo  que usa la app (los que falten muestran las iniciales).

  Uso:  powershell -ExecutionPolicy Bypass -File scripts\download-images.ps1 [-Force] [-All]

  Por defecto SOLO descarga los campeones nuevos (los que no están en data\images.js): los retratos
  antiguos ya están en GitHub y la carpeta img\ local puede estar vacía para ahorrar espacio.
  -All  vuelve a descargar todos los que falten en img\ (comportamiento antiguo).
  -Force  sobrescribe aunque el archivo exista.

  Los retratos son propiedad de Kabam / Marvel: úsalos solo para uso personal.
  Si un campeón no tiene imagen, la app muestra sus iniciales.
#>
param([switch]$Force, [switch]$All, [int]$Size = 128)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Drawing

$root = Resolve-Path (Join-Path $PSScriptRoot '..')
$imgDir = Join-Path $root 'img'
New-Item -ItemType Directory -Force $imgDir | Out-Null
$dataJson = Join-Path $root 'data\data.json'
if (-not (Test-Path $dataJson)) { throw 'Ejecuta antes scripts\excel-to-json.ps1' }

$ua = 'Mozilla/5.0 (fan app MCOC; descarga puntual de retratos)'
$hub = 'https://mcochub.insaneskull.com'

# Nombres que no coinciden con MCOCHUB
$alias = @{
  'Human Torch (Victorious)'      = 'human-torch'
  'Invisible Woman (Victorious)'  = 'invisible-woman'
  'Mister Fantastic (Victorious)' = 'mister-fantastic'
  'Thing (Victorious)'            = 'thing'
  'Platinumpool'                  = 'platinum-pool'
  'Yellowjacket'                  = 'yellowjacket'
  'Daredevil'                     = 'daredevil-classic'
}

Write-Host 'Leyendo lista de campeones de MCOCHUB...'
$list = Invoke-WebRequest -UseBasicParsing "$hub/champions" -UserAgent $ua -TimeoutSec 60
$hubMap = @{}
foreach ($m in [regex]::Matches($list.Content, 'champions/([a-z0-9\-]+)"[^>]*>\s*([^<]+?)\s*</a>')) {
  $hubMap[[System.Net.WebUtility]::HtmlDecode($m.Groups[2].Value)] = $m.Groups[1].Value
}

$data = Get-Content $dataJson -Raw -Encoding UTF8 | ConvertFrom-Json
$rows = $data.sheets.Campeones.rows
$links = $data.sheets.Campeones.links

function Get-FileSlug([string]$s) {
  $n = $s.Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', ''
  $n = $n.ToLowerInvariant() -replace "['’.()]", '' -replace '[^a-z0-9]+', '-'
  return $n.Trim('-')
}

$manifest = [ordered]@{}
# Retratos ya publicados (data\images.js): se conservan en el mapa aunque no estén en img\ local
$known = @{}
$imgJs = Join-Path $root 'data\images.js'
if (Test-Path $imgJs) {
  $prev = ([IO.File]::ReadAllText($imgJs, [Text.Encoding]::UTF8) -replace '^self.MCOC_IMG=', '' -replace ';\s*$', '') | ConvertFrom-Json
  foreach ($p in $prev.PSObject.Properties) { $known[$p.Name] = $p.Value }
}
$ok = 0; $fail = @(); $new = @()
for ($i = 1; $i -lt $rows.Count; $i++) {
  $name = $rows[$i][0]
  if (-not $name) { continue }
  if (-not $All -and -not $Force -and $known.ContainsKey($name)) { $manifest[$name] = $known[$name]; continue }
  $new += $name
  $slug = $null
  if ($alias.ContainsKey($name)) { $slug = $alias[$name] }
  elseif ($links."$i,17") { $slug = ($links."$i,17" -split '/')[-1] }
  elseif ($hubMap.ContainsKey($name)) { $slug = $hubMap[$name] }
  if (-not $slug) { $fail += $name; continue }

  $file = (Get-FileSlug $name) + '.png'
  $dest = Join-Path $imgDir $file
  if ((Test-Path $dest) -and -not $Force) { $manifest[$name] = $file; $ok++; continue }
  $tmp = [IO.Path]::GetTempFileName()
  try {
    try {
      Invoke-WebRequest -UseBasicParsing "$hub/storage/champs/$slug.png" -OutFile $tmp -UserAgent $ua -TimeoutSec 30
    } catch {
      # El nombre del archivo no siempre coincide con el slug: se busca en la ficha.
      $page = (Invoke-WebRequest -UseBasicParsing "$hub/champions/$slug" -UserAgent $ua -TimeoutSec 30).Content
      $hubName = ''
      $mm = [regex]::Match($page, '"about":\{"@type":"Thing","name":"([^"]+)"')
      if ($mm.Success) { $hubName = $mm.Groups[1].Value }
      $url = $null
      foreach ($im in [regex]::Matches($page, 'src="([^"]+/storage/champs/[^"]+\.png)" alt="([^"]*)"')) {
        if ([System.Net.WebUtility]::HtmlDecode($im.Groups[2].Value) -eq $hubName) { $url = $im.Groups[1].Value; break }
      }
      if (-not $url) {
        $og = [regex]::Match($page, 'property="og:image" content="([^"]+\.png)"')
        if ($og.Success -and $og.Groups[1].Value -notmatch '/og/') { $url = $og.Groups[1].Value }
      }
      if (-not $url) { throw 'sin imagen' }
      Start-Sleep -Milliseconds 150
      Invoke-WebRequest -UseBasicParsing $url -OutFile $tmp -UserAgent $ua -TimeoutSec 30
    }
    $src = [System.Drawing.Image]::FromFile($tmp)
    $bmp = New-Object System.Drawing.Bitmap($Size, $Size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.DrawImage($src, 0, 0, $Size, $Size)
    $g.Dispose(); $src.Dispose()
    $bmp.Save($dest, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    $manifest[$name] = $file; $ok++
    Start-Sleep -Milliseconds 150
  } catch {
    $fail += "$name ($slug)"
  } finally { Remove-Item $tmp -ErrorAction SilentlyContinue }
  if ($ok % 25 -eq 0) { Write-Host "  $ok imagenes..." }
}

$json = $manifest | ConvertTo-Json -Compress
[IO.File]::WriteAllText((Join-Path $root 'data\images.js'), "self.MCOC_IMG=$json;", (New-Object System.Text.UTF8Encoding($false)))
Write-Host "OK: $ok imagenes descargadas/comprobadas ($($manifest.Count) en el mapa). Sin imagen: $($fail.Count)"
if ($new.Count) { Write-Host "Campeones nuevos procesados: $($new -join ', ')" } else { Write-Host 'No hay campeones nuevos: no se ha descargado nada.' }
if ($fail.Count) { $fail | ForEach-Object { Write-Host "  - $_" } }
