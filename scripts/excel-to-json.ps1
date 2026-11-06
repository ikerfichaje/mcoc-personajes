<#
  Convierte MCOC_campeones_nov2026.xlsx en los datos de la app (sin dependencias:
  solo PowerShell 5.1+ en Windows).

  Uso (desde la carpeta del proyecto):
    powershell -ExecutionPolicy Bypass -File scripts\excel-to-json.ps1
    powershell -ExecutionPolicy Bypass -File scripts\excel-to-json.ps1 -Excel otro.xlsx

  Genera:
    data\data.json  -> datos crudos por hoja (cabecera + filas + enlaces)
    data\data.js   -> lo mismo como  window.MCOC_RAW = {...}  (funciona abriendo
                           index.html con doble clic, sin servidor)
#>
param(
  [string]$Excel = (Join-Path $PSScriptRoot '..\MCOC_campeones_nov2026.xlsx'),
  [string]$OutDir = (Join-Path $PSScriptRoot '..\data')
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$Excel = (Resolve-Path $Excel).Path
New-Item -ItemType Directory -Force $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path

$zip = [System.IO.Compression.ZipFile]::OpenRead($Excel)
function Read-Entry([string]$name) {
  $e = $zip.GetEntry($name)
  if (-not $e) { return $null }
  $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
  try { $txt = $sr.ReadToEnd() } finally { $sr.Close() }
  $x = New-Object System.Xml.XmlDocument
  $x.LoadXml($txt)
  return $x
}

$nsMain = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$nsRel  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$nsPkg  = 'http://schemas.openxmlformats.org/package/2006/relationships'

# --- Textos compartidos ---
$shared = New-Object System.Collections.Generic.List[string]
$ss = Read-Entry 'xl/sharedStrings.xml'
if ($ss) {
  foreach ($si in $ss.DocumentElement.ChildNodes) {
    if ($si.LocalName -ne 'si') { continue }
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in $si.GetElementsByTagName('t', $nsMain)) {
      # ignorar textos fonéticos (rPh)
      if ($t.ParentNode.LocalName -eq 'rPh') { continue }
      [void]$sb.Append($t.InnerText)
    }
    $shared.Add($sb.ToString())
  }
}

function ColIndex([string]$ref) {
  $letters = ($ref -replace '[0-9]', '')
  $n = 0
  foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int][char]$ch - 64) }
  return $n - 1
}

# --- Hojas ---
$wb = Read-Entry 'xl/workbook.xml'
$wbRels = Read-Entry 'xl/_rels/workbook.xml.rels'
$relMap = @{}
foreach ($r in $wbRels.DocumentElement.ChildNodes) { $relMap[$r.GetAttribute('Id')] = $r.GetAttribute('Target') }

$result = [ordered]@{ generated = (Get-Date).ToString('yyyy-MM-dd'); source = [IO.Path]::GetFileName($Excel); sheets = [ordered]@{} }

foreach ($s in $wb.GetElementsByTagName('sheet', $nsMain)) {
  $name = $s.GetAttribute('name')
  $rid = $s.GetAttribute('id', $nsRel)
  $target = $relMap[$rid]
  $path = 'xl/' + ($target -replace '^/xl/', '' -replace '^/', '')
  $sheet = Read-Entry $path

  # enlaces
  $links = @{}
  $relsPath = ($path -replace 'worksheets/', 'worksheets/_rels/') + '.rels'
  $sRels = Read-Entry $relsPath
  $sRelMap = @{}
  if ($sRels) { foreach ($r in $sRels.DocumentElement.ChildNodes) { $sRelMap[$r.GetAttribute('Id')] = $r.GetAttribute('Target') } }
  foreach ($h in $sheet.GetElementsByTagName('hyperlink', $nsMain)) {
    $hid = $h.GetAttribute('id', $nsRel)
    if ($hid -and $sRelMap.ContainsKey($hid)) { $links[$h.GetAttribute('ref')] = $sRelMap[$hid] }
  }

  $rows = New-Object System.Collections.Generic.List[object]
  $linkOut = [ordered]@{}
  foreach ($row in $sheet.GetElementsByTagName('row', $nsMain)) {
    $cells = @{}
    $max = -1
    foreach ($c in $row.ChildNodes) {
      if ($c.LocalName -ne 'c') { continue }
      $ref = $c.GetAttribute('r')
      $ci = ColIndex $ref
      $t = $c.GetAttribute('t')
      $val = ''
      if ($t -eq 'inlineStr') {
        $val = ($c.GetElementsByTagName('t', $nsMain) | ForEach-Object { $_.InnerText }) -join ''
      } else {
        $v = $c.GetElementsByTagName('v', $nsMain)
        if ($v.Count -gt 0) {
          $val = $v[0].InnerText
          if ($t -eq 's') { $val = $shared[[int]$val] }
        }
      }
      if ($val -ne '') { $cells[$ci] = $val; if ($ci -gt $max) { $max = $ci } }
      if ($links.ContainsKey($ref)) { $linkOut["$($rows.Count),$ci"] = $links[$ref] }
    }
    $arr = @()
    for ($i = 0; $i -le $max; $i++) { if ($cells.ContainsKey($i)) { $arr += [string]$cells[$i] } else { $arr += '' } }
    $rows.Add($arr)
  }
  $result.sheets[$name] = [ordered]@{ rows = $rows; links = $linkOut }
  Write-Host ("{0,-26} {1,5} filas  {2,4} enlaces" -f $name, $rows.Count, $linkOut.Count)
}
$zip.Dispose()

$json = $result | ConvertTo-Json -Depth 6 -Compress
$utf8 = New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText((Join-Path $OutDir 'data.json'), $json, $utf8)
[IO.File]::WriteAllText((Join-Path $OutDir 'data.js'), "window.MCOC_RAW=$json;", $utf8)
Write-Host "OK -> $OutDir"
