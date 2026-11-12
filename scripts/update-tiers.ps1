<#
  Actualiza los tiers del Excel comparando varias tier lists y calcula una TIER LIST GENERAL
  (todas las clases juntas) por consenso.

  Fuentes (todas valoran el uso en ATAQUE salvo las columnas de defensa):
    - mcoc.app: nota individual (0-100) de cada youtuber -> Vega, Lagacy, MetalSonicDude, Seatin,
      y su media (columna «Tier youtubers»).
    - MCOC Guide Blog (CaptainRaghu): tier de ataque (S+..E) y defensa (S..D).
    - MCOC Scout: tier de ataque (A++, S, A..E) y defensa (S..D) de cada ficha. Es una COPIA del Guide Blog
      (A++ = S+), así que solo cuenta para los campeones que el Guide Blog no valora (Deathless, Stellar-Forged…).
  Descartadas tras revisarlas (ver CLAUDE.md §10): gacha.gg, PocketGamer, MMO Culture.

  Método del consenso:
    1. Cada opinión se pasa a una nota 0-100. Los youtubers ya la traen de mcoc.app. Las letras
       se convierten al centro de la franja que usa mcoc.app para esa letra (así todo usa la misma escala).
    2. Nota general = media de las opiniones disponibles. Tier general = franja de mcoc.app de esa nota.
    3. Acuerdo = diferencia entre la nota más alta y la más baja: <=20 Alto, <=40 Medio, >40 Bajo.

  Uso (desde la carpeta del proyecto, con el Excel CERRADO):
    powershell -ExecutionPolicy Bypass -File scripts\update-tiers.ps1            # descarga y actualiza
    powershell -ExecutionPolicy Bypass -File scripts\update-tiers.ps1 -DryRun    # solo muestra cambios
  Después: scripts\excel-to-json.ps1 y subir VERSION en sw.js.
  Las páginas descargadas se guardan en %TEMP%\mcoc-tiers\<fecha> (se reutilizan si repites el mismo día).
#>
param(
  [string]$Excel = (Join-Path $PSScriptRoot '..\MCOC_campeones_nov2026.xlsx'),
  [string]$Cache = (Join-Path $env:TEMP ('mcoc-tiers\' + (Get-Date -Format 'yyyy-MM-dd'))),
  [switch]$DryRun
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$Excel = (Resolve-Path $Excel).Path
$D = [string][char]0x2014
$ua = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/130 Safari/537.36'
New-Item -ItemType Directory -Force $Cache, "$Cache\app", "$Cache\scout" | Out-Null

function Get-Page([string]$url, [string]$file) {
  if (-not (Test-Path $file)) {
    Invoke-WebRequest -UseBasicParsing $url -UserAgent $ua -TimeoutSec 60 -OutFile $file
    Start-Sleep -Milliseconds 250
  }
  return [IO.File]::ReadAllText($file, [Text.Encoding]::UTF8)
}
function Dec([string]$s) { [Net.WebUtility]::HtmlDecode($s).Trim() }

# ---------- nombres del Excel ----------
Write-Host 'Leyendo el Excel...'
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$wb = $x.Workbooks.Open($Excel)
try {
  $wsC = $wb.Worksheets.Item('Campeones')
  $last = $wsC.Cells.Item($wsC.Rows.Count, 1).End(-4162).Row
  $vals = $wsC.Range("A1:F$last").Value2
  $champs = [ordered]@{}
  for ($r = 2; $r -le $last; $r++) {
    $n = [string]$vals[$r, 1]; if (-not $n) { continue }
    $champs[$n] = [pscustomobject]@{ name = $n; cls = [string]$vals[$r, 2]; rar = [string]$vals[$r, 3]; row = $r
      ytOld = [string]$vals[$r, 4]; atkOld = [string]$vals[$r, 5]; defOld = [string]$vals[$r, 6]
      yt = $null; ytScore = $null; atk = $null; def = $null; scAtk = $null; scDef = $null; cr = @{} }
  }
  function Norm([string]$s) { ($s.Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', '').ToLowerInvariant() -replace '[^a-z0-9]', '' }
  $byNorm = @{}; foreach ($k in $champs.Keys) { $byNorm[(Norm $k)] = $k }
  $alias = @{
    'Dr. Bong' = 'Doctor Bong'; 'Stark Spidey' = 'Spider-Man (Stark Enhanced)'; 'Immortal Abomination' = 'Abomination (Immortal)'
    'Immortal Hulk' = 'Hulk (Immortal)'; 'Joaquin Torres (Falcon)' = 'Falcon (Joaquín Torres)'; 'Spiderman 2099' = 'Spider-Man 2099'
    'MODOK' = 'M.O.D.O.K.'; 'Captain America (Classic)' = 'Captain America'; 'Black Panther (Classic)' = 'Black Panther'
    'Daredevil (Classic)' = 'Daredevil'; 'Magneto OG (Red Suit)' = 'Magneto'; 'Magneto (Marvel Now)' = 'Magneto (House of X)'
    'Negasonic Warhead' = 'Negasonic Teenage Warhead'; 'X-23' = 'Wolverine (X-23)'; 'Hulk Buster' = 'Hulkbuster'
    'Infamous Iron Man' = 'Iron Man (Infamous)'; 'Iron Man (IW)' = 'Iron Man (Infinity War)'; 'Storm Pyramid X' = 'Storm (Pyramid X)'
    'Spider-Man Symbiote' = 'Spider-Man (Symbiote)'; 'Kamala Khan' = 'Ms. Marvel (Kamala Khan)'; 'Black Widow (CV)' = 'Black Widow (Claire Voyant)'
    'Spiderman Supreme' = 'Spider-Man (Supreme)'; 'Spider-Woman' = 'Spider-Woman (Jessica Drew)'; 'Captain Marvel' = 'Captain Marvel (Movie)'
    'Carina' = 'Carina Tivan'; 'Scarlet Witch' = 'Scarlet Witch (Sigil)'; 'Spider-Man' = 'Spider-Man (Classic)'
    'Blade (Stellar Forge)' = 'Blade (Stellar Forged)'; 'Platinum Pool' = 'Platinumpool'
  }
  function Resolve([string]$n) {
    if ($alias.ContainsKey($n)) { return $alias[$n] }
    if ($champs.Contains($n)) { return $n }
    $k = Norm $n; if ($byNorm.ContainsKey($k)) { return $byNorm[$k] }
    return $null
  }
  $unres = New-Object System.Collections.Generic.List[string]

  # ---------- mcoc.app ----------
  Write-Host 'mcoc.app: lista y fichas (unos minutos la primera vez)...'
  $list = Get-Page 'https://mcoc.app/champions' "$Cache\mcocapp-champions.html"
  $slugs = [regex]::Matches($list, 'href="/champions/([^"/]+)"') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
  $bands = @{}   # tier -> scores (para deducir las franjas de mcoc.app)
  $creators = 'Vega', 'Lagacy', 'MetalSonicDude', 'Seatin'
  foreach ($s in $slugs) {
    $h = Get-Page "https://mcoc.app/champions/$s" "$Cache\app\$s.html"
    $nm = Dec ([regex]::Match($h, '<title>(.+?) MCOC Rating').Groups[1].Value)
    if (-not $nm) { $nm = Dec ([regex]::Match($h, '<h1[^>]*>([^<]+)</h1>').Groups[1].Value) }
    $k = Resolve $nm
    if (-not $k) { $unres.Add("mcoc.app: $nm"); continue }
    $m = [regex]::Match($h, '(?s)rate [^<]*?<b>(S\+|S|A|B|C|D|F) tier</b> with a score of\s*<b>(\d+)/100')
    if (-not $m.Success) { $m = [regex]::Match($h, '(?s)\n(S\+|S|A|B|C|D|F)\s+tier\s+with a score of\s+(\d+)/100') }
    if (-not $m.Success) { $m = [regex]::Match(([regex]::Replace($h, '<[^>]+>', '')), '(S\+|S|A|B|C|D|F)\s+tier\s+with a score of\s+(\d+)/100') }
    if ($m.Success) {
      $champs[$k].yt = $m.Groups[1].Value; $champs[$k].ytScore = [int]$m.Groups[2].Value
      if (-not $bands.ContainsKey($champs[$k].yt)) { $bands[$champs[$k].yt] = New-Object System.Collections.Generic.List[int] }
      $bands[$champs[$k].yt].Add($champs[$k].ytScore)
    }
    foreach ($cm in [regex]::Matches($h, '<a href="/creators/[^"]+">([^<]+)</a>(?:(?!</li>).)*?<b class="[^"]*">(\d+)</b>')) {
      $champs[$k].cr[$cm.Groups[1].Value] = [int]$cm.Groups[2].Value
    }
  }
  # Franjas de mcoc.app: mínimo de cada tier
  $order = 'S+', 'S', 'A', 'B', 'C', 'D', 'F'
  $minOf = [ordered]@{}; foreach ($t in $order) { if ($bands.ContainsKey($t)) { $minOf[$t] = ($bands[$t] | Measure-Object -Minimum).Minimum } }
  $minOf['F'] = 0
  Write-Host ('Franjas mcoc.app (nota mínima): ' + (($minOf.Keys | ForEach-Object { "$_>=$($minOf[$_])" }) -join '  '))
  function TierOf([double]$v) { foreach ($t in $order) { if ($minOf.Contains($t) -and $v -ge $minOf[$t]) { return $t } }; return 'F' }
  # centro de cada franja
  $mid = @{}; $prevTop = 100
  foreach ($t in $order) { if (-not $minOf.Contains($t)) { continue }; $mid[$t] = [math]::Round(($minOf[$t] + $prevTop) / 2); $prevTop = $minOf[$t] - 1 }

  # ---------- MCOC Guide Blog ----------
  Write-Host 'MCOC Guide Blog...'
  $g = Get-Page 'https://www.mcocguideblog.com/p/mcoc-offense-and-defense-tier-list.html' "$Cache\guideblog.html"
  $gbDate = Dec ([regex]::Match($g, "updated' datetime='[^']+'>([^<]+)<").Groups[1].Value)
  try { $gbDate = ([datetime]::Parse($gbDate, [Globalization.CultureInfo]'en-US')).ToString('dd/MM/yyyy') } catch { $gbDate = '' }
  $panels = [regex]::Matches($g, '(?s)<div class="panel[^"]*" data-c="([^"]+)" data-m="(atk|def)">(.*?)(?=<div class="panel|</section|<script)')
  foreach ($p in $panels) {
    foreach ($row in [regex]::Matches($p.Groups[3].Value, '(?s)<span class="key">([^<]+)</span>.*?<div class="chips">(.*?)</div>')) {
      foreach ($c in [regex]::Matches($row.Groups[2].Value, '<span class="chip">([^<]+)<')) {
        $n = Dec $c.Groups[1].Value; $k = Resolve $n
        if (-not $k) { $unres.Add("Guide Blog: $n"); continue }
        if ($p.Groups[2].Value -eq 'atk') { $champs[$k].atk = Dec $row.Groups[1].Value } else { $champs[$k].def = Dec $row.Groups[1].Value }
      }
    }
  }

  # ---------- MCOC Scout ----------
  Write-Host 'MCOC Scout: fichas...'
  $rel = Get-Page 'https://mcocscout.com/relics' "$Cache\scout-relics.html"
  $sslugs = [regex]::Matches($rel, 'href="/champions/([^"/]+)"') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
  foreach ($s in $sslugs) {
    try { $h = Get-Page "https://mcocscout.com/champions/$s" "$Cache\scout\$s.html" } catch { $unres.Add("Scout (descarga): $s"); continue }
    $nm = Dec ([regex]::Match($h, '<h1[^>]*>([^<]+)</h1>').Groups[1].Value)
    $k = Resolve $nm
    if (-not $k) { $unres.Add("Scout: $nm"); continue }
    $a = [regex]::Match($h, 'is rated (\S+) tier for Attack'); if ($a.Success) { $champs[$k].scAtk = $a.Groups[1].Value }
    $d2 = [regex]::Match($h, 'is rated (\S+) tier for Defense'); if ($d2.Success) { $champs[$k].scDef = $d2.Groups[1].Value }
  }

  # ---------- consenso ----------
  $gbAtkMap = @{ 'S+' = 'S+'; 'S' = 'S'; 'A' = 'A'; 'B' = 'B'; 'C' = 'C'; 'D' = 'D'; 'E' = 'F' }
  # MCOC Scout copia el Guide Blog (llama «A++» a su S+): solo se usa cuando el Guide Blog no valora al campeón.
  $scAtkMap = @{ 'A++' = 'S+'; 'S' = 'S'; 'A' = 'A'; 'B' = 'B'; 'C' = 'C'; 'D' = 'D'; 'E' = 'F' }
  $defMap = @{ 'S' = 'S'; 'A' = 'A'; 'B' = 'B'; 'C' = 'C'; 'D' = 'D' }
  function Score($t, $map) { if ($t -and $map.ContainsKey($t)) { return $mid[$map[$t]] }; return $null }
  function Agree($arr) { if ($arr.Count -lt 2) { return $D }; $sp = ($arr | Measure-Object -Maximum).Maximum - ($arr | Measure-Object -Minimum).Minimum; if ($sp -le 20) { 'Alto' } elseif ($sp -le 40) { 'Medio' } else { 'Bajo' } }
  foreach ($c in $champs.Values) {
    $ops = New-Object System.Collections.Generic.List[double]
    foreach ($cn in $creators) { if ($c.cr.ContainsKey($cn)) { $ops.Add($c.cr[$cn]) } }
    $v = Score $c.atk $gbAtkMap
    if ($null -eq $v) { $v = Score $c.scAtk $scAtkMap }
    if ($null -ne $v) { $ops.Add($v) }
    $c | Add-Member -NotePropertyName nOps -NotePropertyValue $ops.Count
    if ($ops.Count) {
      $avg = ($ops | Measure-Object -Average).Average
      $c | Add-Member score ([math]::Round($avg)); $c | Add-Member gen (TierOf $avg); $c | Add-Member agree (Agree $ops)
    } else { $c | Add-Member score $null; $c | Add-Member gen $D; $c | Add-Member agree $D }
    $dv = Score $c.def $defMap; if ($null -eq $dv) { $dv = Score $c.scDef $defMap }
    if ($null -ne $dv) { $c | Add-Member genDef (TierOf $dv) } else { $c | Add-Member genDef $D }
  }

  # ---------- informe ----------
  if ($unres.Count) { Write-Host "Sin emparejar ($($unres.Count)):"; $unres | Sort-Object -Unique | ForEach-Object { Write-Host "  - $_" } }
  $chg = @()
  foreach ($c in $champs.Values) {
    $ny = if ($c.yt) { $c.yt } else { $D }; $na = if ($c.atk) { $c.atk } else { $D }; $nd = if ($c.def) { $c.def } else { $D }
    if ($ny -ne $c.ytOld -or $na -ne $c.atkOld -or $nd -ne $c.defOld) { $chg += "  $($c.name): $($c.ytOld) | $($c.atkOld)/$($c.defOld)  ->  $ny | $na/$nd" }
    $c | Add-Member ytNew $ny; $c | Add-Member atkNew $na; $c | Add-Member defNew $nd
  }
  Write-Host "Cambios en Campeones (youtubers | GB ataque/defensa): $($chg.Count)"; $chg | ForEach-Object { Write-Host $_ }
  $sorted = $champs.Values | Sort-Object @{ e = { if ($null -eq $_.score) { -1 } else { $_.score } }; Descending = $true }, @{ e = { $_.nOps }; Descending = $true }, name
  Write-Host 'Top 15 general:'; $sorted | Select-Object -First 15 | ForEach-Object { Write-Host ("  {0,-34} {1,-3} {2,3}  ({3} opiniones, acuerdo {4})" -f $_.name, $_.gen, $_.score, $_.nOps, $_.agree) }
  if ($DryRun) { Write-Host 'DryRun: no se ha modificado el Excel.'; return }

  # ---------- escribir en el Excel ----------
  foreach ($c in $champs.Values) {
    $wsC.Cells.Item($c.row, 4).Value2 = $c.ytNew; $wsC.Cells.Item($c.row, 5).Value2 = $c.atkNew; $wsC.Cells.Item($c.row, 6).Value2 = $c.defNew
  }
  $wsR = $wb.Worksheets.Item('Reliquias por campeón')
  $lastR = $wsR.Cells.Item($wsR.Rows.Count, 1).End(-4162).Row
  for ($r = 2; $r -le $lastR; $r++) {
    $n = $wsR.Cells.Item($r, 1).Text
    if ($champs.Contains($n)) { $c = $champs[$n]; $wsR.Cells.Item($r, 4).Value2 = $c.ytNew; $wsR.Cells.Item($r, 5).Value2 = $c.atkNew; $wsR.Cells.Item($r, 6).Value2 = $c.defNew }
  }

  $sheetName = 'Tier list general'
  $ws = $null; foreach ($s in $wb.Worksheets) { if ($s.Name -eq $sheetName) { $ws = $s } }
  if ($ws) { $ws.Cells.Clear() | Out-Null } else { $ws = $wb.Worksheets.Add([Type]::Missing, $wsC); $ws.Name = $sheetName }
  $hdr = @('Puesto', 'Campeón', 'Clase', 'Rareza máx.', 'Tier general', 'Nota general (0-100)', 'Nº opiniones', 'Acuerdo',
    'Vega', 'Lagacy', 'MetalSonicDude', 'Seatin', 'Media youtubers (mcoc.app)', 'Guide Blog ataque', 'MCOC Scout ataque',
    'Tier general defensa', 'Guide Blog defensa', 'MCOC Scout defensa')
  $rows = @($sorted)
  $arr = New-Object 'object[,]' ($rows.Count + 1), $hdr.Count
  for ($j = 0; $j -lt $hdr.Count; $j++) { $arr[0, $j] = $hdr[$j] }
  $pos = 0
  for ($i = 0; $i -lt $rows.Count; $i++) {
    $c = $rows[$i]; $v = @()
    if ($null -ne $c.score) { $pos++; $v += $pos } else { $v += $D }
    $v += $c.name, $c.cls, $c.rar, $c.gen
    $v += $(if ($null -ne $c.score) { $c.score } else { $D }), $c.nOps, $c.agree
    foreach ($cn in $creators) { $v += $(if ($c.cr.ContainsKey($cn)) { $c.cr[$cn] } else { $D }) }
    $v += $(if ($c.yt) { "$($c.yt) ($($c.ytScore))" } else { $D })
    $v += $(if ($c.atk) { $c.atk } else { $D }), $(if ($c.scAtk) { $c.scAtk } else { $D })
    $v += $c.genDef, $(if ($c.def) { $c.def } else { $D }), $(if ($c.scDef) { $c.scDef } else { $D })
    # COM solo acepta tipos .NET simples (no PSObject): números como double, el resto como texto
    for ($j = 0; $j -lt $v.Count; $j++) {
      $b = if ($v[$j] -is [psobject]) { $v[$j].psobject.BaseObject } else { $v[$j] }
      if ($b -is [int] -or $b -is [long] -or $b -is [double]) { $arr[($i + 1), $j] = [double]$b } else { $arr[($i + 1), $j] = [string]$b }
    }
  }
  $rng = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($rows.Count + 1, $hdr.Count))
  # (Value2 falla aquí: PowerShell reutiliza el tipo String de las asignaciones anteriores a Value2)
  $rng.Formula = $arr
  $ws.Rows.Item(1).Font.Bold = $true
  $ws.Rows.Item(1).WrapText = $true
  $ws.Range('A1:R1').Interior.Color = 0x3A2A1F
  $ws.Range('A1:R1').Font.Color = 0xFFFFFF
  $rng.AutoFilter() | Out-Null
  $ws.Activate(); $x.ActiveWindow.SplitRow = 1; $x.ActiveWindow.SplitColumn = 2; $x.ActiveWindow.FreezePanes = $true
  $ws.Columns.Item(2).ColumnWidth = 32
  foreach ($cc in 1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18) { $ws.Columns.Item($cc).ColumnWidth = 12 }
  $ws.Range("A2:R$($rows.Count + 1)").HorizontalAlignment = -4108
  $ws.Range("B2:B$($rows.Count + 1)").HorizontalAlignment = -4131
  # colores por tier en «Tier general»
  $col = @{ 'S+' = 0x4D9BF5; 'S' = 0x6BB0FF; 'A' = 0x6BE3A3; 'B' = 0x9FE7C9; 'C' = 0xE6D9A7; 'D' = 0xD9C6C6; 'F' = 0xC8C8C8 }
  for ($i = 0; $i -lt $rows.Count; $i++) { $t = $rows[$i].gen; if ($col.ContainsKey($t)) { $ws.Cells.Item($i + 2, 5).Interior.Color = $col[$t] } }

  # ---------- Leyenda ----------
  $wsL = $wb.Worksheets.Item('Leyenda')
  $lastL = $wsL.Cells.Item($wsL.Rows.Count, 1).End(-4162).Row
  $gen = "Hoja «Tier list general»: todas las clases juntas, hasta 5 opiniones independientes de ATAQUE (Vega, Lagacy, MetalSonicDude y Seatin, vía mcoc.app, y MCOC Guide Blog). Cada opinión se pasa a una nota 0-100 y se hace la media; las letras se convierten al centro de la franja de mcoc.app (" + (($minOf.Keys | ForEach-Object { "$_ >= $($minOf[$_])" }) -join ', ') + "). MCOC Scout copia el Guide Blog, así que solo se usa cuando el Guide Blog no valora al campeón. Acuerdo: diferencia entre la opinión más alta y la más baja (<=20 Alto, <=40 Medio, >40 Bajo). Tier general defensa: Guide Blog (o MCOC Scout si falta). Actualizado el $(Get-Date -Format 'dd/MM/yyyy')."
  $found = $false
  for ($r = 1; $r -le $lastL; $r++) {
    switch ($wsL.Cells.Item($r, 1).Text) {
      'Tier general' { $wsL.Cells.Item($r, 2).Value2 = $gen; $found = $true }
      'Actualizado' { $wsL.Cells.Item($r, 2).Value2 = (Get-Date).ToString('d \de MMMM \de yyyy', [Globalization.CultureInfo]'es-ES') }
      'Guide Blog ataque/defensa' { if ($gbDate) { $wsL.Cells.Item($r, 2).Value2 = ($wsL.Cells.Item($r, 2).Text -replace 'MCOC Guide Blog, [^.]+\.', "MCOC Guide Blog, $gbDate.") } }
    }
  }
  if (-not $found) {
    $wsL.Rows.Item(5).Insert(-4121, 0) | Out-Null
    $wsL.Cells.Item(5, 1).Value2 = 'Tier general'; $wsL.Cells.Item(5, 2).Value2 = $gen
  }
  $wsC.Activate()
  $wb.Save()
  Write-Host "OK: hoja «$sheetName» con $($rows.Count) campeones. Ahora ejecuta scripts\excel-to-json.ps1"
} finally {
  $wb.Close($false); $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
