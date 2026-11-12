<#
  Extrae del texto de los kits (MCOCHUB, vía el repo Shadey1/Mcoc-help) los efectos que cada campeón
  PROVOCA al rival y los que OBTIENE, incluidos los propios de cada campeón (p. ej. «Ionic Overcharge»),
  que el Excel original no recogía. Solo guarda los NOMBRES de los efectos, nunca el texto de los kits.

  Uso:
    powershell -ExecutionPolicy Bypass -File scripts\kit-effects.ps1 -Abilities <ruta a abilities.json>
  (abilities.json se obtiene descargando https://codeload.github.com/Shadey1/Mcoc-help/zip/refs/heads/main
   y está en data/champions/abilities.json)

  Genera data\kit-effects.js  ->  self.MCOC_KITFX = { "<slug MCOCHUB>": { "prov": [...], "gain": [...] } }
#>
param(
  [Parameter(Mandatory = $true)][string]$Abilities,
  [string]$Out = (Join-Path $PSScriptRoot '..\data\kit-effects.js'),
  [switch]$Preview
)
$ErrorActionPreference = 'Stop'
$j = [IO.File]::ReadAllText((Resolve-Path $Abilities)) | ConvertFrom-Json

# Palabras que no son efectos aunque vayan en mayúscula
$stop = @('The','This','That','These','Those','Each','Every','Any','All','Opponent','Opponents','Champion','Champions','When','While','If','After','Before','During','Upon',
  'Their','His','Her','Its','They','He','She','It','A','An','One','Two','Three','Four','Five','Six','Ten','Up','Special','Attack','Attacks','Heavy','Light','Medium',
  'Basic','Critical','Hit','Hits','Block','Blocked','Max','Signature','Ability','Abilities','Combo','Persistent','Charge','Charges','Debuff','Debuffs','Buff','Buffs',
  'Passive','Passives','Effect','Effects','Duration','Potency','Fight','Start','Seconds','Second','Non','Contact','Damage','Power','Health','Bar','Bars','Stack',
  'Stacks','Incoming','Both','Lower','Increase','Increases','Reduce','Reduces','Remove','Removes','Personal','Outgoing','Dash','Back','Parry','Defender','Attacker','Contest','Mode','Form','Stance','State','Unique','Indefinite','Permanent','New',
  'Additional','Extra','Another','Base','Instant','Over','Time','For','With','Against','From','Of','In','On','To','And','Or','But','Also','Instead','Otherwise')
$stopSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$stop | ForEach-Object { [void]$stopSet.Add($_) }

$name = "[A-Z][A-Za-z'’\-]+(?: (?:of |the )?[A-Z][A-Za-z'’\-]+){0,3}"
$rxProv = @(
  "(?:[Ii]nflicts?|[Aa]pplies|[Aa]pply|[Pp]laces?|[Tt]ransfers?)(?: (?:a|an|the|one|two|three|\d+|up to \d+|a stack of|stacks? of|an indefinite|a non-stacking|a permanent|another))* ($name)",
  "($name) (?:Debuffs?|Passive Debuffs?)\b",
  "(?:Opponent|opponent) (?:is|becomes|are) (?:passively |instantly |immediately )?($name)",
  "($name) the [Oo]pponent\b",
  "replaced (?:with|by) (?:a |an )?(?:instant |indefinite |non-stacking |permanent )*($name)",
  "(?:inflicts?|applies) (?:a |an )?(?:instant |indefinite |non-stacking |permanent )*Passive ($name)"
)
$rxGain = @(
  "(?:[Gg]ains?|[Gg]rants?|[Aa]ctivates?|[Ee]nters?|[Tt]riggers?)(?: (?:a|an|the|one|two|three|\d+|up to \d+|a stack of|stacks? of|an indefinite|a non-stacking|a permanent|another|him|her|them|itself|himself|herself))* ($name)",
  "($name) (?:Buffs?|Passive Buffs?)\b"
)
function Clean([string]$s) {
  $w = New-Object 'System.Collections.Generic.List[string]'
  foreach ($x in ($s -split ' ')) { if ($x) { $w.Add($x) } }
  while ($w.Count -gt 0 -and $stopSet.Contains($w[0])) { $w.RemoveAt(0) }
  while ($w.Count -gt 0 -and $stopSet.Contains($w[$w.Count - 1])) { $w.RemoveAt($w.Count - 1) }
  if ($w.Count -eq 0) { return $null }
  $r = ($w -join ' ').Trim("'’- ")
  if ($r.Length -lt 3 -or $r -match '^\d') { return $null }
  return $r
}

# Glosario inglés -> español (el mismo del Excel)
$glos = @{}
(Get-Content (Join-Path $PSScriptRoot 'glosario.json') -Raw -Encoding UTF8 | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $glos[$_.Name.ToLowerInvariant()] = $_.Value }
$ti = (Get-Culture).TextInfo
# Nombres de campeones (para no confundirlos con efectos)
$champNames = @()
$dj = Join-Path $PSScriptRoot '..\data\data.json'
if (Test-Path $dj) { $champNames = (Get-Content $dj -Raw -Encoding UTF8 | ConvertFrom-Json).sheets.Campeones.rows | Select-Object -Skip 1 | ForEach-Object { ($_[0] -replace '\s*\(.*\)', '').Trim() } | Where-Object { $_.Length -gt 3 } | Sort-Object -Unique }
# Variantes de efectos conocidos
$alias = @{ 'stunned' = 'Stun'; 'stuns' = 'Stun'; 'nullified' = 'Nullify'; 'evasion' = 'Evade'; 'regen' = 'Regeneration'; 'phase' = 'Phasing or Phase';
  'resist physical' = 'Physical Resistance'; 'resist energy' = 'Energy Resistance'; 'roots' = 'Root'; 'rooted' = 'Root'; 'power locked' = 'Power Lock';
  'armor broken' = 'Armor Break'; 'incinerated' = 'Incinerate'; 'poisoned' = 'Poison'; 'shocked' = 'Shock'; 'bleeding' = 'Bleed'; 'heal blocked' = 'Heal Block';
  'special locked' = 'Special Lock'; 'speciallock' = 'Special Lock'; 'power burned' = 'Power Burn'; 'power drained' = 'Power Drain'; 'slowed' = 'Slow'; 'staggered' = 'Stagger'; 'feared' = 'Fear'
  # sinónimos y plurales vistos en los kits (revisión del 12/11/2026)
  'crit resistance' = 'Critical Resistance'; 'shock resist' = 'Shock Resistance'; 'degen' = 'Degeneration'; 'petrified' = 'Petrify'
  'purified' = 'Purify'; 'invisible' = 'Invisibility'; 'disintegrate' = 'Disintegration'; 'suppress' = 'Suppression'
  'instantly power drain' = 'Power Drain'; 'arcane fonts' = 'Arcane Font'; 'mind blasts' = 'Mind Blast'; 'porker poppers' = 'Porker Popper'
  'peter porker power popper' = 'Porker Popper'; 'pym particles' = 'Pym Particle'; 'soul wisps' = 'Soul Wisp'
  'inflict carnage bomb' = 'Carnage Bomb'; 'random carnage bomb' = 'Carnage Bomb'; 'nullify spider-seal' = 'Spider-Seal'; 'nullifying spider-seal' = 'Spider-Seal'
  'cross-fight devolve' = 'Devolve'; 'cross-fight suppression' = 'Suppression'; 'severe power sting' = 'Power Sting'; 'severe rupture' = 'Rupture'
  'severe wither' = 'Wither'; 'severe fumble' = 'Fumble'; 'mind blast' = 'Mind Blast' }
# Palabras sueltas que no son efectos
$junk = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
'Opponent Opponents Gain Heal Whenever Mastery Resistance Sting Drain Rating Lock Shield Steal True Opponent Once Inflict Burn Detonation Detonate Intercepts Interruption Damage-Over-Time Pre-Fight Direct Guaranteed Locked Efficiency Cosmic Psionic Non-Damaging Proficiency Momentum Resolve Mutation Iso Symbiote Dialed Persuasion Then Until Whilst Also Afterwards Additionally Otherwise Instead'.Split(' ') | ForEach-Object { [void]$junk.Add($_) }
# Verbos, nombres de personas y fragmentos sueltos que no son efectos (revisión del 12/11/2026)
'Lowers Make Steals Breaks Random Ranks Prevent Ignore Cause Active Attack Blocks Locks Struck Inflicted Extend Meter Hour Esp Tony Laura Gwen Frank Ohnn Death Energy Mystical Tactical Telepath Dimension Dragon Severe Paused Clues Strikers Seed Soul Morph Ego Solar Cerebro Dormammu Agility Dexterity Analysis Guidance Investigation Protocol Absorb Focus Creep Hate Hater Placebo Dossier Gadget Hydration Careful Rout Om'.Split(' ') | ForEach-Object { [void]$junk.Add($_) }
function Normalize([string]$n) {
  if (-not $n) { return $null }
  $n = $n -replace '^(?i)(Passively|Whenever) ', ''
  $n = $n -replace "['’]s?$", ''
  if ($n -match "['’]" -or $n -match '^(?i)Purify') { return $null }
  if ($junk.Contains($n)) { return $null }
  foreach ($cn in $champNames) { if ($n.StartsWith($cn)) { return $null } }
  if ($n -notmatch ' ' -and $n -match 'ing$') { return $null }     # «Damaging», «Striking», «Intercepting»…
  if ($n -match '(?i)\bMastery$|^Accuracy of |^Bursts of |^Knock(ed|ing) |^Crowd |^Purifies |^Slowed |^Aturdimiento pasivo$') { return $null }   # maestrías y fragmentos
  if ($alias.ContainsKey($n.ToLowerInvariant())) { $n = $alias[$n.ToLowerInvariant()] }
  $n = $n -replace "['’]s$", ''                                   # posesivos: «Storm's»
  if ($n -cmatch '^[A-Z\s\-]+$') { $n = $ti.ToTitleCase($n.ToLowerInvariant()) }   # GAMMA RAGE -> Gamma Rage
  if ($n -match '(?i)\b(immunity|immune|resistance to|vulnerab\w* to)$') { return $null }
  if ($n -match '(?i)^(immunity|immunities|stack|stacks|build ups?)$') { return $null }
  $n = $n -replace '(?i) Build Ups?$', ' Build Up'
  if ($n.Length -lt 3) { return $null }
  if (($n -split ' ').Count -gt 4 -or $n -cmatch ' (?!of\b)[a-z]') { return $null }      # frases, no nombres
  $k = $n.ToLowerInvariant()
  if ($glos.ContainsKey($k)) { return $glos[$k] }
  if ($k.EndsWith('s') -and $glos.ContainsKey($k.TrimEnd('s'))) { return $glos[$k.TrimEnd('s')] }
  foreach ($suf in 'ed', 'd') { if ($k.EndsWith($suf) -and $glos.ContainsKey($k.Substring(0, $k.Length - $suf.Length))) { return $glos[$k.Substring(0, $k.Length - $suf.Length)] } }
  # «Energy Power Burn», «Passive Stun», «Instant Shock» -> el efecto base si está en el glosario
  $m2 = [regex]::Match($k, '^(energy|physical|passive|instant|personal|special|true|indefinite) (.+)$')
  if ($m2.Success -and $glos.ContainsKey($m2.Groups[2].Value)) { return $glos[$m2.Groups[2].Value] }
  if ($n -match "['’]s ") { return $null }                       # «Kushala's Damaging»
  return $n
}
# Quita de una lista los nombres del propio campeón y los fragmentos de otro efecto («Ionic» si está «Ionic Overcharge»)
function Prune($list, [string]$slug) {
  $words = $slug -split '-' | Where-Object { $_.Length -gt 2 }
  $out2 = New-Object 'System.Collections.Generic.List[string]'
  foreach ($n in $list) {
    $low = $n.ToLowerInvariant()
    if ($words | Where-Object { $low -eq $_ -or $low -eq ($_ -replace 'the', '') }) { continue }
    if ($list | Where-Object { $_ -ne $n -and $_.StartsWith($n + ' ') }) { continue }
    $out2.Add($n)
  }
  return , $out2.ToArray()
}

$res = [ordered]@{}
foreach ($p in $j.champions.PSObject.Properties) {
  $k = $p.Value.kit
  $lines = @()
  if ($k.signature) { $lines += $k.signature.lines }
  foreach ($c in $k.cards) { $lines += $c.lines }
  $prov = New-Object 'System.Collections.Generic.List[string]'
  $gain = New-Object 'System.Collections.Generic.List[string]'
  foreach ($l in $lines) {
    if (-not $l) { continue }
    $l = $l -replace 'Special Lock(?:ed|s)?\b', 'SpecialLock' -replace 'Special Attack Lock(?:ed|s)?\b', 'SpecialLock'
    foreach ($rx in $rxProv) { foreach ($m in [regex]::Matches($l, $rx)) { $n = Normalize (Clean $m.Groups[1].Value); if ($n -and -not $prov.Contains($n)) { $prov.Add($n) } } }
    foreach ($rx in $rxGain) { foreach ($m in [regex]::Matches($l, $rx)) { $n = Normalize (Clean $m.Groups[1].Value); if ($n -and -not $gain.Contains($n)) { $gain.Add($n) } } }
  }
  $key = $p.Name; if ($p.Value.source -and $p.Value.source.url) { $key = ($p.Value.source.url -split '/')[-1] }
  $res[$key] = [ordered]@{ prov = (Prune @($prov) $p.Name); gain = (Prune @($gain) $p.Name) }
}

if ($Preview) { return $res }
$json = $res | ConvertTo-Json -Depth 4 -Compress
[IO.File]::WriteAllText((Resolve-Path (Split-Path $Out)).Path + '\' + (Split-Path $Out -Leaf), "self.MCOC_KITFX=$json;", (New-Object System.Text.UTF8Encoding($false)))
Write-Host "OK: $($res.Count) campeones -> $Out"
