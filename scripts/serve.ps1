<#
  Servidor local mínimo (sin instalar nada) para probar la app y su modo offline.
  Uso:  powershell -ExecutionPolicy Bypass -File scripts\serve.ps1 [-Port 8765]
  Luego abre http://localhost:8765/  (Ctrl+C para parar)
#>
param([int]$Port = 8765)
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$types = @{ '.html' = 'text/html; charset=utf-8'; '.js' = 'text/javascript; charset=utf-8'; '.css' = 'text/css; charset=utf-8';
  '.json' = 'application/json'; '.webmanifest' = 'application/manifest+json'; '.png' = 'image/png'; '.svg' = 'image/svg+xml' }
$l = New-Object System.Net.HttpListener
$l.Prefixes.Add("http://localhost:$Port/")
$l.Start()
Write-Host "Sirviendo $root en http://localhost:$Port/"
try {
  while ($l.IsListening) {
    $ctx = $l.GetContext()
    $path = [Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath.TrimStart('/'))
    if (-not $path) { $path = 'index.html' }
    $file = [IO.Path]::GetFullPath((Join-Path $root $path))
    if ($file.StartsWith($root) -and (Test-Path $file -PathType Leaf)) {
      $bytes = [IO.File]::ReadAllBytes($file)
      $ext = [IO.Path]::GetExtension($file).ToLower()
      $ctx.Response.ContentType = if ($types[$ext]) { $types[$ext] } else { 'application/octet-stream' }
      $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    } else { $ctx.Response.StatusCode = 404 }
    $ctx.Response.Close()
  }
} finally { $l.Stop() }
