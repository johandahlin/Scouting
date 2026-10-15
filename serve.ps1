# Tiny static web server so phones on the same Wi-Fi can open the scouting app.
# Usage: powershell -ExecutionPolicy Bypass -File serve.ps1 [-Port 8080]
# Stop with Ctrl+C (or close the window).
param([int]$Port = 8080)

$root = $PSScriptRoot
$types = @{
  ".html" = "text/html; charset=utf-8"; ".js" = "text/javascript; charset=utf-8"; ".json" = "application/json";
  ".png" = "image/png"; ".css" = "text/css"; ".ico" = "image/x-icon"; ".svg" = "image/svg+xml"
}

$ips = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
  Where-Object { $_.IPAddress -notlike "127.*" -and $_.IPAddress -notlike "169.254.*" -and $_.PrefixOrigin -ne "WellKnown" } |
  Select-Object -ExpandProperty IPAddress
Write-Host "Serving $root"
foreach ($ip in $ips) { Write-Host "  On iPhone (same Wi-Fi) open:  http://$($ip):$Port/" }

$listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Any, $Port)
$listener.Start()
try {
  while ($true) {
    $client = $listener.AcceptTcpClient()
    try {
      $stream = $client.GetStream()
      $stream.ReadTimeout = 5000
      $reader = New-Object System.IO.StreamReader($stream)
      $requestLine = $reader.ReadLine()
      while (($line = $reader.ReadLine()) -ne $null -and $line -ne "") { }   # skip headers
      if (-not $requestLine) { continue }

      $path = [Uri]::UnescapeDataString(($requestLine -split " ")[1].Split("?")[0])
      if ($path -eq "/") { $path = "/index.html" }
      $file = [IO.Path]::GetFullPath((Join-Path $root $path.TrimStart("/")))

      if ($file.StartsWith($root) -and (Test-Path $file -PathType Leaf)) {
        $body = [IO.File]::ReadAllBytes($file)
        $type = $types[[IO.Path]::GetExtension($file).ToLower()]; if (-not $type) { $type = "application/octet-stream" }
        $status = "200 OK"
      } else {
        $body = [Text.Encoding]::UTF8.GetBytes("Not found"); $type = "text/plain"; $status = "404 Not Found"
      }
      $head = "HTTP/1.1 $status`r`nContent-Type: $type`r`nContent-Length: $($body.Length)`r`nCache-Control: no-cache`r`nConnection: close`r`n`r`n"
      $hb = [Text.Encoding]::ASCII.GetBytes($head)
      $stream.Write($hb, 0, $hb.Length); $stream.Write($body, 0, $body.Length)
      Write-Host "$(Get-Date -Format HH:mm:ss) $status $path"
    } catch {
      Write-Host "Request error: $($_.Exception.Message)"
    } finally {
      $client.Close()
    }
  }
} finally {
  $listener.Stop()
}
