param(
  [int]$Port = 8789
)

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$server = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Parse("127.0.0.1"), $Port)
$server.Start()

Write-Host "Quote site running: http://127.0.0.1:$Port/"
Write-Host "Press Ctrl+C to stop."

$contentTypes = @{
  ".html" = "text/html; charset=utf-8"
  ".css"  = "text/css; charset=utf-8"
  ".js"   = "application/javascript; charset=utf-8"
  ".json" = "application/json; charset=utf-8"
  ".xls"  = "application/vnd.ms-excel"
  ".xlsx" = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  ".xlsm" = "application/vnd.ms-excel.sheet.macroEnabled.12"
  ".png"  = "image/png"
  ".jpg"  = "image/jpeg"
  ".jpeg" = "image/jpeg"
  ".svg"  = "image/svg+xml"
}

function Send-Response {
  param(
    [System.Net.Sockets.NetworkStream]$Stream,
    [int]$StatusCode,
    [string]$StatusText,
    [string]$ContentType,
    [byte[]]$Body
  )

  $header = "HTTP/1.1 $StatusCode $StatusText`r`nContent-Type: $ContentType`r`nContent-Length: $($Body.Length)`r`nCache-Control: no-store, no-cache, must-revalidate`r`nConnection: close`r`n`r`n"
  $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
  $Stream.Write($headerBytes, 0, $headerBytes.Length)
  $Stream.Write($Body, 0, $Body.Length)
}

function Send-Text {
  param(
    [System.Net.Sockets.NetworkStream]$Stream,
    [int]$StatusCode,
    [string]$StatusText,
    [string]$Text
  )

  $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
  Send-Response $Stream $StatusCode $StatusText "text/plain; charset=utf-8" $bytes
}

try {
  while ($true) {
    $client = $server.AcceptTcpClient()
    try {
      $client.ReceiveTimeout = 2000
      $client.SendTimeout = 2000
      $stream = $client.GetStream()
      $stream.ReadTimeout = 2000
      $stream.WriteTimeout = 2000
      $buffer = New-Object byte[] 8192
      $memoryStream = [System.IO.MemoryStream]::new()

      do {
        $read = $stream.Read($buffer, 0, $buffer.Length)
        if ($read -le 0) { break }
        $memoryStream.Write($buffer, 0, $read)
        $requestText = [System.Text.Encoding]::ASCII.GetString($memoryStream.ToArray())
      } while ($requestText -notmatch "`r`n`r`n" -and $memoryStream.Length -lt 65536)

      if (-not $requestText) {
        continue
      }

      $firstLine = ($requestText -split "`r?`n")[0]
      $parts = $firstLine -split " "
      $requestPath = "index.html"

      if ($parts.Length -ge 2) {
        $requestPath = [Uri]::UnescapeDataString($parts[1].TrimStart("/"))
        if ([string]::IsNullOrWhiteSpace($requestPath)) {
          $requestPath = "index.html"
        }
      }

      $fullPath = Join-Path $root $requestPath
      $resolved = Resolve-Path -LiteralPath $fullPath -ErrorAction SilentlyContinue

      if (-not $resolved) {
        Send-Text $stream 404 "Not Found" "Not Found"
        continue
      }

      $resolvedPath = [System.IO.Path]::GetFullPath($resolved.Path)
      $rootPath = [System.IO.Path]::GetFullPath($root)
      if (-not $resolvedPath.StartsWith($rootPath, [System.StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
        Send-Text $stream 404 "Not Found" "Not Found"
        continue
      }

      $bytes = [System.IO.File]::ReadAllBytes($resolvedPath)
      $ext = [System.IO.Path]::GetExtension($resolvedPath).ToLowerInvariant()
      $contentType = $contentTypes[$ext]
      if (-not $contentType) {
        $contentType = "application/octet-stream"
      }
      Send-Response $stream 200 "OK" $contentType $bytes
    } catch [System.IO.IOException] {
      Write-Host "Request timed out or disconnected."
    } catch {
      try {
        if ($stream) {
          Send-Text $stream 500 "Server Error" "Server Error"
        }
      } catch {}
      Write-Host "Request failed: $($_.Exception.Message)"
    } finally {
      if ($memoryStream) { $memoryStream.Dispose() }
      $client.Close()
    }
  }
} finally {
  $server.Stop()
}
