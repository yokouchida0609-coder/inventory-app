$port = if ($env:PORT) { [int]$env:PORT } else { 8080 }
$root = "C:\Users\yoko7\reminder-app"

$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Any, $port)
$listener.Start()
Write-Host "Serving $root on http://localhost:$port"
[Console]::Out.Flush()

$mimeTypes = @{
    '.html' = 'text/html; charset=utf-8'
    '.js'   = 'application/javascript'
    '.css'  = 'text/css'
    '.json' = 'application/json'
    '.png'  = 'image/png'
    '.ico'  = 'image/x-icon'
}

function Handle-Request($client) {
    $stream = $client.GetStream()
    $reader = [System.IO.StreamReader]::new($stream)
    $requestLine = $reader.ReadLine()
    if (-not $requestLine) { $client.Close(); return }

    # drain headers
    while ($true) {
        $line = $reader.ReadLine()
        if ([string]::IsNullOrEmpty($line)) { break }
    }

    $parts = $requestLine -split ' '
    $urlPath = if ($parts.Count -ge 2) { $parts[1] } else { '/' }
    if ($urlPath -eq '/' -or $urlPath -eq '') { $urlPath = '/routine-check.html' }
    $urlPath = $urlPath -replace '\?.*', ''

    $filePath = Join-Path $root ($urlPath.TrimStart('/').Replace('/', '\'))

    $writer = [System.IO.StreamWriter]::new($stream)
    $writer.NewLine = "`r`n"

    if (Test-Path $filePath -PathType Leaf) {
        $ext = [System.IO.Path]::GetExtension($filePath)
        $mime = if ($mimeTypes[$ext]) { $mimeTypes[$ext] } else { 'application/octet-stream' }
        $bytes = [System.IO.File]::ReadAllBytes($filePath)
        $writer.WriteLine("HTTP/1.1 200 OK")
        $writer.WriteLine("Content-Type: $mime")
        $writer.WriteLine("Content-Length: $($bytes.Length)")
        $writer.WriteLine("Connection: close")
        $writer.WriteLine("")
        $writer.Flush()
        $stream.Write($bytes, 0, $bytes.Length)
    } else {
        $body = [System.Text.Encoding]::UTF8.GetBytes("404 Not Found: $urlPath")
        $writer.WriteLine("HTTP/1.1 404 Not Found")
        $writer.WriteLine("Content-Length: $($body.Length)")
        $writer.WriteLine("Connection: close")
        $writer.WriteLine("")
        $writer.Flush()
        $stream.Write($body, 0, $body.Length)
    }
    $stream.Flush()
    $client.Close()
}

while ($true) {
    $client = $listener.AcceptTcpClient()
    Handle-Request $client
}
