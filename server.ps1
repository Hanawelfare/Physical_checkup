# server.ps1 - Lightweight PowerShell HTTP Server
# Serves the local directory on http://127.0.0.1:8081

$port = 8081
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://127.0.0.1:$port/")

try {
    $listener.Start()
    Write-Host "==========================================================" -ForegroundColor Green
    Write-Host "  Annual Health Check Registration Server Running" -ForegroundColor Green
    Write-Host "  URL: http://127.0.0.1:$port/" -ForegroundColor Cyan
    Write-Host "  Press Ctrl+C to stop the server" -ForegroundColor Yellow
    Write-Host "==========================================================" -ForegroundColor Green

    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response

        $urlPath = $request.Url.LocalPath
        if ($urlPath -eq "/") {
            $urlPath = "/index.html"
        }

        # Resolve local file path
        $filePath = Join-Path $PSScriptRoot $urlPath.TrimStart('/')
        
        if (Test-Path $filePath -PathType Leaf) {
            $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
            $contentType = switch ($extension) {
                ".html" { "text/html; charset=utf-8" }
                ".css"  { "text/css; charset=utf-8" }
                ".js"   { "text/javascript; charset=utf-8" }
                ".png"  { "image/png" }
                ".jpg"  { "image/jpeg" }
                ".ico"  { "image/x-icon" }
                default { "application/octet-stream" }
            }

            $bytes = [System.IO.File]::ReadAllBytes($filePath)
            $response.ContentType = $contentType
            $response.ContentLength64 = $bytes.Length
            $response.Headers.Add("Access-Control-Allow-Origin", "*")
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
        } else {
            $response.StatusCode = 404
            $errorMessage = "404 Not Found"
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($errorMessage)
            $response.ContentType = "text/plain"
            $response.ContentLength64 = $bytes.Length
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
        }
        $response.Close()
    }
} catch {
    Write-Error $_
} finally {
    $listener.Stop()
}
