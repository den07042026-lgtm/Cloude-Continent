$chrome     = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$chromeDir  = "C:\Users\Admin\Desktop\На загрузку\chrome_profile"
$url        = "https://seller.ozon.ru/app/products/import/file"
$port       = 9222

function Test-Port {
    try {
        $t = New-Object Net.Sockets.TcpClient
        $t.Connect("127.0.0.1", $port)
        $t.Close()
        return $true
    } catch { return $false }
}

# Если порт уже открыт — Chrome с отладкой уже запущен, ничего не делаем
if (Test-Port) {
    Write-Host "Chrome with debug port already running."
    exit 0
}

# Запускаем новый Chrome с отдельным профилем — основной Chrome не трогаем
New-Item -ItemType Directory -Force -Path $chromeDir | Out-Null
Write-Host "Starting debug Chrome..."

$proc = Start-Process -FilePath $chrome `
    -ArgumentList "--remote-debugging-port=$port", "--user-data-dir=`"$chromeDir`"", "--no-first-run", $url `
    -PassThru

Write-Host "Chrome PID: $($proc.Id)"

# Ждём открытия порта
for ($i = 0; $i -lt 120; $i++) {
    Start-Sleep -Milliseconds 500
    if ($proc.HasExited) {
        Write-Host "ERROR: Chrome exited with code $($proc.ExitCode)"
        exit 1
    }
    if (Test-Port) {
        Write-Host "Chrome ready after $([math]::Round($i*0.5,1))s"
        exit 0
    }
}

Write-Host "ERROR: port $port did not open in 60s"
exit 1
