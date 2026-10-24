# Start Edge with remote debugging and a temporary user data dir
$edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$userDir = "$env:TEMP\edge_test_profile"
if (Test-Path $userDir) { Remove-Item -Recurse -Force $userDir }

$p = Start-Process -FilePath $edge -ArgumentList "--headless=new", "--remote-debugging-port=9222", "--user-data-dir=$userDir", "--disable-gpu", "http://localhost:8080/RendimientosChuletas002.html" -PassThru

Start-Sleep -Seconds 3

try {
    $tabs = Invoke-RestMethod -Uri "http://127.0.0.1:9222/json"
    Write-Output ("Tabs: " + ($tabs | ConvertTo-Json -Compress))
} catch {
    Write-Output ("Error connecting to CDP: " + $_.Exception.Message)
} finally {
    if ($p -and -not $p.HasExited) {
        Stop-Process -Id $p.Id -Force
    }
}
