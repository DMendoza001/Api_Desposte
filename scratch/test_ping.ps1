$sw = [System.Diagnostics.Stopwatch]::StartNew()
try {
    $res = Invoke-RestMethod -Uri "http://RAEVFINL016:8080/api/health" -TimeoutSec 10
    $sw.Stop()
    Write-Output ("Exito RAEVFINL016: " + ($res | ConvertTo-Json -Compress) + " en " + $sw.ElapsedMilliseconds + "ms")
} catch {
    $sw.Stop()
    Write-Output ("Fallo RAEVFINL016 en " + $sw.ElapsedMilliseconds + "ms: " + $_.Exception.Message)
}

$sw2 = [System.Diagnostics.Stopwatch]::StartNew()
try {
    $res2 = Invoke-RestMethod -Uri "http://localhost:8080/api/health" -TimeoutSec 10
    $sw2.Stop()
    Write-Output ("Exito localhost: " + ($res2 | ConvertTo-Json -Compress) + " en " + $sw2.ElapsedMilliseconds + "ms")
} catch {
    $sw2.Stop()
    Write-Output ("Fallo localhost en " + $sw2.ElapsedMilliseconds + "ms: " + $_.Exception.Message)
}
