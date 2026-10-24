$edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$userDir = "$env:TEMP\edge_test_profile_eval"
if (Test-Path $userDir) { Remove-Item -Recurse -Force $userDir }

$p = Start-Process -FilePath $edge -ArgumentList "--headless=new", "--remote-debugging-port=9223", "--user-data-dir=$userDir", "--disable-gpu", "http://localhost:8080/RendimientosChuletas002.html" -PassThru

# Give it 6 seconds to run DOMContentLoaded, detect API, and query cached data
Start-Sleep -Seconds 6

try {
    $tabs = Invoke-RestMethod -Uri "http://127.0.0.1:9223/json"
    $pageTab = $tabs | Where-Object { $_.type -eq 'page' -and $_.url -like '*RendimientosChuletas002.html*' } | Select-Object -First 1

    if ($pageTab) {
        Write-Output ("Found page tab: " + $pageTab.id)
        $wsUri = [System.Uri]$pageTab.webSocketDebuggerUrl
        $ws = New-Object System.Net.WebSockets.ClientWebSocket
        $cts = New-Object System.Threading.CancellationTokenSource(10000)
        $ws.ConnectAsync($wsUri, $cts.Token).Wait()

        # Send Runtime.evaluate to get dashboard state
        $jsToEval = "JSON.stringify({ statusText: document.getElementById('apiStatusText') ? document.getElementById('apiStatusText').textContent : 'NO_STATUS_EL', rawLength: (typeof rawData !== 'undefined' ? rawData.length : -1), kpiOrigen: document.getElementById('kpiOrigen') ? document.getElementById('kpiOrigen').textContent : 'NO_KPI', tbodyText: document.getElementById('tablePlantillasBody') ? document.getElementById('tablePlantillasBody').innerText.substring(0, 300) : 'NO_TBODY' })"

        $cmdObj = @{
            id = 1
            method = "Runtime.evaluate"
            params = @{
                expression = $jsToEval
                returnByValue = $true
            }
        }
        $cmdJson = $cmdObj | ConvertTo-Json -Compress
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($cmdJson)
        $segment = [System.ArraySegment[byte]]::new($buffer)
        $ws.SendAsync($segment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $cts.Token).Wait()

        $recvBuffer = New-Object byte[] 65536
        $recvSegment = [System.ArraySegment[byte]]::new($recvBuffer)
        $recvResult = $ws.ReceiveAsync($recvSegment, $cts.Token).Result
        $responseJson = [System.Text.Encoding]::UTF8.GetString($recvBuffer, 0, $recvResult.Count)
        Write-Output ("Eval Result: " + $responseJson)

        $ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "Done", $cts.Token).Wait()
    } else {
        Write-Output "No page tab found!"
    }
} catch {
    Write-Output ("Error: " + $_.Exception.ToString())
} finally {
    if ($p -and -not $p.HasExited) {
        Stop-Process -Id $p.Id -Force
    }
}
