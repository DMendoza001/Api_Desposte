$edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$url = "http://localhost:8080/RendimientosChuletas002.html"
$output = & $edgePath --headless=new --dump-dom $url
$output | Out-File -FilePath .\scratch\dump_dom.html -Encoding utf8
Write-Output ("Dump DOM length: " + $output.Length)
