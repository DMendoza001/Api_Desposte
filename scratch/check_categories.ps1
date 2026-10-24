$r = Invoke-RestMethod 'http://localhost:8080/api/ProdChuletas?fechaInicio=2024-01-01&fechaFin=2026-12-31'
$tipos = $r.datos | Select-Object -ExpandProperty tipo -Unique
$chuletas = $r.datos | Select-Object -ExpandProperty tipoChuleta -Unique
$plantillas = $r.datos | Select-Object -ExpandProperty nPlantilla -Unique

Write-Output "=== TIPOS ==="
$tipos | ForEach-Object { Write-Output $_ }
Write-Output "=== TIPOS CHULETA ==="
$chuletas | ForEach-Object { Write-Output $_ }
Write-Output "=== PLANTILLAS (Total: $($plantillas.Count)) ==="
$plantillas | Select-Object -First 20 | ForEach-Object { Write-Output $_ }
