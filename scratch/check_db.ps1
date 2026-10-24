$r = Invoke-RestMethod 'http://localhost:8080/api/ProdChuletas?fechaInicio=2024-01-01&fechaFin=2026-12-31'
Write-Output ("Total registros: " + $r.totalRegistros)
if ($r.totalRegistros -gt 0) {
    Write-Output ("Primer registro fecha: " + $r.datos[0].fechaOrden)
    Write-Output ("Ultimo registro fecha: " + $r.datos[-1].fechaOrden)
    $fechas = $r.datos | ForEach-Object { $_.fechaOrden } | Sort-Object
    Write-Output ("Min fecha: " + $fechas[0])
    Write-Output ("Max fecha: " + $fechas[-1])
}
