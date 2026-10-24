$r = Invoke-RestMethod 'http://localhost:8080/api/ProdChuletas?fechaInicio=2026-10-17&fechaFin=2026-10-24'
Write-Output ("Total registros 17 al 24: " + $r.totalRegistros)
if ($r.totalRegistros -gt 0) {
    Write-Output ("Primer registro: " + ($r.datos[0] | ConvertTo-Json -Compress))
}
