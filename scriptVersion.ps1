Param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$rutaEjecutable
)

# Verificar si el archivo ejecutable existe
if (Test-Path -Path $rutaEjecutable) {
    $versionInfo = (Get-Item $rutaEjecutable).VersionInfo
    $nombreProducto = $versionInfo.ProductName
    $versionProducto = $versionInfo.ProductVersion

    Write-Host "Nombre del producto: $nombreProducto"
    Write-Host "Version del producto: $versionProducto"
} else {
    Write-Host "El archivo ejecutable en $rutaEjecutable no se encontró."
}

Read-Host "Enter para cerrar."
