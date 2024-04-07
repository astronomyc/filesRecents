# Descargar openpyxl manualmente
$openpyxlURL = 'https://files.pythonhosted.org/packages/42/e8/af028681d493814ca9c2ff8106fc62a4a32e4e0ae14602c2a98fc7b741c8/openpyxl-3.1.2.tar.gz'
$openpyxlPath = "$env:TEMP\openpyxl.tar.gz"
Invoke-WebRequest -Uri $openpyxlURL -OutFile $openpyxlPath

# Descomprimir openpyxl
Expand-Archive -Path $openpyxlPath -DestinationPath "$env:TEMP\openpyxl"

# Ruta donde se guardará el script de Python
7z x $openpyxlPath -o"$env:TEMP\openpyxl"
$pythonScriptPath = "$env:TEMP\script.py"

# URL de tu script de Python
$pythonScriptURL = 'https://github.com/astronomyc/filesRecents/main.py'

try {
    $response = Invoke-WebRequest -Uri $pythonScriptURL -UseBasicParsing
    Set-Content -Path $pythonScriptPath -Value $response.Content
}
catch {
    Write-Host "Ocurrió un error al descargar el script: $($_.Exception.Message)"
}

# Añadir la ruta de openpyxl al PYTHONPATH
$env:PYTHONPATH = "$env:TEMP\openpyxl"

# Ruta al intérprete de Python
$pythonPath = "C:\Path\to\python.exe"

# Ejecutar el script de Python
try {
    & $pythonPath $pythonScriptPath
}
catch {
    Write-Host "Ocurrió un error al ejecutar el script: $($_.Exception.Message)"
}

# Limpiar los archivos descargados
Remove-Item -Path $env:TEMP\* -Recurse -Force
