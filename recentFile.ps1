$ErrorActionPreference = "Stop"
# Habilitar TLSv1.2 para compatibilidad con clientes más antiguos
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# URL del script de Python
$PythonScriptURL = 'https://raw.githubusercontent.com/astronomyc/filesRecents/main.py'

try {
    # Descargar el script de Python
    $PythonScriptContent = Invoke-WebRequest -Uri $PythonScriptURL -UseBasicParsing | Select-Object -ExpandProperty Content
}
catch {
    Write-Host "No se pudo descargar el script de Python desde $PythonScriptURL."
    exit
}

# Guardar el script de Python en un archivo temporal
$TempPythonScriptPath = "$env:TEMP\main.py"
Set-Content -Path $TempPythonScriptPath -Value $PythonScriptContent

# Ejecutar el script de Python
python $TempPythonScriptPath

# Eliminar el archivo temporal después de la ejecución
Remove-Item $TempPythonScriptPath

Pause