# Definir la URL del script de Python
$PythonScriptURL = 'https://github.com/astronomyc/filesRecents/main.py'

# Descargar el script de Python desde la URL
$PythonScriptContent = Invoke-WebRequest -Uri $PythonScriptURL -UseBasicParsing | Select-Object -ExpandProperty Content

# Guardar el script de Python en un archivo temporal
$TempPythonScriptPath = "$env:TEMP\main.py"
Set-Content -Path $TempPythonScriptPath -Value $PythonScriptContent

# Ejecutar el script de Python utilizando el intérprete de Python local
python $TempPythonScriptPath
