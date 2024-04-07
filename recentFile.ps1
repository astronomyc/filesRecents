# Configurar la directiva de ejecución para permitir la ejecución de scripts
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force

# Ejecutar el comando ipconfig
ipconfig
