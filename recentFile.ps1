Clear-Host
Write-Host "-------------------------------------------------------------------------"
Write-Host ""
Write-Host ""
Write-Host "                    Bienvenido al Script para"
Write-Host "                    verificar el archivo mas"
Write-Host "                    reciente de cada carpeta"
Write-Host ""
Write-Host "                                                             astronomyc"
Write-Host "-------------------------------------------------------------------------"

# Definir la clave
$pass = "skyrecientes2024*"

# Solicitar la clave al usuario
$inputPass = Read-Host "Ingrese clave de confirmacion"

# Verificar si la clave ingresada es correcta
Clear-Host
if ($inputPass -ne $pass) {
    Write-Host "Clave no autentificada, Cerrando el script"
    exit 3
}

# Obtener la ruta del directorio desde el usuario
$rute_dir = Read-Host "Ingrese la ruta del directorio"

# Obtener la cantidad total de carpetas a procesar
$totalFolders = (Get-ChildItem $rute_dir -Directory).Count
$currentFolderIndex = 0

# Crear un nuevo objeto Excel
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Worksheets.Item(1)

# Configurar el ancho de las columnas
$sheet.Columns.Item(1).ColumnWidth = 16
$sheet.Columns.Item(2).ColumnWidth = 48
$sheet.Columns.Item(3).ColumnWidth = 12

# Escribir las primeras dos filas
$sheet.Cells.Item(1,1).Value2 = "Archivos Recientes"
$sheet.Range("A1:C1").Merge()
$sheet.Cells.Item(1,1).HorizontalAlignment = -4108 # Centrado horizontal
$sheet.Cells.Item(1,1).VerticalAlignment = -4108 # Centrado vertical

$sheet.Cells.Item(2,1).Value2 = "Carpeta"
$sheet.Cells.Item(2,2).Value2 = "Nombre del archivo"
$sheet.Cells.Item(2,3).Value2 = "Fecha"

$sheet.Range("A1:C1").Interior.Color = 15835643 # Color de relleno
$sheet.Range("A2:C2").Interior.Color = 15849703 # Color de relleno

# Función para obtener el archivo más reciente en un directorio (y sus subdirectorios)
function Get-MostRecentFile {
    param(
        [string]$directory
    )

    $recentFile = $null
    $recentDate = $null

    Get-ChildItem $directory -Recurse | Where-Object { -not $_.PSIsContainer } | ForEach-Object {
        $fileDate = $_.LastWriteTime
        if (!$recentFile -or $fileDate -gt $recentDate) {
            $recentFile = $_.Name
            $recentDate = $fileDate
        }
    }

    return $recentFile, $recentDate
}

Clear-Host
# Recorrer cada directorio en la ruta proporcionada
foreach ($folder in Get-ChildItem $rute_dir -Directory) {
    $currentFolderIndex++
    $folderPath = $folder.FullName

    # Obtener el archivo más reciente y la fecha de la carpeta
    $fileRecent, $dateRecent = Get-MostRecentFile $folderPath
    $folderName = $folder.Name

    # Formatear la fecha solo si no es nula
    if ($dateRecent) {
        $formattedDate = Get-Date $dateRecent -Format "dd/MM/yyyy"
    } else {
        $formattedDate = ""
    }


# Calcular el progreso y mostrarlo en la consola
    $progressPercentage = [Math]::Round(($currentFolderIndex / $totalFolders) * 100, 2)
    $status = "Procesando carpeta $currentFolderIndex de $totalFolders ($progressPercentage% completado)..."
    Write-Progress -Activity "Procesando carpetas / Hecho por David para Daniel - Contactame con: devastronomyc@outlook.com" -Status $status -PercentComplete $progressPercentage

    # Imprimir líneas adicionales en la consola
    Write-Host "Carpeta: $folderName"
    Write-Host "Archivo mas reciente: $fileRecent"
    Write-Host "Fecha mas reciente: $formattedDate"
    Write-Host " "
    Write-Host "HECHO POR: ASTRONOMYC"
    Write-Host " "
    Write-Host " "



    # Escribir los datos en el archivo Excel
    $sheet.Cells.Item($sheet.UsedRange.Rows.Count + 1, 1).Value2 = $folderName
    $sheet.Cells.Item($sheet.UsedRange.Rows.Count, 2).Value2 = $fileRecent

    # Formatear la fecha como texto antes de escribirla en la hoja de cálculo
    $sheet.Cells.Item($sheet.UsedRange.Rows.Count, 3).NumberFormat = "@" # Formatear la celda como texto
    $sheet.Cells.Item($sheet.UsedRange.Rows.Count, 3).Value2 = $formattedDate
}

# Ocultar la barra de progreso al finalizar
Write-Progress -Activity "Procesando carpetas" -Completed

# Preguntar al usuario si desea guardar el archivo de Excel
$saveExcel = Read-Host "Desea guardar el archivo de Excel? (Si/No)"

if ($saveExcel -eq "Sí" -or $saveExcel -eq "sí" -or $saveExcel -eq "Si" -or $saveExcel -eq "si") {
    # Guardar el archivo de Excel
    $excel.Visible = $false
    $workbook.SaveAs("$rute_dir\ArchivosRecientes.xlsx")
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable excel, workbook, sheet
    Write-Host "Archivo de Excel guardado exitosamente."
} else {
    # Si el usuario elige no guardar el archivo, cerrar Excel y liberar recursos
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable excel, workbook, sheet
    Write-Host "El archivo de Excel no se guardo."
}
