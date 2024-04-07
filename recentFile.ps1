# Obtener la ruta del directorio desde el usuario
$rute_dir = Read-Host "Ingrese la ruta del directorio"

# Crear un nuevo objeto Excel
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Worksheets.Item(1)

# Configurar el ancho de las columnas
$sheet.Columns.Item(1).ColumnWidth = 16
$sheet.Columns.Item(2).ColumnWidth = 48
$sheet.Columns.Item(3).ColumnWidth = 12
$sheet.Columns.Item(4).ColumnWidth = 12

# Escribir las primeras dos filas
$sheet.Cells.Item(1,1).Value2 = "Archivos Recientes"
$sheet.Range("A1:D1").Merge()
$sheet.Cells.Item(1,1).HorizontalAlignment = -4108 # Centrado horizontal
$sheet.Cells.Item(1,1).VerticalAlignment = -4108 # Centrado vertical

$sheet.Cells.Item(2,1).Value2 = "Carpeta"
$sheet.Cells.Item(2,2).Value2 = "Nombre del archivo"
$sheet.Cells.Item(2,3).Value2 = "Fecha"
$sheet.Cells.Item(2,4).Value2 = "Peso"

$sheet.Range("A1:D2").Interior.Color = 13998081 # Color de relleno

# Función para obtener el archivo más reciente en un directorio
function Get-MostRecentFile {
    param(
        [string]$directory
    )

    $recentFile = $null
    $recentDate = $null

    Get-ChildItem $directory -File | ForEach-Object {
        $fileDate = $_.LastWriteTime
        if ($recentFile -eq $null -or $fileDate -gt $recentDate) {
            $recentFile = $_.Name
            $recentDate = $fileDate
        }
    }

    return $recentFile, $recentDate
}

# Función para calcular el tamaño de un directorio
function Get-FolderSize {
    param(
        [string]$folder
    )

    $totalSize = 0

    Get-ChildItem $folder -Recurse -File | ForEach-Object {
        $totalSize += $_.Length
    }

    return $totalSize
}

# Función para convertir el tamaño de bytes a una cadena legible por humanos
function Convert-BytesToHumanReadable {
    param(
        [long]$sizeInBytes
    )

    $units = "B", "KB", "MB", "GB", "TB"
    $index = 0

    while ($sizeInBytes -ge 1024 -and $index -lt $units.Length) {
        $sizeInBytes /= 1024
        $index++
    }

    return "{0:N2} {1}" -f $sizeInBytes, $units[$index]
}

# Función para recorrer todas las subcarpetas de forma recursiva
function Recursively-GetFiles {
    param(
        [string]$folder
    )

    Get-ChildItem $folder -Directory | ForEach-Object {
        $subfolderPath = $_.FullName
        $subfolderName = $_.Name

        $fileRecent, $dateRecent = Get-MostRecentFile $subfolderPath
        $formattedDate = Get-Date $dateRecent -Format "dd/MM/yyyy"
        $folderSize = Get-FolderSize $subfolderPath
        $formattedSize = Convert-BytesToHumanReadable $folderSize

        $sheet.Cells.Item($sheet.UsedRange.Rows.Count + 1, 1).Value2 = $subfolderName
        $sheet.Cells.Item($sheet.UsedRange.Rows.Count, 2).Value2 = $fileRecent
        $sheet.Cells.Item($sheet.UsedRange.Rows.Count, 3).Value2 = $formattedDate
        $sheet.Cells.Item($sheet.UsedRange.Rows.Count, 4).Value2 = $formattedSize

        Recursively-GetFiles $subfolderPath
    }
}

# Recorrer cada directorio en la ruta proporcionada
foreach ($folder in Get-ChildItem $rute_dir -Directory) {
    $folderPath = $folder.FullName

    # Obtener el archivo más reciente, la fecha y el tamaño de la carpeta
    $fileRecent, $dateRecent = Get-MostRecentFile $folderPath
    $formattedDate = Get-Date $dateRecent -Format "dd/MM/yyyy"
    $folderName = $folder.Name
    $folderSize = Get-FolderSize $folderPath
    $formattedSize = Convert-BytesToHumanReadable $folderSize

    # Escribir los datos en el archivo Excel
    $sheet.Cells.Item($sheet.UsedRange.Rows.Count + 1, 1).Value2 = $folderName
    $sheet.Cells.Item($sheet.UsedRange.Rows.Count, 2).Value2 = $fileRecent
    $sheet.Cells.Item($sheet.UsedRange.Rows.Count, 3).Value2 = $formattedDate
    $sheet.Cells.Item($sheet.UsedRange.Rows.Count, 4).Value2 = $formattedSize

    Recursively-GetFiles $folderPath
}

# Guardar el archivo de Excel
$excel.Visible = $false
$workbook.SaveAs("$rute_dir\ArchivosRecientes.xlsx")
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel, workbook, sheet
Write-Host "Complete"
