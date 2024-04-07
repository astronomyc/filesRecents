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

# Recorrer cada directorio en la ruta proporcionada
foreach ($folder in Get-ChildItem $rute_dir -Directory) {
    $folderPath = $folder.FullName

    # Obtener el archivo más reciente y la fecha de la carpeta
    $fileRecent, $dateRecent = Get-MostRecentFile $folderPath
    $formattedDate = Get-Date $dateRecent -Format "dd/MM/yyyy"
    $folderName = $folder.Name

    # Escribir los datos en el archivo Excel
    $sheet.Cells.Item($sheet.UsedRange.Rows.Count + 1, 1).Value2 = $folderName
    $sheet.Cells.Item($sheet.UsedRange.Rows.Count, 2).Value2 = $fileRecent

    # Formatear la fecha como texto antes de escribirla en la hoja de cálculo
    $sheet.Cells.Item($sheet.UsedRange.Rows.Count, 3).NumberFormat = "@" # Formatear la celda como texto
    $sheet.Cells.Item($sheet.UsedRange.Rows.Count, 3).Value2 = $formattedDate
}

# Guardar el archivo de Excel
$excel.Visible = $false
$workbook.SaveAs("$rute_dir\ArchivosRecientes.xlsx")
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel, workbook, sheet
Write-Host "Complete"
