$ErrorActionPreference = 'Stop'

$manifestUrl = if ($env:WORD_GPT_MANIFEST_URL) {
    $env:WORD_GPT_MANIFEST_URL
} else {
    'https://raw.githubusercontent.com/gzsteven666/word-GPT-Plus/master/release/instant-use/manifest.xml'
}

$installDirectory = Join-Path $env:USERPROFILE 'Documents\WordGPT-Plus'
$manifestPath = Join-Path $installDirectory 'manifest.xml'
$shareName = 'WordGPTPlus'
$catalogId = '{6d77c5ce-b83f-42dd-b8cd-40b48fd32970}'
$catalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"

New-Item -ItemType Directory -Force -Path $installDirectory | Out-Null
Invoke-WebRequest -Uri $manifestUrl -OutFile $manifestPath

[xml]$manifest = Get-Content -Raw $manifestPath
if ($manifest.OfficeApp.DisplayName.DefaultValue -ne 'GPT Plus Steven') {
    throw 'The downloaded file is not the expected GPT Plus Steven manifest.'
}

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdministrator = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($isAdministrator) {
    if (-not (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue)) {
        New-SmbShare -Name $shareName -Path $installDirectory -ReadAccess "$env:USERDOMAIN\$env:USERNAME" | Out-Null
    }
    $catalogUrl = "\\$env:COMPUTERNAME\$shareName"
} else {
    throw 'Run PowerShell as Administrator so the installer can create the local network share.'
}

New-Item -Path $catalogKey -Force | Out-Null
New-ItemProperty -Path $catalogKey -Name Url -PropertyType String -Value $catalogUrl -Force | Out-Null
New-ItemProperty -Path $catalogKey -Name Flags -PropertyType DWord -Value 1 -Force | Out-Null

Write-Host "Installed manifest: $manifestPath" -ForegroundColor Green
Write-Host "Trusted catalog: $catalogUrl" -ForegroundColor Green
if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    Write-Host 'Save your documents, close Word completely, and reopen it.' -ForegroundColor Yellow
}
Write-Host 'In Word, choose Insert > Get Add-ins > Shared Folder > GPT Plus Steven.'
Write-Host 'Future app updates load automatically when the add-in is reopened.'
