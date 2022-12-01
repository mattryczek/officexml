# Mateusz Ryczek
# 11.30.22
# v1.0

$dbl_url = 'https://raw.githubusercontent.com/mattryczek/officexml/main/configs/DBL.xml'
$odt_url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process -FilePath powershell.exe -Verb RunAs -ArgumentList '-Command', 'cd ${Get-Location}; & .\smart_install.ps1' }

$host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.size(130,2000)
$host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.size(130,30)

$sig = (Invoke-WebRequest -URI https://raw.githubusercontent.com/mattryczek/officexml/main/assets/sig.txt).Content
$motd = (Invoke-WebRequest -URI https://raw.githubusercontent.com/mattryczek/officexml/main/assets/motd.txt).Content

Write-Output $sig`n`r
Write-Output `n`n`r$motd
Write-Output '------------------------------------------------------------'
Write-Output "Getting latest Office Deployment Tool URL..."

$links = (Invoke-WebRequest -uri $odt_url -UseBasicParsing).Links.Href

$url = ($links -like "*officedeploymenttool*")[0]

if ($url) {
    Write-Output "Found URL:`n`n`r${url}`n`n`rDownloading..."
    mkdir -Force .deleteme | Out-Null
    attrib +h .deleteme
    Start-BitsTransfer -Priority High -TransferType Download -Source $url -Destination ".deleteme/odt.exe"

    Write-Output "Extracting setup.exe from Office Deployment Tool"
    Start-Process -FilePath .deleteme\odt.exe -Verb RunAs -ArgumentList "/extract:.deleteme/ /quiet"
    Start-Sleep -Seconds 1.5
    Remove-Item .deleteme\*.xml

    Write-Output "Downloading DBL config..."
    Start-BitsTransfer -Priority High -TransferType Download -Source $dbl_url -Destination ".deleteme/"

    Write-Output "Done. Launching setup.exe with DBL config"
    Start-Process -FilePath .deleteme\setup.exe -Verb RunAs -ArgumentList "/configure .deleteme/DBL.xml"
} else {
    Write-Output "Cannot get URL for Office Deployment Tool. Shutting down..."
    return
}