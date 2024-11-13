#requires -RunAsAdministrator
#requires -Version 5

<#
.SYNOPSIS
    This script installs a certificate for signing VBA projects and checks if Microsoft Office is installed.

.DESCRIPTION
    The script performs the following tasks:
    1. Checks if the specified certificate is already installed in the local machine's certificate store.
    2. If the certificate is not found, it installs the certificate from the specified path.
    3. Checks if Microsoft Office is installed on the machine.

.PARAMETER SignToolPath
    The path to the SignTool executable. Default is "C:\SignTool".

.PARAMETER CertPath
    The path to the certificate file (.pfx) to be installed. Default is "C:\SignTool\myCert.pfx".

.PARAMETER Thumbprint
    The thumbprint of the certificate to check for in the local machine's certificate store. Default is "0000000000000000000000000000000000000000".

.EXAMPLE
    .\New-PSMacroSigningDeployment.ps1 -SignToolPath "C:\Tools\SignTool" -CertPath "C:\Certs\myCert.pfx" -Thumbprint "0000000000000000000000000000000000000000"

    This example installs the certificate from "C:\Certs\myCert.pfx" and checks if the certificate with the thumbprint "0000000000000000000000000000000000000000" is already installed.

.NOTES
    This script requires administrative privileges to run.
    Ensure that the certificate file (.pfx) is accessible and the correct thumbprint is provided.
#>

[CmdletBinding()]
param (
    [Parameter()]
    [String]
    $SignToolPath = "C:\SignTool",

    [Parameter()]
    [String]
    $CertPath = "C:\SignTool\myCert.pfx",

    [Parameter()]
    [String]
    $Thumbprint = "0000000000000000000000000000000000000000"
)

$ErrorActionPreference = "Stop"
trap {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Install certificate for signing VBA projects to the local certificate store of computer account
Write-Host "Checking, if certificate is already installed..." -NoNewline
if(-not (Test-Path -Path "Cert:\LocalMachine\My\$Thumbprint")) {
    Write-Host "Installing certificate..." -NoNewline
    $null = Import-Certificate -FilePath $CertPath -CertStoreLocation "Cert:\LocalMachine\My"
}
Write-Host "OK" -ForegroundColor Green

# Check if office is installed
# Office must be installed, to be able to register the VBA signing tools
Write-Host "Checking if Office is installed..." -NoNewline
if(-not (Get-Package -Name "*Office*" -ErrorAction SilentlyContinue)) {
    Write-Host "ERROR" -ForegroundColor Red
    throw "Office is not installed. Please install Office first."
}
Write-Host "OK" -foregroundColor Green

# Create new folder at C:\SignTool.
# This folder must be well secured, as this will contain the tools for the signing process.

if(-not (Test-Path -Path $SignToolPath)) {
    Write-Host "Creating '$SignToolPath'..." -NoNewline
    $null = New-Item -Path $SignToolPath -ItemType Directory
    Write-Host "OK" -ForegroundColor Green
}

# Extract the needed tools from 'Microsoft Office Subject Interface Packages for Digitally Signing VBA Projects'
# You can manually download it at https://www.microsoft.com/en-us/download/details.aspx?id=56617
Write-Host "Extracting 'Microsoft Office Subject Interface Packages for Digitally Signing VBA Projects...'" -NoNewline
if(-not (Test-Path -Path (Join-Path -Path $SignToolPath -ChildPath "offsign.bat"))) {
    $sipPath = Join-Path -Path $PSScriptRoot -ChildPath 'officesips_16.0.16507.43425.exe'
    $null = Start-Process -FilePath $sipPath -ArgumentList "/extract:$SignToolPath /quiet /norestart /passive" -Wait
}
Write-Host "OK" -ForegroundColor Green

# Install signtool from Windows SDK from https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/
# As described at https://support.microsoft.com/en-us/topic/upgrade-signed-office-vba-macro-projects-to-v3-signature-kb5000676-2b8b3cae-ad64-4b4b-aa85-c4a98ca6da87, 
# version 10.0.18362.0 is installed and used
Write-Host "Installing Windows SDK..." -NoNewline
if(-not (Test-Path -Path 'C:\Program Files (x86)\Windows Kits\10\bin\10.0.18362.0\x86\signtool.exe')) {
    $winsdkPath = Join-Path -Path $PSScriptRoot -ChildPath 'winsdksetup.exe'
    $null = Start-Process -FilePath $winsdkPath -ArgumentList "/norestart /quiet /ceip off" -Wait
}
Write-Host "OK" -ForegroundColor Green

# Install Microsoft Visual C++ Runtime Libraries as described in readme.txt of the Microsoft Office Subject Interface Packages for Digitally Signing VBA Projects.
# The installer for the redistributable can be found at https://download.microsoft.com/download/C/6/D/C6D0FD4E-9E53-4897-9B91-836EBA2AACD3/vcredist_x86.exe
Write-Host "Installing Microsoft Visual C++ Runtime Libraries..." -NoNewline
if($null -eq (Get-Package -Name "Microsoft Visual C++ 2010  x86 Redistributable*" -ErrorAction SilentlyContinue)) {
    $vcPath = Join-Path -Path $PSScriptRoot -ChildPath 'vcredist_x86.exe'
    $null = Start-Process -FilePath $vcPath -ArgumentList "/q /norestart" -Wait
}
Write-Host "OK" -ForegroundColor Green

# Register DLLs
Write-Host "Registering VBA signing tools..." -NoNewline
foreach($msoSipDll in @("mspsip.dll", "msosipx.dll")) {
    $msoSipDllPath = Join-Path -Path $SignToolPath -ChildPath $msoSipDll
    $null = Start-Process -FilePath regsvr32.exe -ArgumentList "$msoSipDllPath /s" -Wait
}
Write-Host "OK" -foregroundColor Green

Write-Host "PSMacroSigningDeployment finished." -ForegroundColor Green