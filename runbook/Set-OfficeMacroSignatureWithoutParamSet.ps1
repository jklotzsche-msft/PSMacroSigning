<#
    .SYNOPSIS
    Signs an Office Macro with a certificate.
    .DESCRIPTION
    Signs an Office Macro with a certificate. The certificate can be specified by a file path or by a certificate store.
    .PARAMETER SignToolPath
    Specifies the path to the SignTool.exe. If not specified, the default path is C:\SignTool.
    .PARAMETER FileStream
    Specifies the file stream of the Office Macro.
    .PARAMETER LocalFilePath
    Specifies the local file path of the Office Macro.
    .PARAMETER FileName
    Specifies the file name of the Office Macro.
    .PARAMETER LocalCertPath
    Specifies the local file path of the certificate.
    .PARAMETER LocalCertPassword
    Specifies the password of the certificate.
    .PARAMETER CertIssuer
    Specifies the Issuer of the signing cert, or a substring.
    .PARAMETER CertName
    Specifies the Subject Name of the signing cert, or a substring.
    .PARAMETER FileDigestAlgorithm
    Specifies the file digest algorithm to use for creating file signatures. (Default is SHA1)
    .PARAMETER WindowsKitsPath
    Specifies the path to the Windows Kits folder. If not specified, the default path is 'C:\Program Files (x86)\Windows Kits\10\bin\10.0.18362.0\x86\'.

    .EXAMPLE
    #### For testing: Convert file to stream
    PS C:\> ### $fileStream = [convert]::ToBase64String((Get-Content -path "C:\myPath\TestMakro_self-signed.pptm" -Encoding byte))
    PS C:\> $setOfficeMacroSignatureProps = @{
        FileStream = $fileStream
        FileName = "TestMakro_self-signed-stream.pptm"
        CertIssuer = "Sign"
        CertName = "Sign"
    }
    PS C:\> $result = C:\myPath\Set-OfficeMacroSignature.ps1 @setOfficeMacroSignatureProps
    PS C:\> $result[-1] # contains file stream of new file
    PS C:\> #### For testing: Create file from new stream
    PS C:\> ### $FileStreamBytes = [Convert]::FromBase64String($result[-1])
    PS C:\> ### [IO.File]::WriteAllBytes("C:\temp\newfile.pptm",$FileStreamBytes)

    Signing file from file stream using certificate from cert store of local computer

    .EXAMPLE
    #### For testing: Convert file to stream
    PS C:\> ### $fileStream = [convert]::ToBase64String((Get-Content -path "C:\myPath\TestMakro_self-signed.pptm" -Encoding byte))
    PS C:\> $setOfficeMacroSignatureProps = @{
        FileStream = $fileStream
        FileName = "TestMakro_self-signed-stream.pptm"
        LocalCertPath = "C:\myPath\myCert.pfx"
        LocalCertPassword = (Read-Host -Prompt "Enter .PFX Password" -AsSecureString)
    }
    PS C:\> $result = C:\myPath\Set-OfficeMacroSignature.ps1 @setOfficeMacroSignatureProps
    PS C:\> $result[-1] # contains file stream of new file
    PS C:\> #### For testing: Create file from new stream
    PS C:\> ### $FileStreamBytes = [Convert]::FromBase64String($result[-1])
    PS C:\> ### [IO.File]::WriteAllBytes("C:\temp\newfile.pptm",$FileStreamBytes)

    Signing file from file stream using certificate from file

    .EXAMPLE
    $setOfficeMacroSignatureProps = @{
        LocalFilePath = "C:\myPath"
        FileName = "TestMakro_self-signed.pptm"
        CertIssuer = "Sign"
        CertName = "Sign"
    }
    PS C:\> C:\myPath\Set-OfficeMacroSignature.ps1 @setOfficeMacroSignatureProps

    Signing local .pptm file using certificate from cert store of local computer

    .EXAMPLE
    $setOfficeMacroSignatureProps = @{
        LocalFilePath = "C:\myPath"
        FileName = "TestMakro_self-signed.pptm"
        LocalCertPath = "C:\myPath\myCert.pfx"
        LocalCertPassword = (Read-Host -Prompt "Enter .PFX Password" -AsSecureString)
    }
    PS C:\> C:\myPath\Set-OfficeMacroSignature.ps1 @setOfficeMacroSignatureProps

    Signing local .pptm file using certificate from file
#>
[CmdletBinding()]
param(
    [String]
    $SignToolPath = "C:\SignTool",

    [String]
    $FileStream,

    [String]
    $LocalFilePath,

    [String]
    $FileName,

    [String]
    $LocalCertPath,

    [SecureString] # if you use a automation account, the LocalCertPassword should be stored as a variable inside the automation account
    $LocalCertPassword,

    [String]
    $CertIssuer,

    [String]
    $CertName,

    [String]
    $FileDigestAlgorithm = "SHA256",

    [String]
    $WindowsKitsPath = 'C:\Program Files (x86)\Windows Kits\10\bin\10.0.18362.0\x86\'
)

# Error handling
$ErrorActionPreference = "Stop"
trap {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Error (@{
            "ResultCode" = "500"
            "Body"       = $_.Exception.Message
        } | ConvertTo-Json)

    # Remove temporarily created file if needed
    if ($fileCreated) {
        Write-Host "Removing temporarily created file..." -NoNewline
        $null = Remove-Item -Path $fileToSign -Confirm:$false
        Write-Host "OK" -ForegroundColor Green
    }
    return
}

# Validate file extension based on signtool.exe
$fileExtensionValidation = @{
    "MSOSIP"  = @{
        "Excel"      = @(".xla", ".xls", ".xlt")
        "PowerPoint" = @(".pot", ".ppa", ".pps", ".ppt")
        "Project"    = @(".mpp", ".mpt")
        "Publisher"  = @(".pub")
        "Visio"      = @(".vdw", ".vdx", ".vsd", ".vss", ".vst", ".vsx", ".vtx")
        "Word"       = @(".doc", ".dot", ".wiz")
    }
    "MSOSIPX" = @{
        "Excel"      = @(".xlam", ".xlsb", ".xlsm", ".xltm")
        "PowerPoint" = @(".potm", ".ppam", ".ppsm", ".pptm")
        "Visio"      = @(".vsdm", ".vssm", ".vstm")
        "Word"       = @(".docm", ".dotm")
    }
}
$fileExtensionAllowed = $false
Write-Host "Validating file extension of $FileName..." -NoNewline
foreach ($key in $fileExtensionValidation.GetEnumerator()) {
    foreach ($subkey in $fileExtensionValidation."$($key.Name)".GetEnumerator()) {
        foreach ($value in $fileExtensionValidation."$($key.Name)"."$($subkey.Name)".GetEnumerator()) {
            if ($FileName -like "*$value") {
                Write-Host "Found validate file extension at $($subkey.Name) from $($key.name)" -ForegroundColor "Green"
                $fileExtensionAllowed = $true
            }
        }
    }
}
if (-not $fileExtensionAllowed) {
    throw "File extension of $FileName is not allowed."
}

# Convert password if needed
if ($LocalCertPath -ne "") {
    if (($null -ne $LocalCertPassword) -and ($LocalCertPassword -is [System.Security.SecureString])) {
        # Convert secure string to plain text
        $LocalCertPasswordPlain = [System.Net.NetworkCredential]::new("", $localCertPassword).Password
    }
    elseif ($null -eq $LocalCertPassword) {
        # Assume we are executing from an automation account and the password is stored as a variable inside the automation account
        $LocalCertPasswordPlain = (Get-AutomationVariable -Name 'LocalCertPassword')
    }
    else {
        # Assume we are executing from a local machine and the password is plain text
        $LocalCertPasswordPlain = $LocalCertPassword
    }
}

# Create file from file stream if needed
if ($FileStream) {
    Write-Host "Creating file $FileName from file stream..." -NoNewline
    $fileToSign = Join-Path -Path $env:TEMP -ChildPath $FileName
    $FileStreamBytes = [Convert]::FromBase64String($FileStream)
    [IO.File]::WriteAllBytes($fileToSign, $FileStreamBytes)
    Write-Host "OK" -ForegroundColor Green
    $fileCreated = $true
}
else {
    $fileToSign = Join-Path -Path $LocalFilePath -ChildPath $FileName
}

# Sign the macro
Write-Host "Run offsign.bat to verify and sign the macro at $fileToSign..." -NoNewline
$offsignPath = Join-Path -Path $SignToolPath -ChildPath 'offsign.bat'
## Determine sign arguments
if ($LocalCertPath -and $LocalCertPasswordPlain) {
    ## needed, if certificate is on disk as .pfx file
    $signArguments = 'sign /f "{0}" /p "{1}" /fd "{2}"' -f $LocalCertPath, $LocalCertPasswordPlain, $FileDigestAlgorithm
}
elseif ($CertIssuer -and $CertName) {
    ## needed, if certificate is in cert store
    $signArguments = 'sign /i "{0}" /n "{1}" /sm /fd "{2}"' -f $CertIssuer, $CertName, $FileDigestAlgorithm
}
else {
    throw "No valid set of sign arguments found."
}

## Verify arguments are always the same
$verifyArguments = 'verify /pa'
$offSignResult = & $offsignPath $WindowsKitsPath $signArguments $verifyArguments $fileToSign
if(($offSignResult | ConvertTo-Json) -like "*You should fix the problem and re-run OffSign*") {
    throw "Error signing file: $offSignResult"
}
Write-Host "OK" -ForegroundColor Green

Write-Host "Returning signed file stream..." -NoNewline
Write-Output (@{
        "ResultCode" = "200"
        "Body"       = ([convert]::ToBase64String((Get-Content -Path $fileToSign -Encoding byte)))
    } | ConvertTo-Json)    
Write-Host "OK" -ForegroundColor Green

# Remove temporarily created file if needed
if ($FileStream) {
    Write-Host "Removing temporarily created file $FileName..." -NoNewline
    $null = Remove-Item -Path $fileToSign -Confirm:$false
    Write-Host "OK" -ForegroundColor Green
}

Write-Host "Signing completed."