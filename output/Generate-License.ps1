param(
    [Parameter(Mandatory=$true)]
    [string]$LicenseId,

    [Parameter(Mandatory=$true)]
    [int]$Tier,

    [Parameter(Mandatory=$true)]
    [DateTime]$ExpiryUtc,

    [Parameter(Mandatory=$true)]
    [int]$Seats
)

$ErrorActionPreference = "Stop"

$exeDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$privateKeyPath = Join-Path $exeDir "MetaMech_RSA_PRIVATE.xml"
$outFile = Join-Path $exeDir "license.key"

if (!(Test-Path $privateKeyPath)) {
    Write-Error "Private key not found: $privateKeyPath"
    exit 1
}

# ---------------- PAYLOAD (STRICT FORMAT) ----------------
$payload = "$LicenseId|$Tier|$($ExpiryUtc.ToString("yyyy-MM-ddTHH:mm:ssZ"))|$Seats"

Write-Host "PAYLOAD:"
Write-Host $payload
Write-Host ""

# ---------------- SIGN PAYLOAD ----------------
$rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider
$rsa.FromXmlString((Get-Content $privateKeyPath -Raw))

$payloadBytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
$signatureBytes = $rsa.SignData(
    $payloadBytes,
    [System.Security.Cryptography.CryptoConfig]::MapNameToOID("SHA256")
)

$license =
    [Convert]::ToBase64String($payloadBytes) + "." +
    [Convert]::ToBase64String($signatureBytes)

Set-Content -Path $outFile -Value $license -Encoding ASCII

Write-Host "LICENSE GENERATED SUCCESSFULLY"
Write-Host "Saved to: $outFile"
