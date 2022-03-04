<#
.NOTES
    The function of this script is to safely place the credentials in secure vault
    This script will install the following powershell modules if not already installed
    1. Microsoft.Powershell.SecretManagement
    2. Microsoft.Powershell.SecretStore

#>

# Define a function to read passowrds
Function Read-Password {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$Message
    )
    do {
        $Password1 = Read-host -Prompt $Message -AsSecureString
        # Decode the secure string
        $Ptr1 = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($Password1)
        $result1 = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr1)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr1)
    
        $Password2 = Read-host -Prompt "Please confirm your password" -AsSecureString
        $Ptr2 = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($Password2)
        $result2 = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr2)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr2)
        If ($result1 -ne $result2) {
            Write-Host "Passowrds do not match. Please try again" -ForegroundColor Yellow
        }
    }until ($result1 -eq $result2)
    return $Password1
}

# Check if required modules exist. If not, install them.
Write-host "Checking if Microsoft.Powershell.SecretManagement exist:" -ForegroundColor Yellow -NoNewline
$secMgmtExist = Get-Module -ListAvailable -Name Microsoft.Powershell.SecretManagement
If (-Not($secMgmtExist)) {
    Write-host "Not Found" -ForegroundColor Red
    Write-Host "Installing module: Microsoft.Powershell.SecretManagement" -ForegroundColor Cyan
    Install-Module Microsoft.Powershell.SecretManagement -scope CurrentUser -Confirm:$false -Force
}else{

    Write-host "Module Found" -ForegroundColor Green
}

Write-host "Checking if Microsoft.Powershell.SecretStore exist:" -ForegroundColor Yellow -NoNewline
$secStoreExist = Get-Module -ListAvailable -Name Microsoft.Powershell.SecretStore
If (-Not($secStoreExist)) {
    Write-host "Not Found" -ForegroundColor Red
    Write-Host "Installing module: Microsoft.Powershell.SecretStore" -ForegroundColor Cyan
    Install-Module Microsoft.Powershell.SecretStore -scope CurrentUser -Confirm:$false -Force

    # Configure the vault
    $secretVaultPassword = Read-Password -Message "Enter the password for locking the secret vault"
    Set-SecretStoreConfiguration -Scope CurrentUser -Authentication Password -PasswordTimeout 3600 -Interaction None `
    -Password $secretVaultPassword -Confirm:$false
    Register-SecretVault -Name gappp -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault -ErrorAction SilentlyContinue
    Unlock-SecretStore -Password $secretVaultPassword
}else{
    Write-host "Module Found" -ForegroundColor Green
    $unlockPassword = Read-host "Secret Vault already configured. Please enter password to unlock it" -AsSecureString
    Unlock-SecretStore -Password $unlockPassword
}

# Register secret vault if not already registered
$sv = Get-SecretVault
if (-NOT($sv)) {
    Register-SecretVault -ModuleName Microsoft.PowerShell.SecretStore -Name gappp -DefaultVault -Confirm:$false
}

$VaultAccount = Read-host "Enter your gmail account (someone@example.com)"
$vaultPassword = Read-Password -Message "Please enter your gmail App password"

# Add username and Password to the vault
Set-Secret -Name gapppuser -Vault gappp -Secret $VaultAccount
Set-Secret -Name gapppPassword -Vault gappp -SecureStringSecret $vaultPassword
