<#
.NOTES
    Please update VAR DECLARATION section (lines 8 to 19) before executing the script
#>

##########################  VAR DECLARATION  #######################

$logFilePath = "Enter the path to the log file"
$computerName = "Enter the hostname"
$retryCount = 3 # Specify the number of retires for services which fail to start
$ServiceToMonitor = @(
    "service1",
    "service2",
    "service3"
)
$SMTPServer = "smtp.gmail.com"
$SMTPPort = 587
$Subject = "Service Start Report"
$secretVaultPassword = "Enter the password that was set while registering the App credentials in vault"

######################### END VAR DECLARATION ######################
# Define the function for log writing
Function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter()][String]$FilePath,
        [Parameter()][String]$LogType,
        [Parameter()][String]$LogMessage
    )
    $dateFormate = $(get-date -Format "HH:mm dd MMM yyyy")
    $fullMsg = $dateFormate + " - " + $LogType + " : " + $LogMessage
    $fullMsg | Out-File -FilePath $FilePath -Append -Force
}

# Remove log file
If (Test-Path $logFilePath) {
    Remove-Item -Path $logFilePath -Force -Confirm:$false
}

# Unlock the secret Store
Try{
    Unlock-SecretStore -Password $secretVaultPassword
    $logM = "Secret Store Unlcoked"
    Write-Log -FilePath $logFilePath -LogType "INFO" -LogMessage $logM
}catch{
    $logM = $_.Exception.Message
    Exit
}


# Generate Credentials
$gmailUserName = Get-Secret -Name gapppuser -Vault gappp
$gmailAppPassword = Get-Secret -Name gapppPassword -Vault gappp

# Check if gmail user name does not end with @gmail.com, then append this string accordingly
if (-Not(($gmailUserName).EndsWith('@gmail.com'))) {
    $gmailUserName += "@gmail.com"
}
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $gmailUserName, $gmailAppPassword 

Function Invoke-StartService {
    [CmdletBinding()]
    param (
        [Parameter()][String]$ComputerName = $env:COMPUTERNAME,
        [Parameter(Mandatory)][String]$ServiceName,
        [Parameter(Mandatory)][String]$logFilePath,
        [Parameter()][int]$RetryCount = 1
    )

    $svcObject = Get-Service -ComputerName $ComputerName -Name $ServiceName

    # Identify the services which need to be restarted
    if (($svcObject.StartType -eq 'Automatic') -and ($svcObject.Status -ne 'Running')) {
        # Raise the serviceNotStarted Flag
        $serviceNotStarted = $true
        $rCount = 1
        # Try to start the service
        While ($serviceNotStarted -and ($rCount -le $RetryCount)) {
            Try {
                $logM = "Trying to start the service:$ServiceName on computer: $ComputerName. Attempt Number: $rCount"
                Write-Log -FilePath $logFilePath -LogType "INFO" -LogMessage $logM
                Get-Service -Name $ServiceName -ComputerName $ComputerName | Start-Service -Confirm:$false -ErrorAction Stop
                $serviceNotStarted = $false
                $logM = "Service:$ServiceName on computer: $ComputerName started succesfully. Attempt Number: $rCount"
            }catch{
                $logM = $_.Exception.Message
                Write-Log -FilePath $logFilePath -LogType "ERROR" -LogMessage $logM

                # Increment the count
                $rCount++
            }
        }
        Return [PSCustomObject]@{
            Service = $ServiceName
            Started = $(-Not($serviceNotStarted))
            Tries = $rCount
        }
    }elseif ($svcObject.StartType -ne 'Automatic') {
        $logM =  "The service: $ServiceName is not set to automatic on computer: $ComputerName. Skipping it"
        Write-Log -FilePath $logFilePath -LogType "INFO" -LogMessage $logM
    }else{
        $logM = "The service: $ServiceName already running on computer"
        Write-Log -FilePath $logFilePath -LogType "INFO" -LogMessage $logM
    }
}

# Genearate report array
$report = @()

# Call the service start function
foreach ($svc in $servicesToMonitor) {
    $report += Invoke-StartService -ComputerName $computerName -ServiceName $svc -logFilePath $logFilePath -RetryCount $retryCount 
}

# If there is atleast one failed startup, then send the mail
if (($report.Started).Contains($false)) {
    $mailFrom = $gmailUserName
    $mailTo = $mailFrom
    $body = @"
    <b>Report Generated on: $(Get-Date)</b><br/>
    $('=' * "Report Generated on: $(Get-Date)".Length)<br/>
    1. Following services were successfully started:-<br/>
    <i><b> $(($report | Where-object {$_.Started}).Service -join ", ") </b></i><br/>

    2. Following services failed to start even in $retryCount tries:-<br/>
    <i><b> $(($report | Where-object {-Not($_.Started)}).Service -join ", ") </b></i><br/>
"@

    Send-MailMessage -Body $body -BodyAsHtml -SmtpServer $SMTPServer -UseSsl -Port $SMTPPort `
    -From $mailFrom -To $mailTo -Subject $Subject -Credential $Credentials -Priority High
}
