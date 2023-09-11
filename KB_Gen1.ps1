##CONFIG

$APIKEY = "Enter_API_KEY"
$Endpoint = "https://<NAME>.openai.azure.com/openai/deployments/<DEPLOYMENTNAME>/chat/completions?api-version=2023-07-01-preview"
#Ensure you update the API KEY and Endpoint
$PreInformation = @'

Assume you are working in a standard Windows server environment that primarily uses Microsoft technologies. The server hosts several business applications and is a critical piece of infrastructure.

'@

#----------

# Check and create directory if not exist
$directoryPath = "C:\temp\KBCollection"
if (-not (Test-Path $directoryPath)) {
    New-Item -ItemType Directory -Path $directoryPath
}

# Set the paths for the log and register files within the directory
$errorLogFile = "$directoryPath\Log.json"
$kbRegisterFile = "$directoryPath\KBRegister.json"

# Define the directory for Markdown files
$markdownDirectory = "$directoryPath\KB_Files"
   # Check if the directory exists, and create it if not
                if (-not (Test-Path $markdownDirectory -PathType Container)) {
                    New-Item -Path $markdownDirectory -ItemType Directory -Force
                }

# Define a Mutex for locking
$mutex = New-Object Threading.Mutex($false, "ErrorLogMutex")

# Load the KB Register
if (Test-Path $kbRegisterFile) {
    $kbRegister = Get-Content $kbRegisterFile | ConvertFrom-Json
} else {
    $kbRegister = @()
}



function DetermineITClassification {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ErrorMessage
    )

    # Logic to determine IT classification based on the error message

    if ($ErrorMessage -match "file|directory|path|filesystem|access|permission|disk|storage|backup|restore") {
        return "File/Storage Error"
    } elseif ($ErrorMessage -match "service|network|communication|connection|port|protocol|firewall|routing") {
        return "Service/Network Error"
    } elseif ($ErrorMessage -match "process|application|program|software|execution|runtime|crash|hang|freeze|bug|error") {
        return "Process/Software Error"
    } elseif ($ErrorMessage -match "database|SQL|query|table|record|schema|connection|transaction") {
        return "Database Error"
    } elseif ($ErrorMessage -match "hardware|device|driver|firmware|motherboard|CPU|RAM|disk|keyboard|mouse|monitor") {
        return "Hardware Error"
    } elseif ($ErrorMessage -match "authentication|authorization|login|credentials|token|session|security|SSL|TLS") {
        return "Authentication/Security Error"
    } elseif ($ErrorMessage -match "DNS|domain|hostname|IP address|routing|DNS server|DNS resolution|DNS cache") {
        return "DNS/Host Resolution Error"
    } elseif ($ErrorMessage -match "email|SMTP|IMAP|POP3|email server|send|receive|mailbox|outlook|thunderbird|email client") {
        return "Email/Communication Error"
    } elseif ($ErrorMessage -match "website|web server|HTTP|HTTPS|URL|webpage|browser|web application|web service") {
        return "Website/Web Application Error"
    } elseif ($ErrorMessage -match "printer|printing|print queue|paper jam|printer driver|scanner|fax|printing service") {
        return "Printer/Printing Error"
    } elseif ($ErrorMessage -match "backup|restore|data loss|corruption|disaster recovery|backup solution") {
        return "Backup/Restore Error"
    }

    # If no specific classification is found, return a default classification
    return "Uncategorized"
}

function Get-OpenAIErrorResponse {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Error,
        [Parameter(Mandatory = $true)]
        [string]$Classification
    )

    # Call OpenAI to get the error response based on the IT classification
    # Modify the $body and API call as needed

    $apiEndpoint = $Endpoint
    $headers = @{
        "Content-Type" = "application/json"
        "api-key"      = $APIKEY
    }
    $body = @{
        messages = @(
            @{ role = "system"; content = $PreInformation },  # Add pre-information here
            @{ role = "system"; content = "You are an IT professional looking for technical answers to help resolve the $Classification error." },
            @{ role = "user"; content = "How to fix error: $Error" }
        )
        max_tokens        = 800
        temperature       = 0.7
        frequency_penalty = 0
        presence_penalty  = 0
        top_p             = 0.95
        stop              = $null
    } | ConvertTo-Json

    try {
        $response = Invoke-RestMethod -Uri $apiEndpoint -Method Post -Headers $headers -Body $body
        if ($response -and $response.choices -and $response.choices.Count -gt 0 -and $response.choices[0].message) {
            $errorResponse = $response.choices[0].message.content
            # Remove the redundant text
            $errorResponse = $errorResponse -replace "^To fix the error.*?following steps:", ""
            return $errorResponse
        } else {
            Write-Error "Unexpected response format from OpenAI"
            return $null
        }
    } catch {
        Write-Error "Failed to fetch error response from OpenAI: $_"
        return $null
    }
}





function Handle-Error {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ErrorMessage
    )

    $mutex.WaitOne()

    $errorsList = @()
    if (Test-Path $errorLogFile) {
        $content = Get-Content $errorLogFile | ConvertFrom-Json
        # Ensure $errorsList is always an array
        $errorsList = @($content)
    }

    # Load the existing KB register
    $kbRegister = @()
    if (Test-Path $kbRegisterFile) {
        $kbContent = Get-Content $kbRegisterFile | ConvertFrom-Json
        $kbRegister = @($kbContent)
    }

   $existingMarkdownFile = Get-ChildItem -Path $markdownDirectory -Filter "*.md" | Where-Object {
    $content = Get-Content $_.FullName
    $content -match [regex]::Escape($ErrorMessage)
}

if ($existingMarkdownFile) {
    Write-Host "A Markdown file for this error already exists: $($existingMarkdownFile.FullName)" -ForegroundColor Yellow
    $mutex.ReleaseMutex()
    return
}


     $existingError = $errorsList | Where-Object { $_.error -eq $ErrorMessage }
    if ($existingError) {
        Write-Host "Error found - see details" -ForegroundColor Yellow
        Write-Host $existingError.solutionResponse
    } else {
        # Determine the IT classification (file, service, process, etc.) based on the error message
        $itClassification = DetermineITClassification -ErrorMessage $ErrorMessage
        if ($itClassification) {
            $errorResponse = Get-OpenAIErrorResponse -Error $ErrorMessage -Classification $itClassification
            #$solutionResponse = Get-OpenAISolutionResponse -Error $ErrorMessage
            if ($errorResponse -and $solutionResponse -ne $errorResponse) {
                $errorCode = "" + (Get-UniqueRefCode)
                Write-Host "IT Classification: $itClassification" -ForegroundColor Cyan
                Write-Host "Error Response:"
                Write-Host $errorResponse
                #Write-Host "Solution Response:"
                #Write-Host $solutionResponse

                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"  # Add date and time

                $errorsListEntry = @{
                    "KB" = $errorCode
                    "error" = $ErrorMessage
                    "errorCategory" = $itClassification
                    "errorResponse" = $errorResponse
                    #"solutionResponse" = $solutionResponse
                    "timestamp" = $timestamp  # Add timestamp
                }

                $errorsList += $errorsListEntry

                # Update the KB register
                $title = Get-OpenAITitleResponse -Error $ErrorMessage

                $kbEntry = @{
                    "KB" = $errorCode
                    "title" = $title
                    "category" = $itClassification
                    "timestamp" = $timestamp
                }

                $kbRegister += $kbEntry

                $kbRegister | ConvertTo-Json | Set-Content $kbRegisterFile

             

                # Check if the directory exists, and create it if not
                if (-not (Test-Path $markdownDirectory -PathType Container)) {
                    New-Item -Path $markdownDirectory -ItemType Directory -Force
                }

                 # Create a Markdown file for the error with the unique reference code (KB_)
        $markdownFileName = "$errorCode.md"
        $markdownFilePath = Join-Path -Path $markdownDirectory -ChildPath $markdownFileName

                $markdownContent = @"
---
title: $title
date: $timestamp
category: $itClassification
reference: $errorCode
---

### Error: $ErrorMessage

**IT Classification:** $itClassification

**Error Response:**
$errorResponse


"@

                $markdownContent | Set-Content -Path $markdownFilePath

                $errorsList | ConvertTo-Json | Set-Content $errorLogFile
            } elseif ($errorResponse) {
                $errorCode = "KB_" + (Get-UniqueRefCode)
                Write-Host "IT Classification: $itClassification" -ForegroundColor Cyan
                Write-Host "Error Response:"
                Write-Host $errorResponse

                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"  # Add date and time

                $errorsListEntry = @{
                    "errorCode" = $errorCode
                    "error" = $ErrorMessage
                    "errorCategory" = $itClassification
                    "errorResponse" = $errorResponse
                    "timestamp" = $timestamp  # Add timestamp
                }

                $errorsList += $errorsListEntry

                # Update the KB register
                $title = Get-OpenAITitleResponse -Error $ErrorMessage

                $kbEntry = @{
                    "KB" = $errorCode
                    "title" = $title
                    "category" = $itClassification
                    "timestamp" = $timestamp
                }

                $kbRegister += $kbEntry

                $kbRegister | ConvertTo-Json | Set-Content $kbRegisterFile

                # Define the directory for Markdown files
                #$markdownDirectory = "C:\temp\demo\ErrorMarkdowns"

             

                # Create a Markdown file for the error with the unique reference code (KB_)
                $markdownFileName = "$errorCode.md"
                $markdownFilePath = Join-Path -Path $markdownDirectory -ChildPath $markdownFileName

                $markdownContent = @"
---
title: $title
date: $timestamp
category: $itClassification
reference: $errorCode
---

### Error: $ErrorMessage

**IT Classification:** $itClassification

**Error Response:**
$errorResponse
"@

                $markdownContent | Set-Content -Path $markdownFilePath

                $errorsList | ConvertTo-Json | Set-Content $errorLogFile
            }
        }
    }

    $mutex.ReleaseMutex()
}

#

# Test the error handling using sample commands
# Set-ExecutionPolicy RemoteSigned  # Commented out to avoid changing the execution policy in the script

# Test 1
Write-Host "Running script:" -ForegroundColor Yellow
$scriptText1 = @'
try {
    $process = Get-Process NonExistentProcess1 -ErrorAction Stop
} catch {
    Handle-Error -ErrorMessage $_.Exception.Message
}
'@
Write-Host $scriptText1
Invoke-Expression -Command $scriptText1

# Test 2
Write-Host "`nRunning script:" -ForegroundColor Yellow
$scriptText2 = @'
try {
    $content = Get-Content C:\NonExistentFile1.txt -ErrorAction Stop
} catch {
    Handle-Error -ErrorMessage $_.Exception.Message
}
'@
Write-Host $scriptText2
Invoke-Expression -Command $scriptText2

# Test 3
Write-Host "`nRunning script:" -ForegroundColor Yellow
$scriptText3 = @'
try {

        $service = Get-Service "NonExistentService1" -ErrorAction Stop

        
} catch {
    Handle-Error -ErrorMessage $_.Exception.Message
}
'@
Write-Host $scriptText3
Invoke-Expression -Command $scriptText3
