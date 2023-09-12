##CONFIG

$APIKEY = "ENTER_API_KEY"
$Endpoint = "https://avdtech.openai.azure.com/openai/deployments/AVDTEST/chat/completions?api-version=2023-07-01-preview"


$PreInformation = @'

Assume you are working in a End User computing team on Azure Virtual Desktop and you are responsible for troubleshooting error codes and producing documentation.

'@


#----------

# Check and create directory if not exist
$directoryPath = "C:\temp\KBCollection"
if (-not (Test-Path $directoryPath)) {
    New-Item -ItemType Directory -Path $directoryPath
}



#FSLOGIX Loading
$FslogixCodes = "$directoryPath\FSLogixErrorCodes.json"

# Specify the URL to download from
$url = "https://raw.githubusercontent.com/RMITBLOG/AI_KB_Generator/main/FSLogixErrorCodes.json"


# Check if the file already exists
if (-not (Test-Path $FslogixCodes)) {
    # File doesn't exist, download it
    Invoke-WebRequest -Uri $url -OutFile $FslogixCodes
    Write-Output "File downloaded successfully to $FslogixCodes"
} else {
    Write-Output "File already exists at $FslogixCodes"
}

$fslogixErrorscodes = Get-Content -Path $FslogixCodes | ConvertFrom-Json



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

function Get-UniqueRefCode {
    # Generate a unique 8-digit reference code starting with "KB_"
    $refCode = "KB_" + (Get-Random -Minimum 10000000 -Maximum 99999999)
    return $refCode
}

function DetermineITClassification {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ErrorMessage
    )

    # Define a list of regex patterns and corresponding classifications
    $patternsAndClassifications = @(
        @{ Pattern = "FSLogix|FSLogix profile|ProfileContainer|OfficeContainer|frx|app attach|disk space|profile storage|FSLogix service|FSLogix logs|FSLogix support|FSLogix VirtualDiskAPI"; Classification = "FSLogix Error" },
        @{ Pattern = "file|directory|path|filesystem|access|permission|disk|storage|backup|restore"; Classification = "File/Storage Error" },
        @{ Pattern = "service|network|communication|connection|port|protocol|firewall|routing"; Classification = "Service/Network Error" },
        @{ Pattern = "process|application|program|software|execution|runtime|crash|hang|freeze|bug|error"; Classification = "Process/Software Error" },
        @{ Pattern = "database|SQL|query|table|record|schema|connection|transaction"; Classification = "Database Error" },
        @{ Pattern = "hardware|device|driver|firmware|motherboard|CPU|RAM|disk|keyboard|mouse|monitor"; Classification = "Hardware Error" },
        @{ Pattern = "authentication|authorization|login|credentials|token|session|security|SSL|TLS"; Classification = "Authentication/Security Error" },
        @{ Pattern = "DNS|domain|hostname|IP address|routing|DNS server|DNS resolution|DNS cache"; Classification = "DNS/Host Resolution Error" },
        @{ Pattern = "email|SMTP|IMAP|POP3|email server|send|receive|mailbox|outlook|thunderbird|email client"; Classification = "Email/Communication Error" },
        @{ Pattern = "website|web server|HTTP|HTTPS|URL|webpage|browser|web application|web service"; Classification = "Website/Web Application Error" },
        @{ Pattern = "printer|printing|print queue|paper jam|printer driver|scanner|fax|printing service"; Classification = "Printer/Printing Error" },
        @{ Pattern = "backup|restore|data loss|corruption|disaster recovery|backup solution"; Classification = "Backup/Restore Error" }
    )

    # Loop through patterns and return classification if a match is found
    foreach ($patternAndClassification in $patternsAndClassifications) {
        if ($ErrorMessage -match $patternAndClassification.Pattern) {
            return $patternAndClassification.Classification
        }
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
            @{ role = "system"; content = $PreInformation  },  # Add pre-information here
            @{ role = "system"; content = $fslogixErrorscodes },  # Add pre-information here
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


     #Write-Output "Sending request body to OpenAI: $body"
  

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

# Check if the error message matches any FSLogix error codes
$fslogixError = $fslogixErrors | Where-Object { $ErrorMessage -match $_.ErrorCode }
if ($fslogixError) {
    # Append FSLogix error information to the pre-information content
    $PreInformation += "`n" + $fslogixError.Description
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

# Test 4 - Simulating error codes 
Write-Host "`nRunning script:" -ForegroundColor Yellow
$scriptText4 = @'
try {
    # Simulating an error for testing purposes
    if ($true) {
        throw "Error encountered with code 0x87D00207"
    }

   
} catch {
    Handle-Error -ErrorMessage $_.Exception.Message
}
'@
Write-Host $scriptText4
Invoke-Expression -Command $scriptText4

# Test 5 - Simulating FSLogix error codes 

# Sample error messages
$errorMessages = @(
    "[07:31:01.429][tid:00000d30.00004c04][ERROR:000000b7]  No Create access: \\fslfileserver\profiles\jdoe-S-1-2-34-000000000-000000000-0000000000-00000000-test (Cannot create a file when that file already exists.)",
    "[07:31:01.432][tid:00000d30.00004c04][ERROR:000000b7]  LoadProfile failed. Version: 2.9.8308.44092 User: jdoe. SID: S-1-2-34-000000000-000000000-0000000000-00000000. SessionId: 11. FrxStatus: 31 (Cannot create a file when that file already exists.)",
    "[11:35:33.734][tid:00000d30.00004f38][ERROR:00000005]  VirtualDiskAPI::CreateFormattedDisk failed to create vhd(x): \\fslfileserver\profiles\jdoe-S-1-2-34-000000000-000000000-0000000000-00000000\Profile_jdoe.vhdx (Access is denied.)",
    "[11:35:59.241][tid:00000d30.00004f38][ERROR:80070003]  Failed to save installed AppxPackages (The system cannot find the path specified.)"
)

# Loop through each error message and pass it to the Handle-Error function
foreach ($errorMessage in $errorMessages) {
    Write-Host "Processing error: $errorMessage"
    Handle-Error -ErrorMessage $errorMessage
    Write-Host "`n" # for better readability between messages
}


# Test 6 - Simulating FSLogix reason codes 

# Sample error messages
$errorMessages = @(
    "[07:26:39.015][tid:00000d30.000044b4][INFO]             Volume name: \\?\Volume{c04ba195-522c-40b8-a894-879025d9b386}\",
    "[07:26:39.015][tid:00000d30.000044b4][INFO]             Reason set to 0: The container is attached",
    "[07:26:39.015][tid:00000d30.000044b4][INFO]             queryFreeDiskSpace returning after 0 milliseconds"
)

# Loop through each error message and pass it to the Handle-Error function
foreach ($errorMessage in $errorMessages) {
    Write-Host "Processing error: $errorMessage"
    Handle-Error -ErrorMessage $errorMessage
    Write-Host "`n" # for better readability between messages
}

}
