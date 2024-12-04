# =============================================================================
# Script: Get-MessageCenter.ps1
# Purpose: Retrieves and processes Microsoft 365 Message Center announcements
# Author: Cengiz YILMAZ - Microsoft MVP
# Web: https://yilmazcengiz.tr
# =============================================================================

param(
    [Parameter(Mandatory=$false)]
    [string]$GraphSecret
)

# Configuration settings
$SCRIPT_CONFIG = @{
    # Paths
    DataPath = "./@data"
    ConfigFile = "./@build/config-m365.json"
    SecretsFile = "./@build/secrets-m365.json"
    ArchivePath = "./@data/archive"
    
    # API Settings
    MaxMessages = 999
    SortField = "LastModifiedDateTime"
    SortDirection = "desc"
    
    # Message Processing
    TitleCleanupPattern = '\(Updated\)\s*'
    ContentEnhancementPattern = '\[(.*?)\]'
    ContentReplacement = '<b>$1</b>'
}

# Function to establish connection with Microsoft Graph
function Initialize-GraphConnection {
    param($secretValue)
    
    try {
        # Load configuration
        $m365Settings = Get-Content $SCRIPT_CONFIG.ConfigFile | ConvertFrom-Json
        
        # Handle secret management
        $graphCredential = if ($secretValue) { 
            $secretValue 
        } else { 
            Write-Host "Loading secret from configuration file..."
            (Get-Content $SCRIPT_CONFIG.SecretsFile | ConvertFrom-Json).clientSecret
        }
        
        # Create secure credentials
        [securestring]$secureGraphSecret = ConvertTo-SecureString $graphCredential -AsPlainText -Force
        [pscredential]$graphAuthentication = New-Object System.Management.Automation.PSCredential (
            $m365Settings.clientId, 
            $secureGraphSecret
        )
        
        Write-Host "Initiating Microsoft Graph connection..."
        Connect-MgGraph -TenantId $m365Settings.tenantId -Credential $graphAuthentication -NoWelcome
        Write-Host "Graph connection established successfully"
    }
    catch {
        Write-Error "Graph connection failed: $_"
        throw
    }
}

# Function to retrieve Message Center announcements
function Get-MessageCenterAnnouncements {
    try {
        Write-Host "Fetching Message Center announcements..."
        $announcements = Get-MgServiceAnnouncementMessage -Top $SCRIPT_CONFIG.MaxMessages `
            -Sort "$($SCRIPT_CONFIG.SortField) $($SCRIPT_CONFIG.SortDirection)" -All
        Write-Host "Successfully retrieved $($announcements.Count) announcements"
        return $announcements
    }
    catch {
        Write-Error "Failed to retrieve announcements: $_"
        throw
    }
}

# Function to calculate message counts per service
function Get-ServiceStatistics {
    param([Array]$messages)
    
    $serviceStats = @{}
    
    foreach($message in $messages) {
        foreach($service in $message.Services) {
            if($serviceStats.ContainsKey($service)) {
                $serviceStats[$service]++
            } else {
                $serviceStats[$service] = 1
            }
        }
    }
    
    return $serviceStats
}

# Function to process and format message content
function Format-MessageContent {
    param([object]$message)
    
    $message.Title = $message.Title -replace $SCRIPT_CONFIG.TitleCleanupPattern, ''
    $message.body.content = $message.body.content -replace $SCRIPT_CONFIG.ContentEnhancementPattern, $SCRIPT_CONFIG.ContentReplacement
    
    # Add timestamp for tracking
    $message | Add-Member -NotePropertyName "ProcessedTimestamp" -NotePropertyValue (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") -Force
    
    return $message
}

# Function to ensure required directories exist
function Initialize-DirectoryStructure {
    @($SCRIPT_CONFIG.DataPath, $SCRIPT_CONFIG.ArchivePath) | ForEach-Object {
        if(-not (Test-Path $_)) {
            New-Item -ItemType Directory -Path $_ -Force | Out-Null
            Write-Host "Created directory: $_"
        }
    }
}

# Main execution block
try {
    # Initialize directory structure
    Initialize-DirectoryStructure
    
    # Initialize connection and retrieve messages
    Initialize-GraphConnection -secretValue $GraphSecret
    $messageItems = Get-MessageCenterAnnouncements
    
    Write-Host "Processing $($messageItems.Count) messages..."
    
    # Process each message
    foreach($messageItem in $messageItems) {
        $processedMessage = Format-MessageContent -message $messageItem
        
        # Archive individual messages with error handling
        try {
            $archivePath = Join-Path $SCRIPT_CONFIG.ArchivePath "$($processedMessage.Id).json"
            $processedMessage | ConvertTo-Json -Depth 10 | Set-Content -Path $archivePath
        }
        catch {
            Write-Warning "Failed to archive message $($processedMessage.Id): $_"
        }
    }
    
    # Save all messages to main JSON file
    $messagesPath = Join-Path $SCRIPT_CONFIG.DataPath "messages.json"
    $messageItems | ConvertTo-Json -Depth 10 | Set-Content -Path $messagesPath
    
    # Calculate and save service statistics
    $serviceStatistics = Get-ServiceStatistics -messages $messageItems
    $serviceReport = @()
    
    foreach($service in ($messageItems.Services | Sort-Object | Get-Unique)) {
        $serviceReport += @{
            "serviceName" = $service
            "messageCount" = $serviceStatistics[$service]
            "lastUpdated" = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            "averageMessagesPerDay" = [math]::Round($serviceStatistics[$service] / 30, 2)
        }
    }
    
    # Save service statistics with error handling
    try {
        $servicesPath = Join-Path $SCRIPT_CONFIG.DataPath "services.json"
        $serviceReport | ConvertTo-Json | Set-Content -Path $servicesPath
    }
    catch {
        Write-Error "Failed to save service statistics: $_"
    }
    
    # Output summary
    Write-Host "Message Center update completed successfully"
    Write-Host "Total services processed: $($serviceReport.Count)"
    Write-Host "Total messages processed: $($messageItems.Count)"
    Write-Host "Data saved to: $($SCRIPT_CONFIG.DataPath)"
}
catch {
    Write-Error "Script execution failed: $_"
    throw
}
finally {
    # Clean up connection
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "Graph connection closed successfully"
    }
    catch {
        Write-Warning "Failed to disconnect from Graph: $_"
    }
}
