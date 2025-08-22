<# This script successfully calls the webhook event on TOPdesk for the automated action TD - Ticket cleanup for each ticket in the export file (csv format required).#> 

#TOPdesk Integration
#set variables
function Connect-TD {
    $customerurl = 'https://topdesk.uottawa.ca'
    #replace "password" with your apiuser password
    $Text = 'apiuser:password'
    $Bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $EncodedAppPass = [Convert]::ToBase64String($Bytes)

    $headers = @{
        'Authorization' = 'BASIC ' + $EncodedAppPass
        'Content-Type'  = 'application/json'
    }
    return $headers
}

# Path to CSV file (turn your excel TOPdesk ticket export file into a csv file)
$CsvPath = "Path to your CSV file"

# TOPdesk Webhook URL (on TOPdesk test)
$WebhookURL = "https://uottawa-test.topdesk.net/services/action-v1/api/webhooks/b60f7fd4-673f-4510-84ed-d8c260c605c5"

$connection = Connect-TD

# Read CSV file
if (-not (Test-Path $CsvPath)) {
    Write-Host "CSV file not found at $CsvPath"
    exit
}

# Calls webhook event in TOPdesk and gives it each ticket number (with conditions that ticket number is not blank)
$tickets = Import-Csv -Path $CsvPath -Encoding UTF8

Write-Host "Found $($tickets.Count) tickets to process..."

foreach ($ticket in $tickets) {
    $ticketId = $ticket.'Ticket Number'

    if (-not $ticketId) {
        Write-Error "Skipping row with no Ticket Number."
        continue
    }

    $payload = @{
        ticketNumber = $ticketId
    }
    
    try {
        $jsonPayload = $payload | ConvertTo-Json -Depth 5 -Compress
        Invoke-RestMethod -Uri $WebhookURL -Method Post -Headers $connection -Body $jsonPayload
        Write-Host "Triggered webhook for ticket $ticketId"

    }
    catch {
        Write-Host "Failed for ticket $ticketId. Error: $($_.Exception.Message)"
    }
}