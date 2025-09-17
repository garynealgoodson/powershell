# Purpose
# Connect to Exchange Online and performs a message trace for every day over the past 90 days to identify the most frequent senders into the tenant, which can then be used to help pre-seed a whitelist when onboarding third party email filtering solution 
# There are limits of the Get-MessageTraceV2 cmdlet!
# Maximum historical depth: 90 days
# Maximum query window per request: 10 days
# Maximum results per query: 5000 (script will not handled more than 5000 emails daily)
# Throttling: 100 queries per 5-minute window 

# Connect to Exchange Online
Connect-ExchangeOnline

# Define date range: last 90 days
$fullStartDate = (Get-Date).AddDays(-90)
$fullEndDate = Get-Date

# Get tenant name
$tenantName = (Get-AcceptedDomain | Where-Object {$_.Default} | Select-Object -ExpandProperty DomainName)

# Get internal domains
$internalDomains = Get-AcceptedDomain | Select-Object -ExpandProperty DomainName

# Initialize array to collect all sender domains
$allSenderDomains = @()

# Loop through each day
$currentStart = $fullStartDate
while ($currentStart -lt $fullEndDate) {
    $currentEnd = $currentStart.AddDays(1)

    Write-Host "Querying from $($currentStart.ToShortDateString()) to $($currentEnd.ToShortDateString())..."

    try {
        $messageTraces = Get-MessageTraceV2 -StartDate $currentStart -EndDate $currentEnd -ResultSize 5000

        $externalSenders = $messageTraces |
        Where-Object {
            $_.SenderAddress -and ($_.SenderAddress -match "@") -and
            ($internalDomains -notcontains ($_.SenderAddress -split "@")[1])
        } |
        Select-Object -ExpandProperty SenderAddress |
        ForEach-Object { ($_ -split "@")[1] }

        $allSenderDomains += $externalSenders
    }
    catch {
        Write-Warning "Failed to query $($currentStart.ToShortDateString()) to $($currentEnd.ToShortDateString()): $_"
    }

    $currentStart = $currentEnd
}

# Group and count domains
$domainCounts = $allSenderDomains |
Group-Object |
Sort-Object Count -Descending |
Select-Object @{Name="Domain";Expression={$_.Name}}, @{Name="Count";Expression={$_.Count}}

# Export to CSV
$csvPath = "c:\Temp\InboundDomains_$tenantName.csv"
$domainCounts | Export-Csv -Path $csvPath -NoTypeInformation

# Disconnect session
Disconnect-ExchangeOnline

