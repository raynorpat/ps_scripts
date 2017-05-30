param (
    [string]$message = "This machine just finished deployment.",
    [string]$title = "Deploy Complete",
    [string]$status = "Finished",
    [string]$location = "Test bench",
    [string]$system = "labsys01"
)

# Edit the URI with the webhook url that was provided
$uri = 'https://outlook.office.com/webhook/0392ef4b-ebbd-4c27-8e7b-bfe3211f8f5d@12f5f98d-602c-4c59-9cc1-ccef754e9148/IncomingWebhook/01f301ae310246a68bf49f0e88ead5c8/1c37de73-2d5d-4ec9-b3b8-ff6300f23df1';

# Build the message Body
$body = ConvertTo-Json -Depth 4 @{
    title = $title
    text = ' '
    sections = @(
        @{
            activityText = $message
        },
        @{
            facts = @(
                @{
                    name = 'Status'
                    value = $status
                },
                @{
                    name = 'System Name'
                    value = $system
                },
                @{
                    name = 'Location'
                    value = $location
                }
            )
        }
    )
}

# Send the message to MS Teams
Invoke-RestMethod -uri $uri -Method Post -body $body -ContentType 'application/json';
Write-Output "INFO - Message has been sent to MS Teams channel.";