[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
$apiKey = "YOUR KEY"
$msg = Get-Clipboard
$prompt = "Correct this for Grammar. Only output the correction, without any comment: $msg"
$requestBody = @{
    "model" = "gpt-3.5-turbo-0125"
    "messages" = @(
        @{
            "role" = "system"
            "content" = "You are a helpful assistant."
        },
        @{
            "role" = "user"
            "content" = $prompt
        }
    )
}


$jsonRequest = $requestBody | ConvertTo-Json -Depth 4


$response = Invoke-RestMethod -Uri "https://api.openai.com/v1/chat/completions" `
    -Method Post `
    -Headers @{ "Authorization" = "Bearer $apiKey"; "Content-Type" = "application/json" } `
    -Body $jsonRequest


$correction = $response.choices[0].message.content
$correction | Set-Clipboard
[console]::beep(349,100)
[console]::beep(523,100)
$xml = New-Object Windows.Data.Xml.Dom.XmlDocument
$xml.LoadXml(
@"
<toast><visual><binding template="ToastText02"><text id="2">Corrected</text></binding></visual></toast>
"@)
$toast = New-Object Windows.UI.Notifications.ToastNotification $xml
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("CorrectGPT").Show($toast)