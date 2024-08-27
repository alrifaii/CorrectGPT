Add-Type -AssemblyName System.Runtime.WindowsRuntime
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
$apiKey = "YOUR_API"
$msg = Get-Clipboard
$msg = $msg -replace 'ä', '&auml;'
$msg = $msg -replace 'ö', '&ouml;'
$msg = $msg -replace 'ü', '&uuml;'
$msg = $msg -replace 'ß', '&szlig;'

$prompt = "Correct this for Grammar. Only output the correction, without any comment: $msg"
$requestBody = @{
    "model" = "gpt-3.5-turbo-0125"  
    "messages" = @(
        @{
            "role" = "user"
            "content" = $prompt
        }
    )
}

$jsonRequest = $requestBody | ConvertTo-Json -Depth 4

try {
    $response = Invoke-RestMethod -Uri "https://api.openai.com/v1/chat/completions" `
        -Method Post `
        -Headers @{ "Authorization" = "Bearer $apiKey"; "Content-Type" = "application/json" } `
        -Body $jsonRequest

    $correction = $response.choices[0].message.content
    $tokkens =  $response.usage.total_tokens
    
    $correction = $correction -replace 'Ã¤', 'ä'
    $correction = $correction -replace 'Ã¶', 'ö'
    $correction = $correction -replace 'Ã¼', 'ü'
    $correction = $correction -replace 'Ã.', 'ß'

    if ([string]::IsNullOrEmpty($correction)) {
        "ERROR" | Set-Clipboard
    } else {
        $correction | Set-Clipboard
    }
    $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
    $xml.LoadXml(@"
<toast>
    <visual>
        <binding template='ToastText02'>
            <text id='1'>CorrectGPT</text>
            <text id='2'>Correction copied. Total Tokens: $tokkens</text>
        </binding>
    </visual>
</toast>
"@)

    $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("CorrectGPT").Show($toast)

} catch {


    $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
    $xml.LoadXml(@"
<toast>
    <visual>
        <binding template='ToastText02'>
            <text id='1'>CorrectGPT</text>
            <text id='2'>ERROR: $($_.Exception.Message)</text>
        </binding>
    </visual>
</toast>
"@)

    $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("CorrectGPT").Show($toast)
}
