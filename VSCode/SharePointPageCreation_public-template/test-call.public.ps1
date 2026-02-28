param(
  [Parameter(Mandatory = $true)]
  [string]$FunctionUrl,

  [Parameter(Mandatory = $false)]
  [string]$PayloadPath = ".\payload.public.json"
)

if (-not (Test-Path $PayloadPath)) {
  throw "Payload file not found: $PayloadPath"
}

$body = Get-Content -Raw $PayloadPath
$resp = Invoke-WebRequest -Method Post -Uri $FunctionUrl -ContentType "application/json" -Body $body -SkipHttpErrorCheck

"Status: $($resp.StatusCode)"
"Content-Type: $($resp.Headers['Content-Type'])"

$rawBody = if ($resp.Content -is [byte[]]) {
  [System.Text.Encoding]::UTF8.GetString($resp.Content)
} else {
  [string]$resp.Content
}

try {
  $json = $rawBody | ConvertFrom-Json
  if ($null -ne $json.success) {
    "success: $($json.success)"
    "pageUrl: $($json.pageUrl)"
    "serverRelativeUrl: $($json.serverRelativeUrl)"
    "message: $($json.message)"
  } elseif ($null -ne $json.error) {
    "error: $($json.error)"
    "message: $($json.message)"
    "correlationId: $($json.correlationId)"
  }
}
catch {
  "Body parse warning: response body is not JSON or parsing failed."
}

"`n--- BODY START ---"
$rawBody
"--- BODY END ---"
