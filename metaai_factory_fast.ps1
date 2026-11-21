param(
    [string]$Serial = $env:ADB_SERIAL,
    [string]$AdbPath = "C:\platform-tools\adb.exe",
    [int]$UserId = -1,
    [string]$Password = $env:METAAI_DEFAULT_PASSWORD,
    [int]$YearMin = 1990,
    [int]$YearMax = 2000,
    [int]$EmailPollSeconds = 150,
    [int]$Count = 10,
    [string]$CredFile = (Join-Path $PSScriptRoot "output\MetaAI_Credentials.xlsx"),
    [string]$MailCxToken = "",
    [switch]$SweepInboxes = $false,
    [switch]$KeepLog = $false,
    [string]$ProgressFile = ""
)

$ErrorActionPreference = "Continue"
$script:StartOfBatch = Get-Date

if (-not (Test-Path $AdbPath)) {
    $foundAdb = (Get-Command adb -ErrorAction SilentlyContinue).Source
    if ($foundAdb) {
        $AdbPath = $foundAdb
    } else {
        $candidatePaths = @(
            "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\Google.PlatformTools_Microsoft.Winget.Source_8wekyb3d8bbwe\platform-tools\adb.exe",
            "$env:LOCALAPPDATA\Android\Sdk\platform-tools\adb.exe"
        )
        foreach ($cp in $candidatePaths) {
            if (Test-Path $cp) { $AdbPath = $cp; break }
        }
    }
}

$Pkg       = "com.facebook.stella"

# --- Mail.cx provider ---
$ApiBase   = "https://mail.cx"
$script:ClientId = [System.Guid]::NewGuid().ToString()
$script:SystemDomains = @("uqu.me", "9k3r.com", "ddker.com")

function MailCxHeaders {
    $h = @{
        "X-Client-ID" = $script:ClientId
        "Accept"      = "application/json"
        "User-Agent"  = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    if ($MailCxToken) {
        $h["x-api-token"] = $MailCxToken
    }
    return $h
}

$TempDir   = Join-Path ([System.IO.Path]::GetTempPath()) "opencode"
if (-not (Test-Path $TempDir)) { [void][System.IO.Directory]::CreateDirectory($TempDir) }
$DumpRemote = "/sdcard/metaai_ui.xml"
$DumpLocal = Join-Path $TempDir "metaai_ui.xml"
$script:XmlCache  = $null
$ActiveSerial = ""
$LogLines = @()

function Log([string]$m) {
    $script:LogLines += "[$(Get-Date -Format HH:mm:ss)] $m"
}

function Write-Step([string]$m) { Write-Host "[$(Get-Date -Format HH:mm:ss)] $m" -ForegroundColor Cyan }
function Write-Ok([string]$m)   { Write-Host "  -> $m" -ForegroundColor Green }
function Write-Err([string]$m)  { Write-Host "  !! $m" -ForegroundColor Red }
function Ms($ms) { Start-Sleep -Milliseconds $ms }

function Adb {
    param([string[]]$Cmd)
    & $AdbPath -s $ActiveSerial @Cmd
}

function New-Inbox {
    $letters = 'abcdefghijklmnopqrstuvwxyz'.ToCharArray()
    $name = (-join ($letters | Get-Random -Count 3)) + "cx" + (-join ($letters | Get-Random -Count 4))
    $domain = $script:SystemDomains[0] # "uqu.me"
    $address = "$name@$domain"
    return @{
        alias   = $name
        address = $address
        domain  = $domain
    }
}

function Delete-Inbox([string]$Address) {
    try {
        Invoke-WebRequest -Uri "$ApiBase/v1/inbox/$Address" -Headers (MailCxHeaders) -Method DELETE -UseBasicParsing | Out-Null
        Write-Ok "Cleaned mail.cx inbox $Address"
    } catch { }
}

function Get-MailOtp([string]$Address) {
    try {
        $url = "$ApiBase/v1/inbox/$Address"
        $resp = Invoke-WebRequest -Uri $url -Headers (MailCxHeaders) -Method GET -TimeoutSec 30 -UseBasicParsing

        # 204 No Content means long-poll held for 25s with no new mail yet
        if ($resp.StatusCode -eq 204) {
            return $null
        }

        if ($resp.StatusCode -eq 200 -and $resp.Content) {
            $data = $resp.Content | ConvertFrom-Json
            if ($data.emails -and $data.emails.Count -gt 0) {
                $emailItem = $data.emails[0]
                Write-Ok "Email received from $($emailItem.from_email): '$($emailItem.subject)'"

                # 1. Check subject / preview text first
                if ($emailItem.subject -match '\b([0-9]{6})\b') {
                    return $Matches[1]
                }
                if ($emailItem.preview_text -match '\b([0-9]{6})\b') {
                    return $Matches[1]
                }

                # 2. Fetch full email message
                $fullUrl = "$ApiBase/v1/email/$($emailItem.id)"
                $emailObj = Invoke-RestMethod -Uri $fullUrl -Headers (MailCxHeaders) -Method GET

                # Check full subject
                if ($emailObj.subject -match '\b([0-9]{6})\b') {
                    return $Matches[1]
                }

                # Check plain text body (Mail.cx uses text_body)
                $rawText = if ($emailObj.text_body) { $emailObj.text_body } elseif ($emailObj.text) { $emailObj.text } else { "" }
                if ($rawText) {
                    if ($rawText -match '(?i)confirmation\s*code[^\d]{0,50}?([0-9]{6})') { return $Matches[1] }
                    if ($rawText -match '(?i)\bcode[^\d]{1,30}?([0-9]{6})') { return $Matches[1] }
                    if ($rawText -match '\b([0-9]{6})\b') { return $Matches[1] }
                }

                # Check HTML body (Mail.cx uses html_body)
                $rawHtml = if ($emailObj.html_body) { $emailObj.html_body } elseif ($emailObj.html) { $emailObj.html } else { "" }
                if ($rawHtml) {
                    $clean = ($rawHtml -replace '(?s)<style[^>]*>.*?</style>', ' ' `
                                       -replace '(?s)<script[^>]*>.*?</script>', ' ' `
                                       -replace '(?s)<[^>]+>', ' ' `
                                       -replace '&nbsp;', ' ' `
                                       -replace '&amp;', '&' `
                                       -replace '&#8203;', '' `
                                       -replace '&zwnj;', '' `
                                       -replace '\s+', ' ').Trim()

                    if ($clean -match '(?i)confirmation\s*code[^\d]{0,50}?([0-9]{6})') { return $Matches[1] }
                    if ($rawHtml -match '(?i)>\s*([0-9]{6})\s*<') { return $Matches[1] }
                    if ($clean -match '(?i)\bcode[^\d]{1,30}?([0-9]{6})') { return $Matches[1] }
                    $digits = [regex]::Matches($clean, '\b([0-9]{6})\b')
                    if ($digits.Count -gt 0) { return $digits[0].Groups[1].Value }
                }
            }
        }
    } catch {
        Write-Step "  (mail.cx waiting: $_)"
    }
    return $null
}

function XE([string]$s) { $s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace('"', "&quot;").Replace("'", "&apos;") }

function Read-XlsxRows([string]$Path) {
    $rows = @()
    if (-not (Test-Path $Path)) { return $rows }
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $z = $null
    try {
        $z = [System.IO.Compression.ZipFile]::OpenRead($Path)
        $e = $z.GetEntry("xl/worksheets/sheet1.xml")
        if (-not $e) { return $rows }
        $r = New-Object System.IO.StreamReader($e.Open())
        $xml = $r.ReadToEnd()
        $r.Close()
    } catch { Write-Err "could not read existing xlsx: $_"; return $rows }
    finally { if ($z) { $z.Dispose() } }
    $rowMatches = [regex]::Matches($xml, '<row[^>]*>(.*?)</row>', 'Singleline')
    foreach ($rm in $rowMatches) {
        $cells = @([regex]::Matches($rm.Groups[1].Value, '<t>(.*?)</t>', 'Singleline') | ForEach-Object {
            $_.Groups[1].Value -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>' -replace '&quot;', '"' -replace '&apos;', "'"
        })
        if ($cells.Count -ge 1) { $rows += ,@($cells) }
    }
    return $rows
}

function Save-CredentialsXlsx([string[]]$Emails) {
    $rows = @()
    if (Test-Path $CredFile) {
        $existing = @(Read-XlsxRows $CredFile)
        $startIdx = 1
        if ($existing.Count -gt 0 -and $existing[0][0] -eq "Email") { $startIdx = 1 } else { $startIdx = 0 }
        for ($i = $startIdx; $i -lt $existing.Count; $i++) { $rows += ,@($existing[$i]) }
    }
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm"
    foreach ($em in $Emails) { $rows += ,@($em, $Password, $stamp) }

    $tempZip = Join-Path $TempDir "xlsx_factory.xlsx"
    if (-not (Test-Path $TempDir)) { [void][System.IO.Directory]::CreateDirectory($TempDir) }
    if (Test-Path $tempZip) { Remove-Item $tempZip -Force }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
    [void]$sb.Append('<row r="1"><c r="A1" t="inlineStr"><is><t>Email</t></is></c><c r="B1" t="inlineStr"><is><t>Password</t></is></c><c r="C1" t="inlineStr"><is><t>Created</t></is></c></row>')
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $r = $i + 2
        [void]$sb.Append("<row r=`"$r`"><c r=`"A$r`" t=`"inlineStr`"><is><t>$(XE $rows[$i][0])</t></is></c><c r=`"B$r`" t=`"inlineStr`"><is><t>$(XE $rows[$i][1])</t></is></c><c r=`"C$r`" t=`"inlineStr`"><is><t>$(XE $rows[$i][2])</t></is></c></row>")
    }
    [void]$sb.Append('</sheetData></worksheet>')
    $sheetXml = $sb.ToString()

    $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
    $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    $workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Credentials" sheetId="1" r:id="rId1"/></sheets></workbook>'
    $workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($tempZip, [System.IO.Compression.ZipArchiveMode]::Create)
    function Add-ZipPart($archive, [string]$name, [string]$content) {
        $entry = $archive.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
        $writer = New-Object System.IO.StreamWriter($entry.Open()); $writer.Write($content); $writer.Close()
    }
    Add-ZipPart $zip "[Content_Types].xml" $contentTypes
    Add-ZipPart $zip "_rels/.rels" $rels
    Add-ZipPart $zip "xl/workbook.xml" $workbook
    Add-ZipPart $zip "xl/_rels/workbook.xml.rels" $workbookRels
    Add-ZipPart $zip "xl/worksheets/sheet1.xml" $sheetXml
    $zip.Dispose()

    $credDir = [System.IO.Path]::GetDirectoryName($CredFile)
    if ($credDir -and -not (Test-Path $credDir)) { [void][System.IO.Directory]::CreateDirectory($credDir) }
    Copy-Item $tempZip $CredFile -Force
    Write-Ok "saved $($Emails.Count) credentials to $CredFile"
}

function Dump-Xml {
    param([switch]$Fresh)
    if (-not $Fresh -and $script:XmlCache) { return $script:XmlCache }
    $xml = (& $AdbPath -s $ActiveSerial exec-out uiautomator dump /dev/tty 2>$null) -join "`n"
    if ($xml -and $xml.Length -gt 500) { $script:XmlCache = $xml; return $xml }
    for ($i = 0; $i -lt 2; $i++) {
        $xml = (& $AdbPath -s $ActiveSerial exec-out uiautomator dump /dev/tty 2>$null) -join "`n"
        if ($xml -and $xml.Length -gt 500) { $script:XmlCache = $xml; return $xml }
    }
    return $null
}

function Dump-Fresh { Dump-Xml -Fresh }

function Get-NodesFromXml([string]$Xml) {
    $result = @()
    if (-not $Xml) { return $result }
    $nodeMatches = [regex]::Matches($Xml, '<node([^>]*)>')
    foreach ($nm in $nodeMatches) {
        $attrs = $nm.Groups[1].Value
        $a = @{}
        foreach ($attrMatch in [regex]::Matches($attrs, '([\w-]+)="([^"]*)"')) {
            $a[$attrMatch.Groups[1].Value] = $attrMatch.Groups[2].Value
        }
        $bounds = ''
        if ($a.ContainsKey('bounds')) { $bounds = $a['bounds'] }
        if ($bounds -match '\[(\d+),(\d+)\]\[(\d+),(\d+)\]') {
            $result += @{ text = $a['text']; desc = $a['content-desc']; cls = $a['class']
                         x1 = [int]$Matches[1]; y1 = [int]$Matches[2]; x2 = [int]$Matches[3]; y2 = [int]$Matches[4] }
        }
    }
    return $result
}

function Find-Nodes {
    param([string]$xml, [string]$Text = "", [string]$Class = "", [int]$MinY = -1, [int]$MaxY = 10000)
    $all = @(Get-NodesFromXml $xml)
    $filtered = @($all | Where-Object {
        ($Text -eq "" -or $_.text -like "*$Text*" -or $_.desc -like "*$Text*") -and
        ($Class -eq "" -or $_.cls -eq $Class) -and
        $_.y1 -ge $MinY -and $_.y2 -le $MaxY
    })
    return $filtered
}

function TapNode($node) {
    $cx = [int](($node.x1 + $node.x2) / 2)
    $cy = [int](($node.y1 + $node.y2) / 2)
    & $AdbPath -s $ActiveSerial shell input tap $cx $cy | Out-Null
    $script:XmlCache = $null
}

function Wait-ForNode {
    param([string[]]$Text, [int]$TimeoutSeconds = 30, [string]$Class = "", [switch]$NoCache)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $xml = Dump-Xml -Fresh
        $all = @()
        if ($xml) { $all = @(Get-NodesFromXml $xml) }
        if ($Class) { $all = @($all | Where-Object { $_.cls -eq $Class }) }
        foreach ($t in $Text) {
            $exact = @($all | Where-Object { $_.text -eq $t -or $_.desc -eq $t })
            if ($exact.Count -gt 0) { return $exact[0] }
        }
        foreach ($t in $Text) {
            $fuzzy = @($all | Where-Object { $_.text -like "*$t*" -or $_.desc -like "*$t*" })
            if ($fuzzy.Count -gt 0) { return $fuzzy[0] }
        }
    }
    throw "Timeout waiting for element: '$($Text -join "' or '")'"
}

function Wait-For-Focus([int]$TimeoutSeconds = 6) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $x = Dump-Xml -Fresh
        if ($x -and $x -match 'focused="true"') { return $true }
    }
    return $false
}

function TypeText([string]$text) {
    $escaped = $text -replace '%', '%25' -replace ' ', '%s'
    $escaped = $escaped -replace '([&|()<>;$`"''\\])', '\$1'
    & $AdbPath -s $ActiveSerial shell "input text $escaped" | Out-Null
    $script:XmlCache = $null
}

function Clear-AppData {
    & $AdbPath -s $ActiveSerial shell am force-stop --user $UserId $Pkg | Out-Null
    & $AdbPath -s $ActiveSerial shell pm clear --user $UserId $Pkg | Out-Null
    $script:XmlCache = $null
}

function Launch-App {
    & $AdbPath -s $ActiveSerial shell am start --user $UserId -n "$Pkg/com.facebook.wearable.companion.silverstone.main.view.SilverstoneMainActivity" | Out-Null
    $script:XmlCache = $null
}

function Get-WheelValues {
    for ($try = 0; $try -lt 3; $try++) {
        $all = @(Find-Nodes (Dump-Fresh) "" "android.widget.EditText")
        $monthN = @($all | Where-Object { $_.x1 -ge 200 -and $_.x1 -le 400 })[0]
        $dayN   = @($all | Where-Object { $_.x1 -ge 420 -and $_.x1 -le 620 })[0]
        $yearN  = @($all | Where-Object { $_.x1 -ge 640 -and $_.x1 -le 840 })[0]
        if ($monthN -and $dayN -and $yearN) {
            $month = @{ Jan = 1; Feb = 2; Mar = 3; Apr = 4; May = 5; Jun = 6; Jul = 7; Aug = 8; Sep = 9; Oct = 10; Nov = 11; Dec = 12 }[$monthN.text]
            if ($month) {
                return @{ Month = $month; Day = [int]$dayN.text; Year = [int]$yearN.text; MonthText = $monthN.text; DayText = $dayN.text; YearText = $yearN.text }
            }
        }
    }
    throw "could not read all three wheel columns"
}

function Set-Birthday {
    $yearT  = Get-Random -Minimum $YearMin -Maximum ($YearMax + 1)
    Write-Step "Birthday picker: type year $yearT (month/day stay default)"
    $yearN = $null
    $dlB = (Get-Date).AddSeconds(10)
    while (-not $yearN -and (Get-Date) -lt $dlB) {
        $nodes = @(Find-Nodes (Dump-Fresh) "" "android.widget.EditText")
        $yearN  = @($nodes | Where-Object { $_.x1 -ge 640 -and $_.x1 -le 840 })[0]
    }
    if (-not $yearN) { throw "year field not found" }
    TapNode $yearN
    $null = Wait-For-Focus 5
    TypeText "$yearT"

    $cur = Get-WheelValues
    if ($cur.Year -ne $yearT) {
        Write-Err "year did not stick (now $($cur.Year)); retrying"
        $nodes = @(Find-Nodes (Dump-Xml) "" "android.widget.EditText")
        $yearN  = @($nodes | Where-Object { $_.x1 -ge 640 -and $_.x1 -le 840 })[0]
        if (-not $yearN) { throw "year field missing on retry" }
        TapNode $yearN
        & $AdbPath -s $ActiveSerial shell input keyevent 123 | Out-Null
        for ($d = 0; $d -lt 4; $d++) { & $AdbPath -s $ActiveSerial shell input keyevent 67 | Out-Null }
        $script:XmlCache = $null
        TypeText "$yearT"

        $cur = Get-WheelValues
        if ($cur.Year -ne $yearT) { throw "could not set year to $yearT (now $($cur.Year))" }
    }
    $set = @(Find-Nodes (Dump-Xml) "" "android.widget.Button" | Where-Object { $_.text -eq "SET" })
    if ($set.Count -eq 0) { throw "SET button not found" }
    TapNode $set[0]

    $ok = Find-Nodes (Dump-Xml) "Birthday"
    if ($ok.Count -eq 0) { throw "birthday not applied" }
}

function Is-CodeScreen([string]$sx) {
    if (-not $sx) { return $false }
    return ($sx -match 'Confirm your account' -or `
            $sx -match 'enter the 6-digit code' -or `
            $sx -match 'Enter the confirmation code' -or `
            $sx -match 'Send code again')
}

# ============================ ONE ACCOUNT ============================
function Create-One-Account([int]$accountNo) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $email = ""
    $mailAddress = ""
    try {
        Write-Step "Preparing Mail.cx address"
        $inbox = New-Inbox
        $email = $inbox.address
        $mailAddress = $inbox.address
        Write-Ok "email: $email"

        Clear-AppData
        Launch-App

        Write-Step "Welcome screen -> tap 'Set up a device'"
        $tapped = $false
        for ($tryW = 0; $tryW -lt 3 -and -not $tapped; $tryW++) {
            try {
                $welcome = Wait-ForNode @("Set up a device", "Get started with Meta AI") 15
                Write-Ok "found welcome node '$($welcome.text)' bounds=[$($welcome.x1),$($welcome.y1)][$($welcome.x2),$($welcome.y2)]"
                TapNode $welcome

                $probe = Wait-ForNode @("Next", "Continue") 40
                if ($probe) { $tapped = $true; break }
            } catch { }
            Write-Ok "welcome tap retry $($tryW + 1)"
        }
        if (-not $tapped) { throw "could not advance past welcome screen" }

        Write-Step "Enter email"
        $edits = @(Find-Nodes (Dump-Xml) "" "android.widget.EditText" | Where-Object { $_.y1 -gt 900 -and $_.y1 -lt 1400 })
        if ($edits.Count -eq 0) { $edits = @(Find-Nodes (Dump-Xml) "" "android.widget.EditText") }
        if ($edits.Count -eq 0) { throw "email input not found" }
        TapNode $edits[0]
        $null = Wait-For-Focus 6
        TypeText $email
        $nextBtn = Wait-ForNode @("Next", "Continue") 8
        TapNode $nextBtn

        Write-Step "Entering birthday"
        Set-Birthday
        Write-Step "Birthday set"

        Write-Step "Password"
        $passNodes = @(Find-Nodes (Dump-Xml) "" "android.widget.EditText" | Where-Object { $_.y1 -gt 800 -and $_.y1 -lt 1150 })
        if ($passNodes.Count -eq 0) { throw "password field not found" }
        TapNode $passNodes[0]
        $null = Wait-For-Focus 6
        TypeText $Password
        & $AdbPath -s $ActiveSerial shell input keyevent 4 | Out-Null
        $script:XmlCache = $null

        $next2 = Wait-ForNode @("Continue", "Next") 10
        TapNode $next2

        Write-Step "Confirm account details"
        $confirm = Wait-ForNode @("Confirm", "Continue", "Next") 12
        TapNode $confirm

        Write-Step "Waiting for Meta OTP via Mail.cx long-poll (parallel with code screen)..."
        $code = $null
        $screenReady = $false
        $pollDeadline = (Get-Date).AddSeconds($EmailPollSeconds)
        $noCodeChecks = 0
        $lastResendTime = [DateTime]::MinValue

        while ((Get-Date) -lt $pollDeadline -and (-not $code -or -not $screenReady)) {
            if (-not $code) {
                $code = Get-MailOtp $mailAddress
                if (-not $code) {
                    $noCodeChecks++
                }
            }
            if (-not $screenReady) {
                $sx = Dump-Fresh
                if (Is-CodeScreen $sx) { 
                    $screenReady = $true 
                    Write-Ok "confirmation code screen ready"
                }
            }

            # If code not arrived after 2 checks and screen is ready, tap 'Send code again' every 5s
            if ($screenReady -and (-not $code) -and ($noCodeChecks -ge 2)) {
                $now = Get-Date
                if (($now - $lastResendTime).TotalSeconds -ge 5) {
                    $sx = Dump-Xml
                    $sendAgain = @(Find-Nodes $sx "Send code again")
                    if ($sendAgain.Count -gt 0) {
                        Write-Step "  Code not arrived after $noCodeChecks checks -> Tapping 'Send code again'..."
                        TapNode $sendAgain[0]
                        $lastResendTime = Get-Date
                    } else {
                        Write-Step "  Code not arrived after $noCodeChecks checks -> Tapping 'Send code again' (center bounds)..."
                        & $AdbPath -s $ActiveSerial shell input tap 540 1157 | Out-Null
                        $script:XmlCache = $null
                        $lastResendTime = Get-Date
                    }
                }
            }

            if (-not $code) {
                # 3s breather between long-poll reconnects to stay well below rate limit
                Start-Sleep -Seconds 3
            }
        }
        if (-not $code) { throw "No OTP received within $EmailPollSeconds s" }
        if (-not $screenReady) { 
            $null = Wait-ForNode @("Confirm your account", "enter the 6-digit code", "Enter the confirmation code", "Send code again") 15 
        }
        Write-Ok "OTP: $code"

        Write-Step "Enter OTP"
        $otpDone = $false
        for ($k = 0; $k -lt 3 -and -not $otpDone; $k++) {
            if ($k -gt 0) {
                Write-Step "  retyping OTP (attempt $($k + 1))"
                & $AdbPath -s $ActiveSerial shell input keyevent 123 | Out-Null
                1..8 | ForEach-Object { & $AdbPath -s $ActiveSerial shell input keyevent 67 | Out-Null }
                $script:XmlCache = $null
            }
            $field = $null
            $dl = (Get-Date).AddSeconds(6)
            while (-not $field -and (Get-Date) -lt $dl) {
                $fx = Dump-Fresh
                $field = @(Get-NodesFromXml $fx | Where-Object { $_.cls -eq "android.widget.EditText" } | Sort-Object y1)[0]
            }
            if ($field) { TapNode $field } else { & $AdbPath -s $ActiveSerial shell input tap 540 656 | Out-Null }
            $null = Wait-For-Focus 6
            TypeText $code

            $next3 = $null
            $dl3 = (Get-Date).AddSeconds(12)
            while ((Get-Date) -lt $dl3 -and -not $otpDone) {
                $sx = Dump-Fresh
                if (-not $sx) { continue }
                if (-not (Is-CodeScreen $sx)) {
                    Write-Step "  code screen already advanced after typing"
                    $otpDone = $true
                    break
                }
                if (-not $next3) { $next3 = @(Get-NodesFromXml $sx | Where-Object { $_.text -eq "Continue" -or $_.text -eq "Next" })[0] }
                if ($next3) {
                    TapNode $next3
                    $otpDone = $true
                    break
                }
            }
        }
        if (-not $otpDone) { throw "could not submit OTP" }

        Write-Ok "OTP accepted - account registered"
        try { Save-CredentialsXlsx @($email) }
        catch { Write-Err "could not save credentials for ${email}: $_" }
        $elapsed = [int]$sw.Elapsed.TotalSeconds
        Write-Ok "ACCOUNT $accountNo took ${elapsed}s"
        return $true
    } catch {
        Write-Err "account FAILED: $_"
        return $false
    } finally {
        if ($mailAddress) {
            try { Delete-Inbox $mailAddress } catch { }
        }
    }
}

function Select-Device {
    $lines = & $AdbPath devices
    $attached = @{}
    foreach ($l in $lines) {
        if ($l -match '^(\S+)\s+(\S+)$') {
            $attached[$Matches[1]] = $Matches[2]
        }
    }

    $target = $Serial
    if ($target -and $attached.ContainsKey($target)) {
        if ($attached[$target] -eq "device") { return $target }
        if ($attached[$target] -eq "unauthorized") {
            throw "Device '$target' is unauthorized. Unlock phone and tap 'Allow USB debugging'."
        }
        throw "Device '$target' state is '$($attached[$target])' (expected 'device')."
    }

    # Auto-detect real authorized device
    $found = @($attached.Keys | Where-Object { $_ -notmatch 'emulator' -and $attached[$_] -eq 'device' })
    if ($found.Count -ge 1) { return $found[0] }

    $unauth = @($attached.Keys | Where-Object { $attached[$_] -eq 'unauthorized' })
    if ($unauth.Count -gt 0) {
        throw "Connected device ($($unauth -join ', ')) is unauthorized. Unlock phone and tap 'Allow USB debugging'."
    }

    if ($attached.Count -eq 0) {
        throw "No Android devices connected via ADB. Please plug in your phone with USB debugging enabled."
    }

    throw "No authorized Android device found. Attached: $($attached.Keys -join ', ')"
}

function Select-AndroidUser {
    Write-Step "Detecting Android users on device $ActiveSerial..."
    $userLines = Adb @("shell", "pm", "list", "users")
    $users = @()
    foreach ($line in $userLines) {
        if ($line -match 'UserInfo\{(\d+):([^:]*):[^\}]*\}(.*)') {
            $uId = [int]$Matches[1]
            $uName = $Matches[2].Trim()
            $uStatus = $Matches[3].Trim()
            $users += [PSCustomObject]@{
                Id = $uId
                Name = $uName
                Status = $uStatus
            }
        }
    }

    if ($users.Count -eq 0) {
        Write-Err "Could not detect any users via 'pm list users'. Defaulting to User 0."
        return 0
    }

    Write-Host ""
    Write-Host "--- Available Android Users on $ActiveSerial ---" -ForegroundColor Cyan
    $defaultIndex = 0
    for ($i = 0; $i -lt $users.Count; $i++) {
        $u = $users[$i]
        $statusStr = if ($u.Status) { " ($($u.Status))" } else { "" }
        if ($u.Id -eq 0) { $defaultIndex = $i }
        $mark = if ($u.Id -eq 0) { " [Recommended / Default]" } else { "" }
        Write-Host ("  [{0}] User ID: {1,3}  Name: {2,-15}{3}{4}" -f ($i + 1), $u.Id, $u.Name, $statusStr, $mark) -ForegroundColor Yellow
    }
    $defaultUser = $users[$defaultIndex]

    Write-Host ""
    $prompt = "Select user [1-$($users.Count) or enter User ID directly] (default: User $($defaultUser.Id) - $($defaultUser.Name)): "
    $choice = $null
    try {
        $choice = Read-Host $prompt
    } catch { }
    if ([string]::IsNullOrWhiteSpace($choice)) {
        try {
            $choice = [Console]::In.ReadLine()
        } catch { }
    }
    if ($choice) { $choice = $choice.Trim() }

    $selectedId = $null
    if ([string]::IsNullOrWhiteSpace($choice)) {
        $selectedId = $defaultUser.Id
    } elseif ($choice -match '^\d+$') {
        $num = [int]$choice
        if ($num -ge 1 -and $num -le $users.Count) {
            $selectedId = $users[$num - 1].Id
        } else {
            $matched = @($users | Where-Object { $_.Id -eq $num })
            if ($matched.Count -gt 0) {
                $selectedId = $matched[0].Id
            }
        }
    }

    if ($null -eq $selectedId) {
        Write-Err "Invalid choice '$choice'. Falling back to User $($defaultUser.Id) ($($defaultUser.Name))."
        $selectedId = $defaultUser.Id
    }

    $chosenUser = @($users | Where-Object { $_.Id -eq $selectedId })[0]
    Write-Ok "Selected User ID: $selectedId ($($chosenUser.Name))"

    if ($chosenUser.Status -notmatch 'running') {
        Write-Step "Waking up / starting User $selectedId..."
        try { Adb @("shell", "am", "start-user", "$selectedId") | Out-Null } catch { }
        Start-Sleep -Seconds 1
    }

    return $selectedId
}

# ============================ MAIN (BATCH) ============================
try {
    Write-Step "=== Meta AI factory FAST batch (Mail.cx) ==="
    $ActiveSerial = Select-Device
    Write-Ok "device: $ActiveSerial"

    if ($UserId -lt 0) {
        $UserId = Select-AndroidUser
    } else {
        Write-Ok "Using configured User ID: $UserId"
    }

    $pkgCheck = (Adb @("shell", "pm", "list", "packages", "--user", "$UserId")) -join "`n"
    if ($pkgCheck -notmatch [regex]::Escape($Pkg)) {
        Write-Err "Warning: $Pkg is NOT installed for User $UserId!"
    }

    $okCount = 0
    $i = 1
    if ($ProgressFile) { "0|$Count|0" | Out-File $ProgressFile -Encoding UTF8 }
    while ($i -le $Count) {
        Write-Step "--- Account $i / $Count ---"
        $attempt = 0
        $done = $false
        while (-not $done) {
            $attempt++
            if (Create-One-Account $i) {
                $okCount++
                $done = $true
                Clear-AppData
                Write-Step "app pre-warmed for next account"
            } else {
                $backoff = [Math]::Min(60, 10 * $attempt)
                Write-Err "account $i failed (attempt $attempt) - not counted, retrying in ${backoff}s"
                Start-Sleep -Seconds $backoff
            }
        }
        $i++
        if ($ProgressFile) { "$okCount|$Count|$i" | Out-File $ProgressFile -Encoding UTF8 }
    }

    Write-Step "=== BATCH COMPLETE: $okCount / $Count succeeded ==="
    $batchTotal = [int]((Get-Date) - $script:StartOfBatch).TotalSeconds
    Write-Ok "batch total: ${batchTotal}s"
    Write-Ok "credentials in $CredFile"
    if ($ProgressFile) { "$okCount|$Count|$($Count + 1)" | Out-File $ProgressFile -Encoding UTF8 }
    if ($KeepLog) {
        $script:LogLines | Out-File "$TempDir\metaai_mailcx_log.txt" -Append -Encoding UTF8
    }
    exit 0
} catch {
    Write-Err "FATAL: $_"
    if ($ProgressFile) { "$okCount|$Count|$($Count + 1)" | Out-File $ProgressFile -Encoding UTF8 }
    exit 1
}
