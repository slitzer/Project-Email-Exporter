<#
    .SYNOPSIS
    Export emails from specific Outlook folders (matched by project number) to local .eml files,
    grouped by conversation thread. Launches a Windows Forms GUI if no parameters are supplied.

    .DESCRIPTION
    Reads a CSV file containing project information, finds the matching mail folder in Outlook
    by project number, then downloads every email grouped into subfolders by conversation thread.

    Use -WhatIf to do a dry run: shows exactly what would be downloaded without saving any files.

    Folder structure:
        <OutputFolder>\
            <ClientName>\
                <ProjectNumber> - <ProjectTitle>\
                    <Date> - <ConversationSubject>\
                        2024-12-03_2148 - Plate Storage Racks.eml
                        2024-12-10_0255 - RE_ Plate Storage Racks.eml

    .PARAMETER MailboxUserId
    The UPN or Object ID (GUID) of the mailbox to export from.

    .PARAMETER CsvPath
    Path to the CSV file. Expected columns: Category, Project Number, Project Title, Client Name

    .PARAMETER OutputFolder
    Path to the local folder where emails will be saved.

    .PARAMETER FolderMatchDepth
    How many levels deep to search for the matching Outlook folder. Default: 3

    .PARAMETER WhatIf
    Dry run mode. Lists all emails and thread folders that WOULD be created, without downloading anything.

    .EXAMPLE
    # Launch the GUI (no parameters needed):
    .\Export-ProjectEmails.ps1

    # Or run headlessly:
    Connect-MgGraph -Scopes "Mail.ReadWrite","Mail.ReadWrite.Shared","User.Read"
    .\Export-ProjectEmails.ps1 `
        -MailboxUserId "user@company.com" `
        -CsvPath "C:\TRANSFERSCRIPT\ProjectsToArchive.csv" `
        -OutputFolder "C:\TRANSFERSCRIPT\M365Export" `
        -WhatIf

    .NOTES
    Requires:
        - PowerShell 7.3.4 or later
        - Microsoft.Graph.Authentication module v2.0.0+
        - Microsoft.Graph.Mail module v2.0.0+
        - Windows (for GUI mode — headless mode works on any OS)
#>
#Requires -Version 7.3.4
#Requires -Module @{ ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '2.0.0' }
#Requires -Module @{ ModuleName = 'Microsoft.Graph.Mail'; ModuleVersion = '2.0.0' }

[CmdletBinding(SupportsShouldProcess)]
param (
    [ValidateScript({
        if (-not $_) { return $true }
        if ($_ -match "^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$") { $true }
        elseif ([guid]::TryParse($_, [ref]$null)) { $true }
        else { throw 'Supply a valid UPN (email address) or Azure AD Object ID (GUID).' }
    })]
    [string]$MailboxUserId,

    [ValidateScript({ -not $_ -or (Test-Path $_ -PathType Leaf) })]
    [string]$CsvPath,

    [string]$OutputFolder,

    [switch]$MoveToPredelete,

    [int]$FolderMatchDepth = 3
)

# ══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════════════════

function Get-SafeName {
    param([string]$Name, [int]$MaxLength = 60)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    $safe = ($Name -replace "[$([regex]::Escape($invalid))]", '_').Trim()
    if ($safe.Length -gt $MaxLength) { $safe = $safe.Substring(0, $MaxLength).TrimEnd('_', ' ') }
    return $safe
}


function New-RunLogContext {
    param([string]$OutputFolder,[string]$Mode,[string]$MailboxUserId)
    if (-not $OutputFolder) { $OutputFolder = (Join-Path $PWD 'M365Export') }
    $logRoot = Join-Path $OutputFolder 'Logs'
    New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
    $runStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $runId    = "$runStamp-$Mode"
    $runDir   = Join-Path $logRoot $runId
    New-Item -ItemType Directory -Path $runDir -Force | Out-Null
    $summaryPath = Join-Path $runDir "RunSummary_$runId.csv"
    $eventPath   = Join-Path $runDir "RunEvents_$runId.log"
    [PSCustomObject]@{Timestamp=(Get-Date).ToString('s');RunId=$runId;Mode=$Mode;Mailbox=$MailboxUserId;OutputFolder=$OutputFolder;Status='Started';Details='Run initialised'} | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
    "[$((Get-Date).ToString('s'))] Started $Mode run for mailbox $MailboxUserId" | Out-File -FilePath $eventPath -Encoding UTF8
    return [PSCustomObject]@{RunId=$runId;LogRoot=$logRoot;RunDir=$runDir;SummaryPath=$summaryPath;EventPath=$eventPath}
}

function Get-ProjectLogPath {
    param([object]$LogContext,[string]$ClientName,[string]$ProjectNumber,[string]$ProjectTitle)
    $safeClient = Get-SafeName ($ClientName ?? 'Unknown Client') -MaxLength 60
    $safeProject = Get-SafeName ("$ProjectNumber - $ProjectTitle") -MaxLength 80
    $clientDir = Join-Path $LogContext.RunDir $safeClient
    New-Item -ItemType Directory -Path $clientDir -Force | Out-Null
    return Join-Path $clientDir "$safeProject.csv"
}

function Write-ProjectLog {
    param(
        [object]$LogContext,[string]$MailboxUserId,[string]$Mode,[string]$ClientName,[string]$ProjectNumber,[string]$ProjectTitle,
        [string]$MatchedFolder='',[string]$Action='',[string]$Status='',[int]$MessageCount=0,[int]$ChildFolderCount=0,
        [string]$SourceFolderId='',[string]$DestinationPath='',[string]$Error=''
    )
    if (-not $LogContext) { return }
    $path = Get-ProjectLogPath -LogContext $LogContext -ClientName $ClientName -ProjectNumber $ProjectNumber -ProjectTitle $ProjectTitle
    [PSCustomObject]@{
        Timestamp=(Get-Date).ToString('s');RunId=$LogContext.RunId;Mode=$Mode;Mailbox=$MailboxUserId;ClientName=$ClientName;
        ProjectNumber=$ProjectNumber;ProjectTitle=$ProjectTitle;MatchedFolder=$MatchedFolder;Action=$Action;Status=$Status;
        MessageCount=$MessageCount;ChildFolderCount=$ChildFolderCount;SourceFolderId=$SourceFolderId;DestinationPath=$DestinationPath;Error=$Error
    } | Export-Csv -Path $path -NoTypeInformation -Append -Encoding UTF8
}

function Write-RunLog {
    param([object]$LogContext,[string]$MailboxUserId,[string]$Mode,[string]$Status,[string]$Details)
    if (-not $LogContext) { return }
    [PSCustomObject]@{Timestamp=(Get-Date).ToString('s');RunId=$LogContext.RunId;Mode=$Mode;Mailbox=$MailboxUserId;OutputFolder='';Status=$Status;Details=$Details} | Export-Csv -Path $LogContext.SummaryPath -NoTypeInformation -Append -Encoding UTF8
    "[$((Get-Date).ToString('s'))] [$Status] $Details" | Out-File -FilePath $LogContext.EventPath -Append -Encoding UTF8
}


function Ensure-ArchiveStatusColumns {
    param([object[]]$ProjectList)
    foreach ($row in $ProjectList) {
        if (-not ($row.PSObject.Properties.Name -contains 'Status/Error Info')) { $row | Add-Member -NotePropertyName 'Status/Error Info' -NotePropertyValue '' -Force }
        if (-not ($row.PSObject.Properties.Name -contains 'Matched Folder')) { $row | Add-Member -NotePropertyName 'Matched Folder' -NotePropertyValue '' -Force }
        if (-not ($row.PSObject.Properties.Name -contains 'Processed At')) { $row | Add-Member -NotePropertyName 'Processed At' -NotePropertyValue '' -Force }
    }
}

function New-ArchiveStatusCsvPath {
    param([string]$OutputFolder,[string]$CsvSourcePath,[string]$RunId)
    if (-not $OutputFolder) { $OutputFolder = (Join-Path $PWD 'M365Export') }
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    $baseName = if ($CsvSourcePath) { [IO.Path]::GetFileNameWithoutExtension($CsvSourcePath) } else { 'ProjectsToArchive' }
    return (Join-Path $OutputFolder ("{0}_UPDATED_{1}.csv" -f $baseName,$RunId))
}

function Save-ArchiveStatusCsv {
    param([object[]]$ProjectList,[string]$Path)
    if (-not $Path) { return }
    $ProjectList | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}

function Set-ArchiveProjectStatus {
    param([object]$Project,[string]$Category,[string]$StatusInfo,[string]$MatchedFolder = '',[object[]]$ProjectList,[string]$StatusCsvPath)
    $Project.Category = $Category
    if (-not ($Project.PSObject.Properties.Name -contains 'Status/Error Info')) { $Project | Add-Member -NotePropertyName 'Status/Error Info' -NotePropertyValue $StatusInfo -Force } else { $Project.'Status/Error Info' = $StatusInfo }
    if (-not ($Project.PSObject.Properties.Name -contains 'Matched Folder')) { $Project | Add-Member -NotePropertyName 'Matched Folder' -NotePropertyValue $MatchedFolder -Force } else { $Project.'Matched Folder' = $MatchedFolder }
    if (-not ($Project.PSObject.Properties.Name -contains 'Processed At')) { $Project | Add-Member -NotePropertyName 'Processed At' -NotePropertyValue (Get-Date).ToString('s') -Force } else { $Project.'Processed At' = (Get-Date).ToString('s') }
    Save-ArchiveStatusCsv -ProjectList $ProjectList -Path $StatusCsvPath
}

function Find-ProjectFolderCandidates {
    param([object[]]$AllFolders, [string]$ProjectNumber)
    $prefixMatches = @($AllFolders | Where-Object { $_.DisplayName -match "^$([regex]::Escape($ProjectNumber))\b" })
    if ($prefixMatches.Count -gt 0) { return $prefixMatches }
    return @($AllFolders | Where-Object { $_.DisplayName -like "*$ProjectNumber*" })
}

function Get-MgAccessToken {
    param([string]$UserId)
    try {
        $resp = Invoke-MgGraphRequest `
            -Uri        "https://graph.microsoft.com/v1.0/users/$UserId/mailFolders?`$top=1" `
            -Method     GET `
            -OutputType HttpResponseMessage `
            -ErrorAction Stop
        $token = $resp.RequestMessage.Headers.Authorization.Parameter
        $resp.Dispose()
        if ($token) { return $token }
    } catch {}
    try {
        $token = [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.AuthContext.AccessToken
        if ($token) { return $token }
    } catch {}
    throw "Could not retrieve access token. Please reconnect: Connect-MgGraph -Scopes 'Mail.ReadWrite','Mail.ReadWrite.Shared','User.Read'"
}

function Get-AllMailFolders {
    param(
        [string]$UserId,
        [string]$ParentFolderId = $null,
        [int]$CurrentDepth = 0,
        [int]$MaxDepth = 3
    )
    $results = [System.Collections.Generic.List[object]]::new()
    if ($CurrentDepth -gt $MaxDepth) { return $results }
    try {
        if ($ParentFolderId) {
            $folders = Get-MgUserMailFolderChildFolder `
                -UserId $UserId -MailFolderId $ParentFolderId -All -ErrorAction Stop
        } else {
            $folders = Get-MgUserMailFolder -UserId $UserId -All -ErrorAction Stop
        }
    } catch {
        Write-Warning "  Could not list mail folders (depth $CurrentDepth): $_"
        return $results
    }
    foreach ($folder in $folders) {
        $results.Add([PSCustomObject]@{
            Id          = $folder.Id
            DisplayName = $folder.DisplayName
            ChildCount  = $folder.ChildFolderCount
        })
        if ($folder.ChildFolderCount -gt 0 -and $CurrentDepth -lt $MaxDepth) {
            $children = Get-AllMailFolders `
                -UserId $UserId -ParentFolderId $folder.Id `
                -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth
            foreach ($child in $children) { $results.Add($child) }
        }
    }
    return $results
}

function Find-ProjectFolder {
    param([object[]]$AllFolders, [string]$ProjectNumber)
    $match = $AllFolders |
             Where-Object { $_.DisplayName -match "^$([regex]::Escape($ProjectNumber))\b" } |
             Select-Object -First 1
    if (-not $match) {
        $match = $AllFolders |
                 Where-Object { $_.DisplayName -like "*$ProjectNumber*" } |
                 Select-Object -First 1
    }
    return $match
}

function Get-MessageDiagnostic {
    param([string]$UserId, [string]$MessageId, [string]$AccessToken)
    try {
        $meta = Invoke-MgGraphRequest `
            -Uri    "https://graph.microsoft.com/v1.0/users/$UserId/messages/$MessageId`?`$select=id,subject,hasAttachments,internetMessageId" `
            -Method GET `
            -ErrorAction Stop
        $uri = "https://graph.microsoft.com/v1.0/users/$UserId/messages/$MessageId/`$value"
        $resp = Invoke-WebRequest `
            -Uri     $uri `
            -Method  HEAD `
            -Headers @{ Authorization = "Bearer $AccessToken" } `
            -ErrorAction SilentlyContinue
        if ($resp -and $resp.StatusCode -eq 200) { return "HEAD request OK — retry may succeed" }
        return "HEAD request returned: $($resp.StatusCode ?? 'no response')"
    } catch {
        $errMsg = $_.ToString()
        if ($errMsg -match '403')  { return "Access denied (403) — insufficient permissions or IRM-protected message" }
        if ($errMsg -match '404')  { return "Message not found (404) — may have been deleted or moved since folder scan" }
        if ($errMsg -match '423')  { return "Message locked (423) — possibly being processed by Exchange" }
        return "Error accessing message metadata: $errMsg"
    }
}

function Save-MessageAsEml {
    param(
        [string]$UserId,
        [string]$MessageId,
        [string]$DestinationPath,
        [string]$AccessToken,
        [bool]$IsDryRun = $false
    )
    if ($IsDryRun) {
        Write-Host "      [WHATIF] Would download to: $(Split-Path $DestinationPath -Leaf)" -ForegroundColor Cyan
        return 'whatif'
    }
    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/messages/$MessageId/`$value"
    try {
        Invoke-WebRequest `
            -Uri     $uri `
            -Method  GET `
            -Headers @{ Authorization = "Bearer $AccessToken" } `
            -OutFile $DestinationPath `
            -ErrorAction Stop | Out-Null
        if ((Get-Item $DestinationPath -ErrorAction SilentlyContinue).Length -gt 0) {
            return 'ok'
        } else {
            Remove-Item $DestinationPath -Force -ErrorAction SilentlyContinue
            $reason = Get-MessageDiagnostic -UserId $UserId -MessageId $MessageId -AccessToken $AccessToken
            Write-Warning "  Empty file [$($MessageId.Substring([Math]::Max(0,$MessageId.Length-12)))]: $reason"
            return 'failed'
        }
    } catch {
        Remove-Item $DestinationPath -Force -ErrorAction SilentlyContinue
        $httpCode = if ($_ -match '(\d{3})') { $Matches[1] } else { 'unknown' }
        Write-Warning "  HTTP $httpCode [$($MessageId.Substring([Math]::Max(0,$MessageId.Length-12)))]: $_"
        return 'failed'
    }
}

function Export-FolderMessages {
    param(
        [string]$UserId,
        [string]$FolderId,
        [string]$DestinationFolder,
        [string]$AccessToken,
        [bool]$IsDryRun,
        [string]$CsvSourcePath
    )

    $getParams = @{
        UserId       = $UserId
        MailFolderId = $FolderId
        All          = $true
        Select       = 'id,subject,receivedDateTime,conversationId'
    }

    try {
        $messages = Get-MgUserMailFolderMessage @getParams -ErrorAction Stop
    } catch {
        Write-Warning "  Could not retrieve messages: $_"
        return @{ Saved = 0; Failed = 0; Threads = 0; WhatIfCount = 0 }
    }

    if (-not $messages -or $messages.Count -eq 0) {
        Write-Host "  No emails found in this folder."
        return @{ Saved = 0; Failed = 0; Threads = 0; WhatIfCount = 0 }
    }

    $conversations = $messages | Group-Object -Property ConversationId

    if ($IsDryRun) {
        Write-Host "  [WHATIF] Would process $($messages.Count) email(s) across $($conversations.Count) thread(s):" -ForegroundColor Cyan
    } else {
        Write-Host "  Found $($messages.Count) email(s) across $($conversations.Count) conversation thread(s)."
    }

    $savedCount  = 0
    $failedCount = 0
    $whatIfCount = 0
    $threadNum   = 0

    foreach ($conv in $conversations) {
        $threadNum++
        $threadMessages = $conv.Group | Sort-Object ReceivedDateTime
        $firstMsg       = $threadMessages | Select-Object -First 1

        $firstDate   = if ($firstMsg.ReceivedDateTime) {
            $firstMsg.ReceivedDateTime.ToString('yyyy-MM-dd')
        } else { 'unknown-date' }

        $baseSubject = ($firstMsg.Subject ?? 'no-subject') -replace '^(RE_\s*|FW_\s*|Re:\s*|Fw:\s*)+', ''
        $safeSubject = Get-SafeName $baseSubject

        $threadFolderName = "${firstDate} - ${safeSubject}"
        $threadFolder     = Join-Path $DestinationFolder $threadFolderName
        if ((Test-Path $threadFolder) -and -not $IsDryRun) {
            $threadFolder = "${threadFolder}_t${threadNum}"
        }

        if ($IsDryRun) {
            Write-Host ("  [{0,3}/{1}] Thread: '{2}' — {3} email(s)" -f `
                $threadNum, $conversations.Count, $safeSubject, $threadMessages.Count) -ForegroundColor Cyan
            Write-Host "            Folder: $threadFolder" -ForegroundColor DarkCyan
            foreach ($msg in $threadMessages) {
                $d = if ($msg.ReceivedDateTime) { $msg.ReceivedDateTime.ToString('yyyy-MM-dd HH:mm') } else { '?' }
                $s = $msg.Subject ?? 'no-subject'
                Write-Host ("            • {0}  {1}" -f $d, $s) -ForegroundColor DarkCyan
                $whatIfCount++
            }
            continue
        }

        New-Item -ItemType Directory -Path $threadFolder -Force | Out-Null
        Write-Host "  Thread $threadNum/$($conversations.Count): '$safeSubject' ($($threadMessages.Count) email(s))"

        foreach ($msg in $threadMessages) {
            $datePrefix = if ($msg.ReceivedDateTime) {
                $msg.ReceivedDateTime.ToString('yyyy-MM-dd_HHmm')
            } else { 'unknown-date' }

            $safeMsg  = Get-SafeName ($msg.Subject ?? 'no-subject')
            $shortId  = $msg.Id.Substring([Math]::Max(0, $msg.Id.Length - 8))
            $fileName = "${datePrefix}_${safeMsg}_${shortId}.eml"
            $filePath = Join-Path $threadFolder $fileName

            if ((Test-Path $filePath) -and (Get-Item $filePath).Length -gt 0) {
                Write-Verbose "    Skipping (exists): $fileName"
                $savedCount++
                continue
            }

            $result = Save-MessageAsEml `
                -UserId          $UserId `
                -MessageId       $msg.Id `
                -DestinationPath $filePath `
                -AccessToken     $AccessToken `
                -IsDryRun        $false

            switch ($result) {
                'ok'     { $savedCount++ }
                'failed' { $failedCount++ }
            }
        }
    }
    return @{ Saved = $savedCount; Failed = $failedCount; Threads = $threadNum; WhatIfCount = $whatIfCount }
}

# ══════════════════════════════════════════════════════════════════════════════
#  MOVE TO PREDELETE
# ══════════════════════════════════════════════════════════════════════════════

# Ensure Inbox\PREDELETE exists, creating it if necessary. Returns the folder ID.
function Get-OrCreate-PredeleteFolder {
    param([string]$UserId)

    # Find Inbox
    $inbox = Get-MgUserMailFolder -UserId $UserId -All |
             Where-Object { $_.DisplayName -eq 'Inbox' } |
             Select-Object -First 1
    if (-not $inbox) { throw "Could not find Inbox folder for '$UserId'." }

    # Look for existing PREDELETE child
    $predelete = Get-MgUserMailFolderChildFolder -UserId $UserId -MailFolderId $inbox.Id -All |
                 Where-Object { $_.DisplayName -eq 'PREDELETE' } |
                 Select-Object -First 1

    if ($predelete) { return $predelete.Id }

    # Create it
    $body = @{ displayName = 'PREDELETE' } | ConvertTo-Json
    $created = Invoke-MgGraphRequest `
        -Uri    "https://graph.microsoft.com/v1.0/users/$UserId/mailFolders/$($inbox.Id)/childFolders" `
        -Method POST `
        -Body   $body `
        -ContentType 'application/json' `
        -ErrorAction Stop
    Write-Host "  Created folder: Inbox\PREDELETE"
    return $created.id
}

function Invoke-MoveToPredelete {
    param(
        [string]$MailboxUserId,
        [string]$OutputFolder,
        [object[]]$ProjectList,
        [int]$FolderMatchDepth,
        [bool]$IsDryRun,
        [string]$CsvSourcePath
    )

    Write-Host ""
    if ($IsDryRun) {
        Write-Host "╔══════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "║       DRY RUN — MOVE FOLDERS TO PREDELETE   ║" -ForegroundColor Yellow
        Write-Host "║  No folders or messages will be moved.       ║" -ForegroundColor Yellow
        Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host ""
    }

    $logContext = New-RunLogContext -OutputFolder $OutputFolder -Mode 'MoveToPredelete' -MailboxUserId $MailboxUserId
    Write-Host "Logging to: $($logContext.RunDir)" -ForegroundColor DarkGray
    Ensure-ArchiveStatusColumns -ProjectList $ProjectList
    foreach ($projectRow in $ProjectList) {
        $projectRow.Category = 'Review Required'
        $projectRow.'Status/Error Info' = 'Pending: run started but this row has not completed yet.'
        $projectRow.'Processed At' = ''
    }
    $archiveStatusCsvPath = New-ArchiveStatusCsvPath -OutputFolder $OutputFolder -CsvSourcePath $CsvSourcePath -RunId $logContext.RunId
    Save-ArchiveStatusCsv -ProjectList $ProjectList -Path $archiveStatusCsvPath
    Write-Host "Status CSV: $archiveStatusCsvPath" -ForegroundColor DarkGray

    $mgContext = Get-MgContext
    if (-not $mgContext) { throw "Not connected to Microsoft Graph." }
    Write-Host "Connected as: $($mgContext.Account)"

    # Moving a folder in another mailbox via Graph requires Mail.ReadWrite.Shared
    # in addition to the Exchange mailbox delegation/full access.
    if ($mgContext.Account -ne $MailboxUserId) {
        Write-Warning "  Signed-in account ($($mgContext.Account)) differs from target mailbox ($MailboxUserId)."
        Write-Warning "  Folder move operations require Mail.ReadWrite.Shared in the Graph token."
        Write-Warning "  If moves fail with 403, reconnect with: Connect-MgGraph -Scopes 'Mail.ReadWrite','Mail.ReadWrite.Shared','User.Read'"
    }

    Write-Host "`nScanning mailbox folders (depth: $FolderMatchDepth)..."
    $allFolders = Get-AllMailFolders -UserId $MailboxUserId -MaxDepth $FolderMatchDepth
    Write-Host "Found $($allFolders.Count) total mail folder(s)."

    Write-Host "Resolving Inbox\PREDELETE..."
    $predeleteId = Get-OrCreate-PredeleteFolder -UserId $MailboxUserId
    Write-Host "  Target folder ID: $($predeleteId.Substring(0, [Math]::Min(20,$predeleteId.Length)))..."

    # Get current PREDELETE child folders so we can warn about duplicates before Graph throws a vague error.
    $existingPredeleteChildren = @{}
    try {
        $children = Get-MgUserMailFolderChildFolder -UserId $MailboxUserId -MailFolderId $predeleteId -All -ErrorAction Stop
        foreach ($child in $children) { $existingPredeleteChildren[$child.DisplayName.ToLowerInvariant()] = $true }
    } catch {
        Write-Warning "  Could not check existing PREDELETE child folders: $_"
    }

    $totalFoldersMoved  = 0
    $totalFoldersFailed = 0
    $totalFoldersSkipped = 0
    $totalMessagesInside = 0

    foreach ($project in $ProjectList) {
        $projectNumber = $project.'Project Number'.Trim()
        $projectTitle  = $project.'Project Title'.Trim()
        $clientName    = $project.'Client Name'.Trim()

        Write-Host "`n──────────────────────────────────────────────"
        Write-Host "Project : $projectNumber - $projectTitle"
        Write-Host "Client  : $clientName"

        $folderCandidates = @(Find-ProjectFolderCandidates -AllFolders $allFolders -ProjectNumber $projectNumber)
        if ($folderCandidates.Count -eq 0) {
            $reason = "Source folder missing: no Outlook folder found matching project number '$projectNumber'."
            Write-Warning "  $reason Skipping."
            Write-ProjectLog -LogContext $logContext -MailboxUserId $MailboxUserId -Mode 'MoveToPredelete' -ClientName $clientName -ProjectNumber $projectNumber -ProjectTitle $projectTitle -Action 'FindFolder' -Status 'Skipped-NoFolder' -DestinationPath 'Inbox\PREDELETE' -Error $reason
            Set-ArchiveProjectStatus -Project $project -Category 'Review Required' -StatusInfo $reason -ProjectList $ProjectList -StatusCsvPath $archiveStatusCsvPath
            $totalFoldersSkipped++
            continue
        }
        if ($folderCandidates.Count -gt 1) {
            $candidateNames = ($folderCandidates | Select-Object -ExpandProperty DisplayName) -join ' | '
            $reason = "Ambiguous folder match: found $($folderCandidates.Count) folders for project number '$projectNumber': $candidateNames"
            Write-Warning "  $reason"
            Write-Warning "  Skipping so the wrong folder does not get moved."
            Write-ProjectLog -LogContext $logContext -MailboxUserId $MailboxUserId -Mode 'MoveToPredelete' -ClientName $clientName -ProjectNumber $projectNumber -ProjectTitle $projectTitle -Action 'FindFolder' -Status 'Skipped-AmbiguousFolder' -DestinationPath 'Inbox\PREDELETE' -Error $reason
            Set-ArchiveProjectStatus -Project $project -Category 'Review Required' -StatusInfo $reason -ProjectList $ProjectList -StatusCsvPath $archiveStatusCsvPath
            $totalFoldersSkipped++
            continue
        }

        $matchedFolder = $folderCandidates[0]
        Write-Host "  Matched folder: '$($matchedFolder.DisplayName)'"
        $sourceFolderId = $matchedFolder.Id

        if ($sourceFolderId -eq $predeleteId -or $matchedFolder.DisplayName -eq 'PREDELETE') {
            Write-Warning "  Refusing to move PREDELETE into itself. Skipping."
            $reason = 'Safety check: matched folder is PREDELETE or same as destination.'
            Write-ProjectLog -LogContext $logContext -MailboxUserId $MailboxUserId -Mode 'MoveToPredelete' -ClientName $clientName -ProjectNumber $projectNumber -ProjectTitle $projectTitle -MatchedFolder $matchedFolder.DisplayName -Action 'MoveFolder' -Status 'Skipped-SelfMove' -SourceFolderId $sourceFolderId -DestinationPath 'Inbox\PREDELETE' -Error $reason
            Set-ArchiveProjectStatus -Project $project -Category 'Review Required' -StatusInfo $reason -MatchedFolder $matchedFolder.DisplayName -ProjectList $ProjectList -StatusCsvPath $archiveStatusCsvPath
            $totalFoldersSkipped++
            continue
        }

        if ($existingPredeleteChildren.ContainsKey($matchedFolder.DisplayName.ToLowerInvariant())) {
            Write-Warning "  PREDELETE already contains a folder named '$($matchedFolder.DisplayName)'. Skipping to avoid duplicate/conflict."
            Write-Warning "  Rename or remove the existing PREDELETE child folder, then re-run this project."
            $reason = "Destination already contains a folder named '$($matchedFolder.DisplayName)'. Rename/remove that PREDELETE child folder and re-run."
            Write-ProjectLog -LogContext $logContext -MailboxUserId $MailboxUserId -Mode 'MoveToPredelete' -ClientName $clientName -ProjectNumber $projectNumber -ProjectTitle $projectTitle -MatchedFolder $matchedFolder.DisplayName -Action 'MoveFolder' -Status 'Skipped-DuplicateDestination' -ChildFolderCount $matchedFolder.ChildCount -SourceFolderId $sourceFolderId -DestinationPath "Inbox\PREDELETE\$($matchedFolder.DisplayName)" -Error $reason
            Set-ArchiveProjectStatus -Project $project -Category 'Review Required' -StatusInfo $reason -MatchedFolder $matchedFolder.DisplayName -ProjectList $ProjectList -StatusCsvPath $archiveStatusCsvPath
            $totalFoldersSkipped++
            continue
        }

        # Count direct messages for reporting only. Folder move preserves direct messages and child folders.
        $messageCount = 0
        try {
            $messages = Get-MgUserMailFolderMessage `
                -UserId $MailboxUserId -MailFolderId $sourceFolderId `
                -All -Select 'id' -ErrorAction Stop
            $messageCount = @($messages).Count
        } catch {
            Write-Warning "  Could not count messages in folder before move: $_"
        }
        $totalMessagesInside += $messageCount

        Write-Host "  Folder contains $messageCount direct message(s) and $($matchedFolder.ChildCount) child folder(s)."

        if ($IsDryRun) {
            Write-Host "  [WHATIF] Would move folder '$($matchedFolder.DisplayName)' into Inbox\PREDELETE." -ForegroundColor Cyan
            $reason = "WHATIF only: would move folder '$($matchedFolder.DisplayName)' to Inbox\PREDELETE."
            Write-ProjectLog -LogContext $logContext -MailboxUserId $MailboxUserId -Mode 'MoveToPredelete' -ClientName $clientName -ProjectNumber $projectNumber -ProjectTitle $projectTitle -MatchedFolder $matchedFolder.DisplayName -Action 'MoveFolder' -Status 'WhatIf' -MessageCount $messageCount -ChildFolderCount $matchedFolder.ChildCount -SourceFolderId $sourceFolderId -DestinationPath "Inbox\PREDELETE\$($matchedFolder.DisplayName)"
            Set-ArchiveProjectStatus -Project $project -Category 'Review Required' -StatusInfo $reason -MatchedFolder $matchedFolder.DisplayName -ProjectList $ProjectList -StatusCsvPath $archiveStatusCsvPath
            $totalFoldersMoved++
            continue
        }

        try {
            $body = @{ destinationId = $predeleteId } | ConvertTo-Json

            # Move the folder itself. This preserves the folder name, contained emails, and child folders.
            Invoke-MgGraphRequest `
                -Uri    "https://graph.microsoft.com/v1.0/users/$MailboxUserId/mailFolders/$sourceFolderId/move" `
                -Method POST `
                -Body   $body `
                -ContentType 'application/json' `
                -ErrorAction Stop | Out-Null

            Write-Host "  Moved folder to: Inbox\PREDELETE\$($matchedFolder.DisplayName)" -ForegroundColor Green
            $reason = "Moved folder '$($matchedFolder.DisplayName)' to Inbox\PREDELETE. Direct messages: $messageCount; child folders: $($matchedFolder.ChildCount)."
            Write-ProjectLog -LogContext $logContext -MailboxUserId $MailboxUserId -Mode 'MoveToPredelete' -ClientName $clientName -ProjectNumber $projectNumber -ProjectTitle $projectTitle -MatchedFolder $matchedFolder.DisplayName -Action 'MoveFolder' -Status 'Moved' -MessageCount $messageCount -ChildFolderCount $matchedFolder.ChildCount -SourceFolderId $sourceFolderId -DestinationPath "Inbox\PREDELETE\$($matchedFolder.DisplayName)"
            Set-ArchiveProjectStatus -Project $project -Category 'Archiving Done' -StatusInfo $reason -MatchedFolder $matchedFolder.DisplayName -ProjectList $ProjectList -StatusCsvPath $archiveStatusCsvPath
            $totalFoldersMoved++
            $existingPredeleteChildren[$matchedFolder.DisplayName.ToLowerInvariant()] = $true
        } catch {
            Write-Warning ("  Failed to move folder [{0}]: {1}" -f $matchedFolder.DisplayName, $_)
            $reason = "Move failed: $($_.Exception.Message)"
            Write-ProjectLog -LogContext $logContext -MailboxUserId $MailboxUserId -Mode 'MoveToPredelete' -ClientName $clientName -ProjectNumber $projectNumber -ProjectTitle $projectTitle -MatchedFolder $matchedFolder.DisplayName -Action 'MoveFolder' -Status 'Failed' -MessageCount $messageCount -ChildFolderCount $matchedFolder.ChildCount -SourceFolderId $sourceFolderId -DestinationPath "Inbox\PREDELETE\$($matchedFolder.DisplayName)" -Error $_.ToString()
            Set-ArchiveProjectStatus -Project $project -Category 'Review Required' -StatusInfo $reason -MatchedFolder $matchedFolder.DisplayName -ProjectList $ProjectList -StatusCsvPath $archiveStatusCsvPath
            $totalFoldersFailed++
        }
    }

    Write-Host ""
    Write-Host "══════════════════════════════════════════════"
    if ($IsDryRun) {
        Write-Host "DRY RUN complete. No folders were moved." -ForegroundColor Yellow
        Write-Host "  Folders that would be moved : $totalFoldersMoved"
        Write-Host "  Folders skipped             : $totalFoldersSkipped"
        Write-Host "  Direct messages inside      : $totalMessagesInside"
    } else {
        Write-Host "Folder move complete."
        Write-Host "  Folders moved    : $totalFoldersMoved"
        Write-Host "  Folders failed   : $totalFoldersFailed"
        Write-Host "  Folders skipped  : $totalFoldersSkipped"
        Write-Host "  Destination      : Inbox\PREDELETE"
    }
    Write-Host "══════════════════════════════════════════════"
    Save-ArchiveStatusCsv -ProjectList $ProjectList -Path $archiveStatusCsvPath
    Write-Host "  Status CSV       : $archiveStatusCsvPath"
    Write-RunLog -LogContext $logContext -MailboxUserId $MailboxUserId -Mode 'MoveToPredelete' -Status 'Completed' -Details "Moved=$totalFoldersMoved; Failed=$totalFoldersFailed; Skipped=$totalFoldersSkipped; DirectMessages=$totalMessagesInside; Destination=Inbox\PREDELETE; StatusCsv=$archiveStatusCsvPath"
}

# ══════════════════════════════════════════════════════════════════════════════
#  WINDOWS FORMS GUI
# ══════════════════════════════════════════════════════════════════════════════

function Show-ExportGui {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # ── Form ──────────────────────────────────────────────────────────────────
    $form                  = [System.Windows.Forms.Form]::new()
    $form.Text             = 'Project Email Exporter'
    $form.Size             = [System.Drawing.Size]::new(680, 660)
    $form.MinimumSize      = [System.Drawing.Size]::new(640, 660)
    $form.StartPosition    = 'CenterScreen'
    $form.Font             = [System.Drawing.Font]::new('Segoe UI', 9)
    $form.BackColor        = [System.Drawing.Color]::FromArgb(245, 245, 245)
    $form.FormBorderStyle  = 'FixedDialog'
    $form.MaximizeBox      = $false

    # ── Helpers ───────────────────────────────────────────────────────────────
    function New-Label {
        param([string]$Text, [int]$X, [int]$Y, [int]$W = 160, [int]$H = 18)
        $l           = [System.Windows.Forms.Label]::new()
        $l.Text      = $Text
        $l.Location  = [System.Drawing.Point]::new($X, $Y)
        $l.Size      = [System.Drawing.Size]::new($W, $H)
        $l.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
        return $l
    }

    function New-TextBox {
        param([int]$X, [int]$Y, [int]$W = 380, [string]$PlaceHolder = '')
        $t               = [System.Windows.Forms.TextBox]::new()
        $t.Location      = [System.Drawing.Point]::new($X, $Y)
        $t.Size          = [System.Drawing.Size]::new($W, 24)
        $t.BackColor     = [System.Drawing.Color]::White
        $t.BorderStyle   = 'FixedSingle'
        if ($PlaceHolder) {
            $t.ForeColor = [System.Drawing.Color]::Gray
            $t.Text      = $PlaceHolder
            $t.Add_Enter({ if ($this.ForeColor -eq [System.Drawing.Color]::Gray) { $this.Text = ''; $this.ForeColor = [System.Drawing.Color]::Black } })
            $t.Add_Leave({ if ($this.Text -eq '') { $this.ForeColor = [System.Drawing.Color]::Gray; $this.Text = $PlaceHolder } })
        }
        return $t
    }

    function New-Button {
        param([string]$Text, [int]$X, [int]$Y, [int]$W = 110, [int]$H = 28)
        $b             = [System.Windows.Forms.Button]::new()
        $b.Text        = $Text
        $b.Location    = [System.Drawing.Point]::new($X, $Y)
        $b.Size        = [System.Drawing.Size]::new($W, $H)
        $b.FlatStyle   = 'Flat'
        $b.BackColor   = [System.Drawing.Color]::White
        $b.ForeColor   = [System.Drawing.Color]::FromArgb(40, 40, 40)
        $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
        $b.Cursor      = 'Hand'
        return $b
    }

    function New-GroupBox {
        param([string]$Text, [int]$X, [int]$Y, [int]$W, [int]$H)
        $g             = [System.Windows.Forms.GroupBox]::new()
        $g.Text        = $Text
        $g.Location    = [System.Drawing.Point]::new($X, $Y)
        $g.Size        = [System.Drawing.Size]::new($W, $H)
        $g.BackColor   = [System.Drawing.Color]::FromArgb(245, 245, 245)
        $g.ForeColor   = [System.Drawing.Color]::FromArgb(80, 80, 80)
        return $g
    }

    # ── Section: Connection ───────────────────────────────────────────────────
    $grpConn = New-GroupBox 'Connection' 12 10 648 90
    $form.Controls.Add($grpConn)

    $grpConn.Controls.Add((New-Label 'Mailbox user (UPN):' 10 24))
    $txtUser = New-TextBox 170 22 300 'user@company.com'
    $grpConn.Controls.Add($txtUser)

    $btnConnect = New-Button 'Connect-MgGraph' 480 21 150 26
    $btnConnect.BackColor = [System.Drawing.Color]::FromArgb(220, 240, 228)
    $btnConnect.ForeColor = [System.Drawing.Color]::FromArgb(20, 100, 55)
    $btnConnect.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(140, 200, 160)
    $grpConn.Controls.Add($btnConnect)

    $lblConnStatus = New-Label 'Not connected' 170 54 460 18
    $lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(160, 80, 0)
    $grpConn.Controls.Add($lblConnStatus)

    # ── Section: Paths ────────────────────────────────────────────────────────
    $grpPaths = New-GroupBox 'Paths' 12 108 648 90
    $form.Controls.Add($grpPaths)

    $grpPaths.Controls.Add((New-Label 'CSV file:' 10 24))
    $txtCsv = New-TextBox 170 22 280
    $grpPaths.Controls.Add($txtCsv)
    $btnBrowseCsv = New-Button 'Browse…' 458 21 80 26
    $grpPaths.Controls.Add($btnBrowseCsv)

    $grpPaths.Controls.Add((New-Label 'Output folder:' 10 58))
    $txtOutput = New-TextBox 170 56 280 'C:\TRANSFERSCRIPT\M365Export'
    $grpPaths.Controls.Add($txtOutput)
    $btnBrowseOut = New-Button 'Browse…' 458 55 80 26
    $grpPaths.Controls.Add($btnBrowseOut)

    # ── Section: Options ──────────────────────────────────────────────────────
    $grpOpts = New-GroupBox 'Options' 12 206 648 52
    $form.Controls.Add($grpOpts)

    $grpOpts.Controls.Add((New-Label 'Folder depth:' 10 20))
    $numDepth = [System.Windows.Forms.NumericUpDown]::new()
    $numDepth.Location = [System.Drawing.Point]::new(170, 17)
    $numDepth.Size     = [System.Drawing.Size]::new(55, 24)
    $numDepth.Minimum  = 1
    $numDepth.Maximum  = 6
    $numDepth.Value    = 3
    $grpOpts.Controls.Add($numDepth)

    $chkDryRun = [System.Windows.Forms.CheckBox]::new()
    $chkDryRun.Text     = 'Dry run  (preview only — no files saved or moved)'
    $chkDryRun.Location = [System.Drawing.Point]::new(270, 18)
    $chkDryRun.Size     = [System.Drawing.Size]::new(360, 22)
    $chkDryRun.Checked  = $true
    $chkDryRun.ForeColor = [System.Drawing.Color]::FromArgb(140, 90, 0)
    $grpOpts.Controls.Add($chkDryRun)

    # ── Section: Projects grid ────────────────────────────────────────────────
    $grpProj = New-GroupBox 'Projects' 12 266 648 230
    $form.Controls.Add($grpProj)

    $dgv                          = [System.Windows.Forms.DataGridView]::new()
    $dgv.Location                 = [System.Drawing.Point]::new(10, 20)
    $dgv.Size                     = [System.Drawing.Size]::new(625, 160)
    $dgv.BackgroundColor          = [System.Drawing.Color]::White
    $dgv.BorderStyle              = 'FixedSingle'
    $dgv.RowHeadersVisible        = $false
    $dgv.AllowUserToAddRows       = $true
    $dgv.AllowUserToDeleteRows    = $true
    $dgv.AutoSizeColumnsMode      = 'Fill'
    $dgv.SelectionMode            = 'FullRowSelect'
    $dgv.Font                     = [System.Drawing.Font]::new('Segoe UI', 9)
    $dgv.GridColor                = [System.Drawing.Color]::FromArgb(210, 210, 210)
    $dgv.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
    $dgv.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 248)
    $dgv.ColumnHeadersDefaultCellStyle.BackColor   = [System.Drawing.Color]::FromArgb(235, 235, 235)
    $dgv.ColumnHeadersDefaultCellStyle.ForeColor   = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $dgv.ColumnHeadersDefaultCellStyle.Font        = [System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $dgv.EnableHeadersVisualStyles = $false

    # Checkbox column for include/skip
    $colInclude             = [System.Windows.Forms.DataGridViewCheckBoxColumn]::new()
    $colInclude.HeaderText  = ''
    $colInclude.Width       = 28
    $colInclude.AutoSizeMode = 'None'
    $dgv.Columns.Add($colInclude) | Out-Null

    foreach ($col in @(
        @{ H='Category';       W=90 },
        @{ H='Project #';      W=70 },
        @{ H='Project Title';  W=0  },   # W=0 → Fill
        @{ H='Client Name';    W=0  }
    )) {
        $c = [System.Windows.Forms.DataGridViewTextBoxColumn]::new()
        $c.HeaderText = $col.H
        if ($col.W -gt 0) {
            $c.AutoSizeMode = 'None'
            $c.Width        = $col.W
        } else {
            $c.AutoSizeMode = 'Fill'
        }
        $dgv.Columns.Add($c) | Out-Null
    }

    $grpProj.Controls.Add($dgv)

    $btnAddRow = New-Button 'Add row' 10 188 80 26
    $grpProj.Controls.Add($btnAddRow)

    $btnDelRow = New-Button 'Delete row' 98 188 80 26
    $grpProj.Controls.Add($btnDelRow)

    $btnLoadCsv = New-Button 'Load from CSV' 186 188 110 26
    $grpProj.Controls.Add($btnLoadCsv)

    $btnSaveCsv = New-Button 'Save to CSV' 304 188 100 26
    $grpProj.Controls.Add($btnSaveCsv)

    # ── Section: Run ──────────────────────────────────────────────────────────
    $grpRun = New-GroupBox 'Run' 12 504 648 80
    $form.Controls.Add($grpRun)

    $btnRun = New-Button 'Run Export' 10 22 120 30
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(30, 130, 80)
    $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(20, 100, 60)
    $btnRun.Font      = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $grpRun.Controls.Add($btnRun)

    $btnMove = New-Button 'Move Folder to PREDELETE' 140 22 190 30
    $btnMove.BackColor = [System.Drawing.Color]::FromArgb(180, 60, 30)
    $btnMove.ForeColor = [System.Drawing.Color]::White
    $btnMove.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(140, 40, 20)
    $btnMove.Font      = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $grpRun.Controls.Add($btnMove)

    $lblStatus = New-Label 'Ready.' 310 28 320 22
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $grpRun.Controls.Add($lblStatus)

    # ── Populate grid helper ──────────────────────────────────────────────────
    function Add-ProjectRow {
        param([bool]$Checked=$true,[string]$Cat='',[string]$Num='',[string]$Title='',[string]$Client='')
        $row = $dgv.Rows.Add()
        $dgv.Rows[$row].Cells[0].Value = $Checked
        $dgv.Rows[$row].Cells[1].Value = $Cat
        $dgv.Rows[$row].Cells[2].Value = $Num
        $dgv.Rows[$row].Cells[3].Value = $Title
        $dgv.Rows[$row].Cells[4].Value = $Client
    }

    function Load-GridFromCsv {
        param([string]$Path)
        try {
            $rows = Import-Csv -Path $Path
            $dgv.Rows.Clear()
            foreach ($r in $rows) {
                Add-ProjectRow -Checked $true `
                    -Cat    $r.Category `
                    -Num    $r.'Project Number' `
                    -Title  $r.'Project Title' `
                    -Client $r.'Client Name'
            }
            $lblStatus.Text = "Loaded $($rows.Count) project(s) from CSV."
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Could not load CSV:`n$_", 'Error', 'OK', 'Error') | Out-Null
        }
    }

    # Pre-populate with CSV if already set
    if ($CsvPath -and (Test-Path $CsvPath)) {
        $txtCsv.Text      = $CsvPath
        $txtCsv.ForeColor = [System.Drawing.Color]::Black
        Load-GridFromCsv $CsvPath
    } else {
        # Seed with sample rows
        Add-ProjectRow $true  'ARCHIVING' '6109' 'Plate Storage Rack Certifications - Auckland' 'Vulcan Steel Ltd.'
        Add-ProjectRow $true  'ARCHIVING' '6117' 'Plate Storage Rack Certifications - Tauranga' 'Vulcan Steel Ltd.'
        Add-ProjectRow $true  'ARCHIVING' '6130' 'River Crossing Platform Certification'         'CUDDON LTD'
        Add-ProjectRow $true  'ARCHIVING' '6367' 'Samsung Battery Cabinet - Seismic Mapping'    'Eaton Industries Company'
    }
    if ($OutputFolder) {
        $txtOutput.Text      = $OutputFolder
        $txtOutput.ForeColor = [System.Drawing.Color]::Black
    }
    if ($MailboxUserId) {
        $txtUser.Text      = $MailboxUserId
        $txtUser.ForeColor = [System.Drawing.Color]::Black
    }

    # ── Check existing Graph connection ───────────────────────────────────────
    try {
        $ctx = Get-MgContext
        if ($ctx) {
            $lblConnStatus.Text      = "Connected as: $($ctx.Account)"
            $lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(20, 120, 60)
        }
    } catch {}

    # ── Event handlers ────────────────────────────────────────────────────────
    $btnConnect.Add_Click({
        $lblStatus.Text = 'Connecting to Microsoft Graph…'
        try {
            Connect-MgGraph -Scopes 'Mail.ReadWrite','Mail.ReadWrite.Shared','User.Read' -ErrorAction Stop
            $ctx = Get-MgContext
            $lblConnStatus.Text      = "Connected as: $($ctx.Account)"
            $lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(20, 120, 60)
            $lblStatus.Text = 'Connected successfully.'
        } catch {
            $lblConnStatus.Text      = "Connection failed: $_"
            $lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 0)
            $lblStatus.Text = 'Connection failed.'
        }
    })

    $btnBrowseCsv.Add_Click({
        $ofd = [System.Windows.Forms.OpenFileDialog]::new()
        $ofd.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
        $ofd.Title  = 'Select project CSV file'
        if ($ofd.ShowDialog() -eq 'OK') {
            $txtCsv.Text      = $ofd.FileName
            $txtCsv.ForeColor = [System.Drawing.Color]::Black
            Load-GridFromCsv $ofd.FileName
        }
    })

    $btnBrowseOut.Add_Click({
        $fbd = [System.Windows.Forms.FolderBrowserDialog]::new()
        $fbd.Description = 'Select output folder'
        if ($txtOutput.ForeColor -ne [System.Drawing.Color]::Gray) { $fbd.SelectedPath = $txtOutput.Text }
        if ($fbd.ShowDialog() -eq 'OK') {
            $txtOutput.Text      = $fbd.SelectedPath
            $txtOutput.ForeColor = [System.Drawing.Color]::Black
        }
    })

    $btnAddRow.Add_Click({ Add-ProjectRow })

    $btnDelRow.Add_Click({
        foreach ($row in $dgv.SelectedRows) {
            if (-not $row.IsNewRow) { $dgv.Rows.Remove($row) }
        }
    })

    $btnLoadCsv.Add_Click({
        if ($txtCsv.ForeColor -ne [System.Drawing.Color]::Gray -and $txtCsv.Text) {
            Load-GridFromCsv $txtCsv.Text
        } else {
            $btnBrowseCsv.PerformClick()
        }
    })

    $btnSaveCsv.Add_Click({
        $sfd = [System.Windows.Forms.SaveFileDialog]::new()
        $sfd.Filter   = 'CSV files (*.csv)|*.csv'
        $sfd.FileName = 'ProjectsToArchive.csv'
        if ($sfd.ShowDialog() -eq 'OK') {
            try {
                $lines = @('Category,Project Number,Project Title,Client Name')
                foreach ($row in $dgv.Rows) {
                    if ($row.IsNewRow) { continue }
                    $cat    = $row.Cells[1].Value
                    $num    = $row.Cells[2].Value
                    $title  = $row.Cells[3].Value
                    $client = $row.Cells[4].Value
                    $lines += "$cat,$num,`"$title`",`"$client`""
                }
                $lines | Set-Content -Path $sfd.FileName -Encoding UTF8
                $lblStatus.Text = "CSV saved to $($sfd.FileName)"
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Save failed:`n$_", 'Error', 'OK', 'Error') | Out-Null
            }
        }
    })

    $btnRun.Add_Click({
        # ── Validate ──────────────────────────────────────────────────────────
        $user = $txtUser.Text.Trim()
        if ($txtUser.ForeColor -eq [System.Drawing.Color]::Gray -or -not $user) {
            [System.Windows.Forms.MessageBox]::Show('Please enter a mailbox user (UPN or Object ID).', 'Validation', 'OK', 'Warning') | Out-Null
            return
        }
        $outDir = $txtOutput.Text.Trim()
        if ($txtOutput.ForeColor -eq [System.Drawing.Color]::Gray -or -not $outDir) {
            [System.Windows.Forms.MessageBox]::Show('Please select an output folder.', 'Validation', 'OK', 'Warning') | Out-Null
            return
        }
        try { $ctx = Get-MgContext; if (-not $ctx) { throw } } catch {
            [System.Windows.Forms.MessageBox]::Show('Not connected to Microsoft Graph. Click "Connect-MgGraph" first.', 'Not Connected', 'OK', 'Warning') | Out-Null
            return
        }

        # ── Collect project rows ──────────────────────────────────────────────
        $projectList = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($row in $dgv.Rows) {
            if ($row.IsNewRow) { continue }
            if (-not $row.Cells[0].Value) { continue }
            $projectList.Add([PSCustomObject]@{
                Category      = $row.Cells[1].Value
                'Project Number' = $row.Cells[2].Value
                'Project Title'  = $row.Cells[3].Value
                'Client Name'    = $row.Cells[4].Value
            })
        }
        if ($projectList.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('No projects selected. Check at least one row.', 'Validation', 'OK', 'Warning') | Out-Null
            return
        }

        $depth  = [int]$numDepth.Value
        $dryRun = $chkDryRun.Checked

        # ── Disable UI during run ─────────────────────────────────────────────
        $btnRun.Enabled  = $false
        $btnMove.Enabled = $false
        $lblStatus.Text  = 'Running export… see terminal for progress.'
        $form.Refresh()

        try {
            Invoke-ExportRun `
                -MailboxUserId    $user `
                -OutputFolder     $outDir `
                -ProjectList      $projectList `
                -FolderMatchDepth $depth `
                -IsDryRun         $dryRun
            $lblStatus.Text = if ($dryRun) { 'Dry run complete.' } else { 'Export complete.' }
        } catch {
            $lblStatus.Text = "Error: $_"
            [System.Windows.Forms.MessageBox]::Show("Export failed:`n$_", 'Error', 'OK', 'Error') | Out-Null
        } finally {
            $btnRun.Enabled  = $true
            $btnMove.Enabled = $true
        }
    })

    $btnMove.Add_Click({
        # ── Validate ──────────────────────────────────────────────────────────
        $user = $txtUser.Text.Trim()
        if ($txtUser.ForeColor -eq [System.Drawing.Color]::Gray -or -not $user) {
            [System.Windows.Forms.MessageBox]::Show('Please enter a mailbox user (UPN or Object ID).', 'Validation', 'OK', 'Warning') | Out-Null
            return
        }
        try { $ctx = Get-MgContext; if (-not $ctx) { throw } } catch {
            [System.Windows.Forms.MessageBox]::Show('Not connected to Microsoft Graph. Click "Connect-MgGraph" first.', 'Not Connected', 'OK', 'Warning') | Out-Null
            return
        }

        # ── Collect project rows ──────────────────────────────────────────────
        $projectList = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($row in $dgv.Rows) {
            if ($row.IsNewRow) { continue }
            if (-not $row.Cells[0].Value) { continue }
            $projectList.Add([PSCustomObject]@{
                Category         = $row.Cells[1].Value
                'Project Number' = $row.Cells[2].Value
                'Project Title'  = $row.Cells[3].Value
                'Client Name'    = $row.Cells[4].Value
            })
        }
        if ($projectList.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('No projects selected. Check at least one row.', 'Validation', 'OK', 'Warning') | Out-Null
            return
        }

        $outDir = $txtOutput.Text.Trim()
        if ($txtOutput.ForeColor -eq [System.Drawing.Color]::Gray -or -not $outDir) {
            [System.Windows.Forms.MessageBox]::Show('Please select an output folder for logs.', 'Validation', 'OK', 'Warning') | Out-Null
            return
        }

        $dryRun = $chkDryRun.Checked
        $verb   = if ($dryRun) { 'preview moving' } else { 'MOVE' }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "This will $verb all emails from the matched project folders to Inbox\PREDELETE.`n`nContinue?",
            'Confirm Folder Move',
            'YesNo',
            $(if ($dryRun) { 'Question' } else { 'Warning' })
        )
        if ($confirm -ne 'Yes') { return }

        # ── Disable UI during run ─────────────────────────────────────────────
        $btnRun.Enabled  = $false
        $btnMove.Enabled = $false
        $lblStatus.Text  = 'Moving emails… see terminal for progress.'
        $form.Refresh()

        try {
            Invoke-MoveToPredelete `
                -MailboxUserId    $user `
                -OutputFolder     $outDir `
                -ProjectList      $projectList `
                -FolderMatchDepth ([int]$numDepth.Value) `
                -IsDryRun         $dryRun `
                -CsvSourcePath    $csv
            $lblStatus.Text = if ($dryRun) { 'Dry run complete — no folders moved.' } else { 'Move folders to PREDELETE complete.' }
        } catch {
            $lblStatus.Text = "Error: $_"
            [System.Windows.Forms.MessageBox]::Show("Move failed:`n$_", 'Error', 'OK', 'Error') | Out-Null
        } finally {
            $btnRun.Enabled  = $true
            $btnMove.Enabled = $true
        }
    })

    $form.ShowDialog() | Out-Null
    $form.Dispose()
}

# ══════════════════════════════════════════════════════════════════════════════
#  CORE EXPORT RUNNER  (used by both GUI and headless paths)
# ══════════════════════════════════════════════════════════════════════════════

function Invoke-ExportRun {
    param(
        [string]$MailboxUserId,
        [string]$OutputFolder,
        [object[]]$ProjectList,
        [int]$FolderMatchDepth,
        [bool]$IsDryRun
    )

    if ($IsDryRun) {
        Write-Host ""
        Write-Host "╔══════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "║           DRY RUN MODE (-WhatIf)             ║" -ForegroundColor Yellow
        Write-Host "║  No files will be created or downloaded.     ║" -ForegroundColor Yellow
        Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host ""
    }

    $logContext = New-RunLogContext -OutputFolder $OutputFolder -Mode 'Export' -MailboxUserId $MailboxUserId
    Write-Host "Logging to: $($logContext.RunDir)" -ForegroundColor DarkGray

    $mgContext = Get-MgContext
    if (-not $mgContext) {
        throw "Not connected to Microsoft Graph. Run: Connect-MgGraph -Scopes 'Mail.ReadWrite','Mail.ReadWrite.Shared','User.Read'"
    }
    Write-Host "Connected as: $($mgContext.Account)"

    Write-Host "Retrieving access token..."
    $Script:AccessToken = Get-MgAccessToken -UserId $MailboxUserId
    Write-Host "Access token acquired. $(($Script:AccessToken).Substring(0,20))..."

    if (-not $IsDryRun) {
        if (-not (Test-Path $OutputFolder)) {
            New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
            Write-Host "Created output folder: $OutputFolder"
        }
    }

    Write-Host "Loaded $($ProjectList.Count) project(s)."
    Write-Host "`nScanning mailbox folders (depth: $FolderMatchDepth)..."
    $allFolders = Get-AllMailFolders -UserId $MailboxUserId -MaxDepth $FolderMatchDepth
    Write-Host "Found $($allFolders.Count) total mail folder(s)."

    $totalSaved   = 0
    $totalFailed  = 0
    $totalThreads = 0
    $totalWhatIf  = 0

    foreach ($project in $ProjectList) {
        $projectNumber = $project.'Project Number'.Trim()
        $projectTitle  = $project.'Project Title'.Trim()
        $clientName    = $project.'Client Name'.Trim()

        Write-Host "`n──────────────────────────────────────────────"
        Write-Host "Project : $projectNumber - $projectTitle"
        Write-Host "Client  : $clientName"

        $matchedFolder = Find-ProjectFolder -AllFolders $allFolders -ProjectNumber $projectNumber

        if (-not $matchedFolder) {
            Write-Warning "  No Outlook folder found matching '$projectNumber'. Skipping."
            continue
        }

        Write-Host "  Matched folder : '$($matchedFolder.DisplayName)'"

        # ── New structure: OutputFolder \ ClientName \ ProjectNumber - Title ──
        $safeClient      = Get-SafeName $clientName  -MaxLength 60
        $safeProjFolder  = Get-SafeName "$projectNumber - $projectTitle" -MaxLength 80
        $projectFolder   = Join-Path $OutputFolder $safeClient $safeProjFolder

        if (-not $IsDryRun -and -not (Test-Path $projectFolder)) {
            New-Item -ItemType Directory -Path $projectFolder -Force | Out-Null
        }

        $result = Export-FolderMessages `
            -UserId            $MailboxUserId `
            -FolderId          $matchedFolder.Id `
            -DestinationFolder $projectFolder `
            -AccessToken       $Script:AccessToken `
            -IsDryRun          $IsDryRun

        if ($IsDryRun) {
            Write-Host ("  [WHATIF] Would download {0} email(s) across {1} thread(s)" -f `
                $result.WhatIfCount, $result.Threads) -ForegroundColor Yellow
            Write-Host "  [WHATIF] Path: $projectFolder" -ForegroundColor DarkYellow
            Write-ProjectLog -LogContext $logContext -MailboxUserId $MailboxUserId -Mode 'Export' -ClientName $clientName -ProjectNumber $projectNumber -ProjectTitle $projectTitle -MatchedFolder $matchedFolder.DisplayName -Action 'ExportMessages' -Status 'WhatIf' -MessageCount $result.WhatIfCount -ChildFolderCount $matchedFolder.ChildCount -SourceFolderId $matchedFolder.Id -DestinationPath $projectFolder
        } else {
            Write-Host "  Saved: $($result.Saved)  |  Failed: $($result.Failed)  |  Threads: $($result.Threads)"
            $status = if ($result.Failed -gt 0) { 'Completed-WithFailures' } else { 'Completed' }
            Write-ProjectLog -LogContext $logContext -MailboxUserId $MailboxUserId -Mode 'Export' -ClientName $clientName -ProjectNumber $projectNumber -ProjectTitle $projectTitle -MatchedFolder $matchedFolder.DisplayName -Action 'ExportMessages' -Status $status -MessageCount $result.Saved -ChildFolderCount $matchedFolder.ChildCount -SourceFolderId $matchedFolder.Id -DestinationPath $projectFolder -Error "Failed=$($result.Failed); Threads=$($result.Threads)"
        }

        $totalSaved   += $result.Saved
        $totalFailed  += $result.Failed
        $totalThreads += $result.Threads
        $totalWhatIf  += $result.WhatIfCount
    }

    Write-Host ""
    Write-Host "══════════════════════════════════════════════"
    if ($IsDryRun) {
        Write-Host "DRY RUN complete. No files were downloaded." -ForegroundColor Yellow
        Write-Host "  Threads that would be created : $totalThreads"
        Write-Host "  Emails that would be saved    : $totalWhatIf"
        Write-Host ""
        Write-Host "  To run for real, uncheck 'Dry run' and click Run Export." -ForegroundColor Green    } else {
        Write-Host "Export complete."
        Write-Host "  Conversation threads : $totalThreads"
        Write-Host "  Total emails saved   : $totalSaved"
        Write-Host "  Total emails failed  : $totalFailed"
        Write-Host "  Output folder        : $OutputFolder"
    }
    Write-Host "══════════════════════════════════════════════"
    Write-RunLog -LogContext $logContext -MailboxUserId $MailboxUserId -Mode 'Export' -Status 'Completed' -Details "Saved=$totalSaved; Failed=$totalFailed; Threads=$totalThreads; WhatIf=$totalWhatIf; Output=$OutputFolder"
}

# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

$headless = $MailboxUserId -and $CsvPath -and $OutputFolder

if ($headless) {
    # ── Headless / command-line mode ──────────────────────────────────────────
    $isDryRun    = [bool]$WhatIfPreference
    $projectList = Import-Csv -Path $CsvPath

    if ($MoveToPredelete) {
        Invoke-MoveToPredelete `
            -MailboxUserId    $MailboxUserId `
            -OutputFolder     $OutputFolder `
            -ProjectList      $projectList `
            -FolderMatchDepth $FolderMatchDepth `
            -IsDryRun         $isDryRun `
            -CsvSourcePath    $CsvPath
    } else {
        Invoke-ExportRun `
            -MailboxUserId    $MailboxUserId `
            -OutputFolder     $OutputFolder `
            -ProjectList      $projectList `
            -FolderMatchDepth $FolderMatchDepth `
            -IsDryRun         $isDryRun
    }
} else {
    # ── GUI mode ──────────────────────────────────────────────────────────────
    Show-ExportGui
}
