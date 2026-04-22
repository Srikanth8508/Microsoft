# ==============================================================================
# STEP 1: Connect to the Outlook Application
# ==============================================================================
# This "wakes up" the Outlook app already running on your computer.
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# ==============================================================================
# STEP 2: Configure Your Targets
# ==============================================================================
$accountName = "<<username>>@gmail.com"        # Your email address as seen in Outlook
$folderName = "Mail_backup_userwise"     # The top-level folder containing subfolders

try {
    # Navigate the folder tree: Account -> Target Folder
    $root = $namespace.Folders.Item($accountName)
    $targetFolder = $root.Folders.Item($folderName)
} catch {
    Write-Host "❌ Error: Could not find folder. Is Outlook open?" -ForegroundColor Red
    return
}

# ==============================================================================
# STEP 3: Setup Destination Path
# ==============================================================================
# Creates a 'Mail_backup' folder in your personal Documents folder.
$destinationPath = Join-Path -Path ([Environment]::GetFolderPath("MyDocuments")) -ChildPath "Mail_backup"
if (!(Test-Path $destinationPath)) { 
    New-Item -ItemType Directory -Path $destinationPath | Out-Null 
}

Write-Host "🚀 Starting export to: $destinationPath" -ForegroundColor Yellow

# ==============================================================================
# STEP 4: Loop Through Each Subfolder (001, 002, etc.)
# ==============================================================================
foreach ($subfolder in $targetFolder.Folders) {
    $csvFileName = "$($subfolder.Name).csv"
    $fullPath = Join-Path -Path $destinationPath -ChildPath $csvFileName

    $results = foreach ($mail in $subfolder.Items) {
        
        # --- Handle 'From' Address ---
        # If it's an internal Exchange mail, resolve it to a standard @domain.com address
        $fromAddress = $mail.SenderEmailAddress
        if ($mail.SenderEmailType -eq "EX") {
            $sender = $mail.Sender.GetExchangeUser()
            if ($null -ne $sender) { $fromAddress = $sender.PrimarySmtpAddress }
        }

        # --- Handle 'To' Address ---
        # Extracts all recipients and joins them with a semicolon
        $toAddresses = @()
        foreach ($recipient in $mail.Recipients) {
            if ($recipient.AddressEntry.Type -eq "EX") {
                $user = $recipient.AddressEntry.GetExchangeUser()
                if ($null -ne $user) { $toAddresses += $user.PrimarySmtpAddress }
                else { $toAddresses += $recipient.Address }
            } else {
                $toAddresses += $recipient.Address
            }
        }

        # --- Clean Body Content ---
        # Removes line breaks so the CSV row stays on one line in Excel
        $cleanBody = ""
        if ($mail.Body) {
            $cleanBody = $mail.Body -replace "`r`n", " " -replace "`n", " " -replace "`r", " "
        }

        # Create the data row
        [PSCustomObject]@{
            From       = $fromAddress
            To         = ($toAddresses -join "; ")
            Subject    = $mail.Subject
            Received   = $mail.ReceivedTime
            Body       = $cleanBody
            Size_KB    = [math]::Round($mail.Size / 1KB, 2)
        }
    }

    # ==============================================================================
    # STEP 5: Save Results
    # ==============================================================================
    if ($results) {
        # Export data to CSV with UTF8 encoding for special characters
        $results | Export-Csv -Path $fullPath -NoTypeInformation -Encoding UTF8
        Write-Host "✅ Exported: $csvFileName ($($subfolder.Items.Count) emails)" -ForegroundColor Cyan
    } else {
        # If folder is empty, create an empty file placeholder
        New-Item -Path $fullPath -ItemType File -Force | Out-Null
        Write-Host "⚪ Folder Empty: $($subfolder.Name)" -ForegroundColor Gray
    }
}

Write-Host "`n✨ Export Complete!" -ForegroundColor Green
explorer $destinationPath
