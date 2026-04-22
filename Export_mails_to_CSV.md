# 📧 Outlook Mail Backup Script (PowerShell)

A clean and structured script to export Outlook emails from subfolders into CSV files.

---

## 🔹 Overview

* Connects to Outlook
* Targets a specific mailbox & folder
* Exports each subfolder into a CSV
* Saves output to **Documents/Mail_backup**

---

## ⚙️ Configuration

```powershell
$accountName = "<<username>>@gmail.com"   # Outlook account
$folderName  = "Mail_backup_userwise"     # Parent folder containing subfolders
```

---

## 🚀 Script

```powershell
# ==========================================================
# 1. CONNECT TO OUTLOOK
# ==========================================================
$outlook   = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# ==========================================================
# 2. GET TARGET FOLDER
# ==========================================================
try {
    $root         = $namespace.Folders.Item($accountName)
    $targetFolder = $root.Folders.Item($folderName)
} catch {
    Write-Host "❌ Folder not found. Ensure Outlook is open." -ForegroundColor Red
    return
}

# ==========================================================
# 3. PREPARE DESTINATION
# ==========================================================
$destinationPath = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "Mail_backup"

if (!(Test-Path $destinationPath)) {
    New-Item -ItemType Directory -Path $destinationPath | Out-Null
}

Write-Host "🚀 Exporting to: $destinationPath" -ForegroundColor Yellow

# ==========================================================
# 4. PROCESS EACH SUBFOLDER
# ==========================================================
foreach ($subfolder in $targetFolder.Folders) {

    $filePath = Join-Path $destinationPath "$($subfolder.Name).csv"

    $results = foreach ($mail in $subfolder.Items) {

        # ---------- FROM ----------
        $from = $mail.SenderEmailAddress
        if ($mail.SenderEmailType -eq "EX") {
            $sender = $mail.Sender.GetExchangeUser()
            if ($sender) { $from = $sender.PrimarySmtpAddress }
        }

        # ---------- TO ----------
        $to = foreach ($r in $mail.Recipients) {
            if ($r.AddressEntry.Type -eq "EX") {
                $u = $r.AddressEntry.GetExchangeUser()
                if ($u) { $u.PrimarySmtpAddress } else { $r.Address }
            } else {
                $r.Address
            }
        }

        # ---------- BODY CLEAN ----------
        $body = ""
        if ($mail.Body) {
            $body = $mail.Body -replace "`r`n|`n|`r", " "
        }

        # ---------- OUTPUT ----------
        [PSCustomObject]@{
            From     = $from
            To       = ($to -join "; ")
            Subject  = $mail.Subject
            Received = $mail.ReceivedTime
            Body     = $body
            Size_KB  = [math]::Round($mail.Size / 1KB, 2)
        }
    }

    # ======================================================
    # 5. EXPORT
    # ======================================================
    if ($results) {
        $results | Export-Csv $filePath -NoTypeInformation -Encoding UTF8
        Write-Host "✅ Exported: $($subfolder.Name)" -ForegroundColor Cyan
    } else {
        New-Item $filePath -ItemType File -Force | Out-Null
        Write-Host "⚪ Empty: $($subfolder.Name)" -ForegroundColor Gray
    }
}

# ==========================================================
# 6. COMPLETE
# ==========================================================
Write-Host "`n✨ Export Complete!" -ForegroundColor Green
explorer $destinationPath
```

---

## ✅ Improvements Made

* Removed repeated separators
* Grouped logic into clear sections
* Simplified variable naming
* Reduced comments to only useful ones
* Compact regex for body cleanup
* Cleaner output logs

---

## 📁 Output Example

```
Documents/
└── Mail_backup/
    ├── 001.csv
    ├── 002.csv
    └── 003.csv
```

---

## 💡 Notes

* Works only when **Outlook is open**
* Supports **Exchange + SMTP emails**
* Handles **empty folders safely**
* CSV is **Excel-friendly (UTF-8)**

---
