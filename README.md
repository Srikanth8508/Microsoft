This script automates the extraction of emails from specific Outlook subfolders into individual CSV files. It captures full SMTP email addresses (even from internal Exchange accounts) and cleans the email body for perfect CSV formatting.

Open Outlook: You must have Outlook open and be logged into cd@fructidor.com.

Open PowerShell: Search for "PowerShell" in the Start menu.

Enable Scripting: Paste this command first and hit Enter:
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

Run the Export: Paste the long script above and hit Enter.

Outlook Security Prompt:

  * Because the script is reading email data, Outlook will show a security warning: "A program is trying to access email address information..."
  * Check the box "Allow access for".
  * Select 10 minutes (since you have 500+ folders, it will take a few minutes to read them all).
  * Click Allow.

What will be inside the CSVs?
Each CSV (e.g., 001.csv) will now have a header row and a list of all emails in that folder with these columns:

  * Subject: The title of the email.
  * Sender: Who sent it.
  * Received: The date and time it arrived.
  * Size_Bytes: How large the email is.
  * To: Who it was sent to.
  * Body: The complete text content of the email.
