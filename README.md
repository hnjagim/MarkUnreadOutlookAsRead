# MarkUnreadOutlookAsRead
Using PowerShell to mark all unread messages as read

1. Open PowerShell on your computer.
2. Run the following command to import the Outlook module:

```powershell
Import-Module Outlook
```

3. Run the following command to connect to your Outlook application:

```powershell
$Outlook = New-Object -ComObject Outlook.Application
```

4. Run the following command to access the Inbox folder:

```powershell
$Inbox = $Outlook.Session.GetDefaultFolder(6)
```

Note: If your unread messages are in a different folder, replace `6` with the appropriate folder ID. You can find the folder ID using the following command:

```powershell
$Outlook.Session.Folders | Select-Object Name, EntryID
```

5. Run the following command to iterate through the unread messages in the Inbox folder and mark them as read:

```powershell
$Inbox.Items | Where-Object { $_.UnRead -eq $true } | ForEach-Object { $_.UnRead = $false; $_.Save() }
```

This PowerShell script should mark all unread messages in the specified folder as read.

If you still encounter issues, please provide more details or error messages, and I'll be glad to assist you further.
