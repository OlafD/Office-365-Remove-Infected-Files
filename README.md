# Office-365-Remove-Infected-Files
PowerShell script to remove malware infected files by the export of the audit log in Office 365. The script will cover OneDrive, SharePoint Online and Teams sites (sites assigned to a Group).

Using the script ExportMalwareDetections.ps1, the csv-file is created that will be processed for the removal of the files in the tenant. 

The script RemoveInfectedFiles.ps1 processes the output (csv-file) and deletes the files reported in the audit log. The script must run at least as a SharePoint Administrator in the tenant. Before anything is done with a file, the script promotes the user, who is running the script, to a site collection administrator to have sufficient permissions. When the action is done for a file, the user will be removed from the site collection administrators of the specific site collection.
When a file is deleted, an email notification is sent. Depending on the type of the site, the notification is sent to
- the site collection administrator for a OneDrive
- the owners group for a Teams site
- the members of the associated owners group for a SharePoint Online site

The RemoveInfectedFiles.Param.xml file is used for parameters, needed in the script but normally not changed for each execution. In this file set
- the tenant url (https://company-admin.sharepoint.com)
- the sender address for the notification mails
- the smtp server for the notification mails
- the smtp port for the notification mails

When the script is executed, a transcript is written to the folder specified, when the script was started.
