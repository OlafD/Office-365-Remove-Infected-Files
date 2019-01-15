param (
    [Parameter(Mandatory=$true)]
    [string]$InputFile,
    [Parameter(Mandatory=$true)]
    [string]$TranscriptPath,
    $Credential
)

# ------------ Constants -------------

$mailSubject = "Notification - Infected file removed"

$mailBody = @"
<div>Hello,</div>
<br />
<div>this automatic E-mail is send to you as a notification for the removal of the file [*file_placeholder*] in the site [*site_placeholder*] because a virus infection was detected in the file. 
<br />
<br />
<div>Kind regards</div>
<br />
<div>Antivirus Service.</div>
"@

# ------------ Functions -------------

function FindNthOccurence
{
    param (
        [Parameter(Mandatory=$true)]
        [string]$StringValue,
        [Parameter(Mandatory=$true)]
        [string]$SearchFor,
        [Parameter(Mandatory=$true)]
        [int]$Occurrence
    )

    [int]$i = 0
    [int]$pos = 0

    while ($pos -ge 0)
    {
        $pos = $StringValue.IndexOf($SearchFor, $pos)

        if ($pos -ge 0)
        {
            $i++
        }
        else
        {
            return -1
        }

        if ($i -eq $Occurrence)
        {
            return $pos
        }

        $pos++
    }

    return -1
}

function SetSiteCollectionAdmin
{
    param (
        [Parameter(Mandatory=$true)]
        $TenantObject,
        [Parameter(Mandatory=$true)]
        $TenantConnection,
        [Parameter(Mandatory=$true)]
        [string]$Url,
        [Parameter(Mandatory=$true)]
        [string]$User
    )

    $TenantObject.SetSiteAdmin($Url, $User, $true) | Out-Null

    Invoke-PnPQuery -Connection $TenantConnection

    Write-Host "$User added as site collection admin"
}

function RemoveSiteCollectionAdmin
{
    param (
        [Parameter(Mandatory=$true)]
        $TenantObject,
        [Parameter(Mandatory=$true)]
        $TenantConnection,
        [Parameter(Mandatory=$true)]
        [string]$Url,
        [Parameter(Mandatory=$true)]
        [string]$User
    )

    $TenantObject.SetSiteAdmin($Url, $User, $false) | Out-Null

    Invoke-PnPQuery -Connection $TenantConnection

    Write-Host "$User removed from site collection admins"
}

function GetSiteCollectionUrl
{
    param (
        [Parameter(Mandatory=$true)]
        [string]$Url
    )

    $result = ""

    $i = FindNthOccurence -StringValue $Url -SearchFor "/" -Occurrence 5

    if ($i -ge 0)
    {
        $result = $Url.Substring(0, $i)
    }

    return $result
}

function GetServerRelativeUrl
{
    param (
        [Parameter(Mandatory=$true)]
        [string]$Url
    )

    $result = ""

    $i = FindNthOccurence -StringValue $Url -SearchFor "/" -Occurrence 5

    if ($i -ge 0)
    {
        $result = $Url.Substring($i)
    }

    return $result
}

function RemoveFile
{
    param (
        [Parameter(Mandatory=$true)]
        [string]$FileUrl,
        [Parameter(Mandatory=$true)]
        $Connection
    )

    $result = $false

    $file = Get-PnPFile -Url $FileUrl -Connection $Connection -ErrorAction SilentlyContinue

    if ($file -ne $null)
    {
        $serverRelativeFileUrl = Get-PnPProperty -ClientObject $file -Property ServerRelativeUrl -Connection $Connection  # get the value of the property

        Write-Host -NoNewline "Remove file $serverRelativeFileUrl..."

        Remove-PnPFile -ServerRelativeUrl $serverRelativeFileUrl -Recycle -Force -Connection $Connection

        Write-Host "Done."

        $result = $true
    }
    else
    {
        Write-Host "File $fileUrl does not exist anymore."

        $result = $false
    }

    return $result
}

function GetMailReceipients
{
    param (
        [Parameter(Mandatory=$true)]
        $ValueObject,
        [Parameter(Mandatory=$true)]
        $Connection
    )

    $result = ""

    $web = Get-PnPWeb -Connection $Connection
    Get-PnPProperty -ClientObject $web -Property WebTemplate,Configuration -Connection $Connection

    if (($web.WebTemplate -eq "GROUP") -and ($web.Configuration -eq "0"))  # Teams
    {
        $admin = Get-PnPSiteCollectionAdmin -Connection $Connection | Where { $_.Title -like "*Owners" }

        if ($admin -ne $null)
        {
            $result = $admin[0].Email
        }
    }
    elseif ($web.WebTemplate -eq  "SPSPERS")  # OneDrive
    {
        $admin = Get-PnPSiteCollectionAdmin -Connection $Connection | Where { $_.Email -ne $Connection.PSCredential.UserName }

        if ($admin -ne $null)
        {
            $result = $admin[0].Email
        }
    }
    else  # any other SharePoint site
    {
        $owners = Get-PnPGroup -AssociatedOwnerGroup -Connection $Connection

        if ($owners -ne $null)
        {
            foreach ($user in $owners.Users)
            {
                if ($user.LoginName -ne "SHAREPOINT\system")
                {
                    if ($result -eq "")
                    {
                        $result = $user.Email
                    }
                    else
                    {
                        $result = $result + ";" + $user.Email
                    }
                }
            }
        }
    }

    return $result
}

function GetValueFromXml
{
       param (
             [Parameter(Mandatory=$true)]
             [string]$NodeName,
             [string]$XmlFilename = "$PSScriptRoot\RemoveInfectedFiles.Param.xml"
       )

       $result = ""

       $xmlDoc = New-Object System.Xml.XmlDocument

       $xmlDoc.Load($XmlFilename)

       $rootNode = $xmlDoc.DocumentElement

       $node = $rootNode.SelectSingleNode("//RemoveInfectedFiles/$NodeName")

       if ($node -ne $null)
       {
             $result = $node.InnerText
       }

       return $result
}

function SendNotificationMail
{
    param (
        [Parameter(Mandatory=$true)]
        [string]$Receipients,
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        [Parameter(Mandatory=$true)]
        [string]$Body,
        [Parameter(Mandatory=$true)]
        $Credentials
    )

    $to = $Receipients.Split(";")

    if ($to.Count -gt 0)
    {
        $smtpServer = GetValueFromXml -NodeName "SmtpServer"
        $smtpPort = GetValueFromXml -NodeName "SmtpPort"
        $from = GetValueFromXml "MailFrom"

    	Send-MailMessage -From $from -To $to -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $smtpServer -Port $smtpPort -Credential $Credentials -UseSsl

        Write-Host "Notification mail sent."
    }
    else
    {
        Write-Host "No receipients specified to send mail to."
    }
}

# ------------ Main -------------

if ($Credential -eq $null)
{
    $Credential = Get-Credential
}

$tenantAdmin = $Credential.UserName

$transcriptExtension = Get-Date -Format yyyyMMdd-HHmmss
$transcriptFile = "$TranscriptPath\SAP_RemoveInfectedFiles_Transcript_$transcriptExtension.txt"

Start-Transcript -Path $transcriptFile


Write-Host "Script is running with user $tenantAdmin"

$TenantUrl = GetValueFromXml -NodeName "TenantUrl"

$tenantConnection = Connect-PnPOnline -Url $TenantUrl -Credentials $Credential -ReturnConnection

[Reflection.Assembly]::LoadWithPartialName("Microsoft.Online.SharePoint.TenantAdministration")

$ctx = Get-PnPContext
$tenant = New-Object Microsoft.Online.SharePoint.TenantAdministration.Tenant($ctx)
$ctx.Load($tenant)
Invoke-PnPQuery -Connection $tenantConnection


$csvContent = Import-Csv -Path $InputFile -Delimiter ";"

foreach ($csvItem in $csvContent)
{
    if (($csvItem.Workload -eq "SharePoint") -or ($csvItem.Workload -eq "OneDrive"))
    {
# store csv item values to variables
        $creationTime = $csvItem.CreationTime
        $id = $csvItem.Id
        $operation = $csvItem.Operation
        $OrganizationId = $csvItem.OrganizationId
        $recordType = $csvItem.RecordType
        $userKey = $csvItem.UserKey
        $userType = $csvItem.UserType
        $version = $csvItem.Version
        $workload = $csvItem.Workload
        $clientIP = $csvItem.ClientIP
        $objectId = $csvItem.ObjectId
        $userId = $csvItem.UserId
        $correlationId = $csvItem.CorrelationId
        $eventSource = $csvItem.EventSource
        $itemType = $csvItem.ItemType
        $listId = $csvItem.ListId
        $listItemUniqueId = $csvItem.ListItemUniqueId
        $site = $csvItem.Site
        $userAgent = $csvItem.UserAgent
        $webId = $csvItem.WebId
        $sourceFileExtension = $csvItem.SourceFileExtension
        $virusInfo = $csvItem.VirusInfo
        $virusVendor = $csvItem.VirusVendor
        $siteUrl = $csvItem.SiteUrl
        $sourceFileName = $csvItem.SourceFileName
        $sourceRelativeUrl = $csvItem.SourceRelativeUrl

        $fileUrl = "/" + $sourceRelativeUrl + "/" + $sourceFileName

        Write-Host -ForegroundColor Yellow "File: $objectId"

# make site collection admin

        SetSiteCollectionAdmin -TenantObject $tenant -TenantConnection $tenantConnection -Url $siteUrl -User $tenantAdmin

# connect to site url

        $siteConnection = Connect-PnPOnline -Url $siteUrl -Credentials $Credential -ReturnConnection

        Write-Host "Connected to site collection"

# connect to the web with the affected file

        $ctx = Get-PnPContext
        $site = $ctx.Site

        $web = $site.OpenWebById($webId)
        $ctx.Load($web)
        Invoke-PnPQuery

        $webUrl = $web.Url

        $webConnection = Connect-PnPOnline -Url $webUrl -Credentials $Credential -ReturnConnection

        Write-Host "Connected to web $weburl"

        Disconnect-PnPOnline -Connection $siteConnection

        Write-Host "Disconnected from site collection"

# collect mail receipients

        $mailReceipients = GetMailReceipients -ValueObject $csvItem -Connection $webConnection

        Write-Host "Found mail receipients: $mailReceipients"

# remove the file

        $success = RemoveFile -FileUrl $fileUrl -Connection $webConnection

# send email

        if ($success -eq $true)
        {
            $subject = $mailSubject
            $body = $mailBody.Replace("[*site_placeholder*]", $siteUrl)
            $body = $body.Replace("[*file_placeholder*]", $objectId)

            if ($mailReceipients -ne "")
            {
                SendNotificationMail -Receipients $mailReceipients -Subject $subject -Body $body -Credentials $Cred
            }
            else
            {
                Write-Host "No receipients found, no mail sent."
            }
        }

# remove from site collection admins

        $admins = Get-PnPSiteCollectionAdmin -Connection $siteConnection

        if ($admins.Count -gt 1)
        {
            RemoveSiteCollectionAdmin -TenantObject $tenant -TenantConnection $tenantConnection -Url $siteUrl -User $tenantAdmin
        }

# disconnect from web

        Disconnect-PnPOnline -Connection $webConnection

        Write-Host "Disconnected from web"
    }
    else
    {
        Write-Host -ForegroundColor Magenta "Element from workload " + $csvItem.Workload + " skipped."
    }

    Write-Host
}

Write-Host -ForegroundColor Green "Done."

Stop-Transcript
