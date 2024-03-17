param
(
    [parameter(Mandatory=$false)]
    [Alias("f")]
    [string]$Function,
    [parameter(Mandatory=$false)]
    [Alias("u")]
    [string]$User,
    [parameter(Mandatory=$false)]
    [Alias("i")]
    [string]$IP,
    [parameter(Mandatory=$false)]
    [Alias("o")]
    [string]$OutFolder,
    [parameter(Mandatory=$false)]
    [Alias("s")]
    [string]$StartDate,
    [parameter(Mandatory=$false)]
    [Alias("e")]
    [string]$EndDate,
    [parameter(Mandatory=$false)]
    [Alias("a")]
    [string]$Analyst,
    [parameter(Mandatory=$false)]
    [Alias("h")]
    [switch]$Help

)

#Set the Company Domain Here
$Domain = ''

#
function identify {
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        Import-Module ExchangeOnlineManagement
    } 
    else {
        Install-Module -Name ExchangeOnlineManagement
        Import-Module ExchangeOnlineManagement
    }

    #Verify or set Analyst and then connect using that account
    if(!$Analyst){$Analyst = $env:UserName + '@' +$Domain}
    Write-Verbose ("[i] Connecting to Echange Online as " + $Analyst)
    Connect-ExchangeOnline -UserPrincipalName $Analyst -ShowBanner:$false

    #Verify date or Set It if not set
    if(!$EndDate){$EndDate = (Get-Date).ToShortDateString();}
    if(!$StartDate){$StartDate = ([DateTime]$EndDate).adddays(-30).ToShortDateString();}

    Write-Verbose ("[i] Pulling data for " + $User)
    $Records = Search-UnifiedAuditLog -UserIds $User -StartDate $StartDate -EndDate $EndDate -Operations MailItemsAccessed -ResultSize 5000

    $ClientIPs = @();
    ForEach($Record in $Records)
    {
        $AuditData = ConvertFrom-Json $Record.Auditdata
        $ClientIPs += $AuditData.ClientIPAddress
    }

    $ClientIPs = $ClientIPs | Sort-Object | Select-Object -Unique
    Write-Host "[i] Found" $ClientIPs.Count "IPs:"
    $ClientIPs
}

function review {
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        Import-Module ExchangeOnlineManagement
    } 
    else {
        Install-Module -Name ExchangeOnlineManagement
        Import-Module ExchangeOnlineManagement
    }

    #Verify or set Analyst and then connect using that account
    if(!$Analyst){$Analyst = $env:UserName + '@' +$Domain}
    Write-Verbose ("[i] Connecting to Echange Online as " + $Analyst)
    Connect-ExchangeOnline -UserPrincipalName $Analyst -ShowBanner:$false

    #Verify date or Set It if not set
    if(!$EndDate){$EndDate = (Get-Date).ToShortDateString();}
    if(!$StartDate){$StartDate = ([DateTime]$EndDate).adddays(-30).ToShortDateString();}

    Write-Verbose ("[i] Pulling data for " + $User)
    $Records = Search-UnifiedAuditLog -UserIds $User -StartDate $StartDate -EndDate $EndDate -Operations MailItemsAccessed -ResultSize 5000 | Where {$_.AuditData -like '*"MailAccessType","Value":"Bind"*'}
    $MessageIDs = @();
    $IPs = $IP.Split(",")
    ForEach($Record in $Records)
    {
        $AuditData = ConvertFrom-Json $Record.Auditdata
        $FolderItems = $AuditData.Folders.FolderItems
        if ($IPs -contains $AuditData.ClientIPAddress)
            {
            ForEach($Item in $FolderItems)
            {
                $MessageIDs += $Item.InternetMessageId
            }
        }
        
    }
    $MessageIDs = $MessageIDs | Sort-Object | Select-Object -Unique
    Write-Host "[i] Found" $MessageIDs.Count "Emails accessed by:" $IP
}

function export {
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        Import-Module ExchangeOnlineManagement
    } 
    else {
        Install-Module -Name ExchangeOnlineManagement
        Import-Module ExchangeOnlineManagement
    }

    #Verify or set Analyst and then connect using that account
    if(!$Analyst){$Analyst = $env:UserName + '@' +$Domain}
    Write-Verbose ("[i] Connecting to Echange Online as " + $Analyst)
    Connect-ExchangeOnline -UserPrincipalName $Analyst -ShowBanner:$false

    #Verify date or Set It if not set
    if(!$EndDate){$EndDate = (Get-Date).ToShortDateString();}
    if(!$StartDate){$StartDate = ([DateTime]$EndDate).adddays(-30).ToShortDateString();}

    if(!$OutFolder) 
    {
        $UserFolder = "\" + $User + "\"
        $OutFolder = Join-Path $PSScriptRoot $UserFolder
    }
    if(-not (test-path $OutFolder))
    {
        New-Item -Path $OutFolder -ItemType Directory | Out-Null
    }
    
    Write-Verbose ("[i] Pulling data for " + $User)
    $Records = Search-UnifiedAuditLog -UserIds $User -StartDate $StartDate -EndDate $EndDate -Operations MailItemsAccessed -ResultSize 5000 | Where {$_.AuditData -like '*"MailAccessType","Value":"Bind"*'}
    $IPs = $IP.Split(",")
    ForEach($Record in $Records)
    {
        $AuditData = ConvertFrom-Json $Record.Auditdata
        $FolderItems = $AuditData.Folders.FolderItems
        if ($IPs -contains $AuditData.ClientIPAddress)
            {
            ForEach($Item in $FolderItems)
            {
                $ItemFile = $Item.InternetMessageId.trim("<").trim(">") + ".txt"
                $SaveLocation = $OutFolder + $ItemFile
                Get-MessageTrace -StartDate (Get-Date).adddays(-10).ToShortDateString() -EndDate (Get-Date).ToShortDateString() -MessageID $Item.InternetMessageId | fl * | Out-File -FilePath $SaveLocation
            }
        }
        
    }
}

$HelpMessage=@'
This is the Help Message. Hopefully you don't need help yet cause this is still a WIP.
'@

function main {
    Write-Verbose ("=========================")
    Write-Verbose ("Function:       " + $Function)
    Write-Verbose ("User:           " + $User)
    Write-Verbose ("IP:             " + $IP)
    Write-Verbose ("OutFolder:      " + $OutFolder)
    Write-Verbose ("Analyst:        " + $Analyst)
    Write-Verbose ("Start Date:     " + $StartDate)
    Write-Verbose ("End Date:       " + $EndDate)
    Write-Verbose ("=========================")

    if($Help)
    {
        $HelpMessage
        exit
    }

    if($Parameters.Count -eq 0)
    {
        $HelpMessage
    }else {
       switch ($Function)
        {
            'identify'{identify}
            'review'{review}
            'export'{export}
            Default {$HelpMessage}
        }     
    }
}

$Parameters = $PSBoundParameters
main