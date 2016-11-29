#Import the required DLL
Import-Module 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll'
Import-Module 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'
#OR
Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll'

$YOURTENANT = ''
$DIRECTORYSITE = ''
$DIRECTORYNAME = ''
$DIRECTORYDESCRIPTION = ''
$SITECOLLECTIONALIAS = 'Site'
$SUBSITEALIAS = 'SubSite'

$Username = "admin@$YOURTENANT.onmicrosoft.com"

$AdminUrl = "https://$YOURTENANT-admin.sharepoint.com/"
$DestinationSiteURL = "https://$YOURTENANT.sharepoint.com/sites/$DIRECTORYSITE/"
$DestinationListName = $DIRECTORYNAME
$Password = Read-Host -Prompt 'Please enter your password' -AsSecureString
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($DestinationSiteURL)
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $Username, $Password
$SPOCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username,$Password)
$Context.Credentials = $SPOCredentials

#Delete Existing List
function Delete-List() {
    [CmdletBinding()]
    param()
    $list = $context.Web.Lists.GetByTitle($DestinationListName)
    $context.Load($list)
    $list.DeleteObject()
    $list.Update()

    #Expected exception: http://social.technet.microsoft.com/wiki/contents/articles/29531.csom-sharepoint-online-delete-list-using-powershell.aspx
    $Context.ExecuteQuery()
}

#Create List
function New-List() {
    $ListInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
    $ListInfo.Title = $DestinationListName
    $ListInfo.TemplateType = '100'
    $List = $Context.Web.Lists.Add($ListInfo)
    $List.Description = $DIRECTORYDESCRIPTION
    $List.Update()
    $Context.ExecuteQuery()

    $List.Fields.AddFieldAsXml("<Field Type='URL' DisplayName='$SUBSITEALIAS' Format='Hyperlink' Name='$SUBSITEALIAS'/>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
    $clientfield = $list.Fields.GetByInternalNameOrTitle('Title')
    $clientfield.Title = $SITECOLLECTIONALIAS;
    $clientfield.Update()
    $List.Update()
    $Context.ExecuteQuery()
}


function Get-SPOWebs() {
    [CmdLetBinding()]
    param(
        $Url = $(throw 'Please provide a Site Collection Url'),
        $Credential = $(throw 'Please provide a Credentials')
    )
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)  
    $context.Credentials = $Credential 
    $web = $context.Web
    $context.Load($web)
    $context.Load($web.Webs)
    $context.ExecuteQuery()
    foreach($web in $web.Webs)
    {
        Get-SPOWebs -Url $web.Url -Credential $Credential 
        $web
    }
}

function Export-List() {
    param(
    [Parameter(valueFromPipeLine=$true)]
       $ExportList = $(throw 'Please enter a list of sites')
    )
    
    #Write-Host 'Exporting List'
    $List = $Context.Web.Lists.GetByTitle($DestinationListName)
    $Context.Load($List)
    $Context.ExecuteQuery()

    foreach ($Entry in $ExportList) {
        $ListItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
        $Item = $List.AddItem($ListItemInfo)
        $Item["Title"] = $Entry.$SITECOLLECTIONALIAS
        $Item["Project"] = $Entry.Url+ ", " + $Entry.$SUBSITEALIAS
        $Item.Update()
        $Context.ExecuteQuery()
    }
    $ExportList | Export-Csv -Path C:\sitelist.csv -NoTypeInformation
}

#Filter List
function Filter-List() {
    param(
       [Parameter(valueFromPipeLine=$true)]
       $UnFiltered = $(throw 'Please enter a list of sites')
    )
    $FilteredList = $UnFiltered | Where-Object {$_.$SITECOLLECTIONALIAS -ne ''} 
    Export-List($FilteredList)
}

function publish-sharepointdirectory() {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [Switch]$ListSites
    )
    try {
        #Retrieve all site collection infos
        Connect-SPOService -Url $AdminUrl -Credential $Credentials
    }
    catch {
        Write-Error "Password not valid" -ErrorAction Stop
    }
        $sites = Get-SPOSite 

        $SiteCollection=@()
        $SitesNotAdded=@()
        $IdField = 0

        #Retrieve and print all sites
        foreach ($site in $sites)
        {
            $AllWebs = Get-SPOWebs -Url $site.Url -Credential $SPOCredentials -ErrorAction SilentlyContinue -ErrorVariable GetSiteError
            if ($GetSiteError) {
                $SitesNotAdded += $site.Title+' | ' 
                if($ListSites) {Write-Host $site.Title -ForegroundColor Red}
            } else {
                Write-Host $site.Title -ForegroundColor Green
                $AllWebs | ForEach-Object {
                    $IdField = $IdField + 1
                    $SubSiteItem = New-Object PSObject
                    Add-Member -InputObject $SubSiteItem -MemberType NoteProperty -Name Id -Value $IdField
                    Add-Member -InputObject $SubSiteItem -MemberType NoteProperty -Name $SITECOLLECTIONALIAS -Value $site.Title
                    Add-Member -InputObject $SubSiteItem -MemberType NoteProperty -Name $SUBSITEALIAS -Value $_.Title
                    Add-Member -InputObject $SubSiteItem -MemberType NoteProperty -Name Url -Value $_.Url
                    $SiteCollection += $SubSiteItem
                    Write-Host " "$_.Title
                } 
            }
        }
        Delete-List -ErrorAction SilentlyContinue
        New-List
        Filter-List($SiteCollection)
        Write-Host 'Sites not added:'$SitesNotAdded -ForegroundColor Red
}

Export-ModuleMember -Function publish-sharepointdirectory