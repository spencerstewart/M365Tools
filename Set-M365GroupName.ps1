<#
 .Synopsis
  Helps rename a M365 Group and all its artifacts

 .Description
  Changes the Microsoft 365 Group's DisplayName, Alias, PrimarySmtpEmail, and SharePoint site URL.
  
 .Parameter GroupIdentity
  Mandatory. Can be group display name or email.


 .Example
  # Rename everthing
  Set-M365GroupName -GroupIdentity "D365UsersJEGroup@contoso.com" -NewDisplayName "FIN-ALL-D365-GL User Group" -NewAlias "FIN-ALL-D365-GLUserGroup" `
                    -NewPrimarySmtpAddress "FIN-ALL-D365-GLUserGroup@contoso.com" -UpdateSharePointSiteUrl



#>

function Set-M365GroupName {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $GroupIdentity,
        [Parameter()]
        [string]
        $NewDisplayName,
        [Parameter()]
        [string]
        $NewAlias,
        [Parameter()]
        [string]
        $NewPrimarySmtpAddress,
        [Parameter()]
        [switch]
        $UpdateSharePointSiteUrl,
        [Parameter()]
        [string]
        $NewSharePointSiteUrl,
        [Parameter()]
        [string]
        $SharePointAdminSite = 'https://contoso-admin.sharepoint.com',
        [Parameter()]
        [switch]
        $Force
    )

    # If SharePoint Site is being renamed, need to make sure we can connect to SPO
    if (($NewSharePointSiteUrl -or $UpdateSharePointSiteUrl) -and !(Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable)) {
        Write-Host "Must have SharePoint Online Management Shell installed:" -ForegroundColor Red
        Write-Host 'https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online?view=sharepoint-ps' -ForegroundColor Red
        return
    } else {
        try {
            Get-SPOTenant -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Host "Authenticate to SharePoint Online Management Shell" -ForegroundColor  Yellow
            Connect-SPOService -Url $SharePointAdminSite
        }
    }

    Get-ExchangeOnlineSession

    try {
        $Group = Get-UnifiedGroup $GroupIdentity -ErrorAction Stop
        if ($Group.Count -gt 1) {
            Write-Host "The group identity $GroupIdentity resolves to more than one group: $($Group.Alias -join ', ')" -ForegroundColor Red
            throw "Specify the exact group using an email or more unique identifier."
        }
    } catch {
        throw $_.Exception.Message
    }
    
    if ($NewPrimarySmtpAddress) {
        Write-Host "[$($Group.Alias)] Changing PrimarySmtpAddress from '$($Group.PrimarySmtpAddress)' to '$NewPrimarySmtpAddress'"
        Set-UnifiedGroup -Identity $Group.Alias -PrimarySmtpAddress $NewPrimarySmtpAddress

        
        Write-Host "[$($Group.Alias)] Removing old PrimarySmtpAddress of '$($Group.PrimarySmtpAddress)'"
        Set-UnifiedGroup -Identity $Group.Alias -EmailAddresses @{remove="$($Group.PrimarySmtpAddress)"}
    }

    if ($NewDisplayName) {
        Write-Host "[$($Group.Alias)] Changing DisplayName from '$($Group.DisplayName)' to '$NewDisplayName'."
        Set-UnifiedGroup -Identity $Group.Alias -DisplayName $NewDisplayName
    }

    if ($NewAlias) {
        Write-Host "[$($Group.Alias)] Changing Alias from '$($Group.Alias)' to '$NewAlias'."
        Set-UnifiedGroup -Identity $Group.Alias -Alias $NewAlias
    }

    if ($UpdateSharePointSiteUrl) {

        # Generate update SharePoint Site URL if one not given
        if (!$NewSharePointSiteUrl) {
            $NewSharePointSiteUrl = $Group.SharePointSiteUrl -replace $Group.Alias,$NewAlias
            Write-Host "[$($Group.Alias)] Assuming new SharePoint Site URL should simply use the new alias: '$NewSharePointSiteUrl'." -ForegroundColor Cyan
        }

        Write-Host "[$($Group.Alias)] Changing SPOSite from '$($Group.SharePointSiteUrl)' to '$NewSharePointSiteUrl'."
        Start-SPOSiteRename -Identity $Group.SharePointSiteUrl -NewSiteUrl $NewSharePointSiteUrl -NewSiteTitle $NewDisplayName -Confirm:$Force   
    }
    


}
