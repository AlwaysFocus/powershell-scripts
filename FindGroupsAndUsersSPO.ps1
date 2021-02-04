# Data collection variable
$ADGroupCollection = @()
Function Save-File ([string]$initialDirectory) {

	$SaveInitialPath = "C:\"
	$SaveFileName = "Result.csv"

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $SaveInitialPath
	$OpenFileDialog.FileName = $SaveFileName
    $OpenFileDialog.ShowDialog() | Out-Null

    return $OpenFileDialog.filename

}

$t = @"
           _                         ______                   
     /\   | |                       |  ____|                  
    /  \  | |_      ____ _ _   _ ___| |__ ___   ___ _   _ ___ 
   / /\ \ | \ \ /\ / / _` | | | / __|  __/ _ \ / __| | | / __|
  / ____ \| |\ V  V / (_| | |_| \__ \ | | (_) | (__| |_| \__ \
 /_/    \_\_| \_/\_/ \__,_|\__, |___/_|  \___/ \___|\__,_|___/
                            __/ |                             
                           |___/                                                              
"@

for ($i=0;$i -lt $t.length;$i++) {
if ($i%2) {
 $c = "green"
}
elseif ($i%5) {
 $c = "blue"
}
elseif ($i%7) {
 $c = "red"
}
else {
   $c = "white"
}
write-host $t[$i] -NoNewline -ForegroundColor $c
}
# Install Sharepoint powershell module
Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser

$CompanySite = Read-Host -Prompt 'Enter the name of your sharepoint online site (companyname-admin.sharepoint.com)'

$SearchType = Read-Host -Prompt 'Search all (a) or search for specific user (u)? '
# Connect to Sharepoint
Connect-SPOService -Url $CompanySite




if ( 'a' -eq $SearchType ) {
    # Iterate through all sites in the site collection
    foreach ($Site in Get-SPOSite) { 
        Write-host -foregroundcolor blue "Processing Site: "$Site.Url

        $ADGroups = Get-SPOUser -Site $Site.Url -Limit ALL

        # Iterate through each AD group
        foreach ($Group in $ADGroups) {
            # Get all direct permissions
            $Permissions = $Group.UserType
        
            # Get user groups that contain AD groups as users
            $SiteGroups = $Group.Groups

            # Add all data to object array
            $ADGroup = New-Object psobject
            $ADGroup | add-member -type noteproperty -name "Sharepoint Site URL" -value $Site.Url
            $ADGroup | add-member -type noteproperty -name "Group/User Name" -value $Group.DisplayName
            $ADGroup | add-member -type noteproperty -name "Permission Set" -value ($Permissions -join ",")
            $ADGroup | add-member -type noteproperty -name "SharePoint Groups" -value ($SiteGroups -join ",")
            # Append object to Array
            $ADGroupCollection += $ADGroup
        }
    }
    $ReportResultsPath = Save-File
    # Write all the collected data to a CSV file in the selected location
    $ADGroupCollection | export-csv -path $ReportResultsPath -notypeinformation -Force
}
elseif ('u' -eq $SearchType) {

    $User = Read-Host -Prompt 'Enter the name of the user (eg: user@your-domain.com): '

    # Iterate through all sites in the site collection
    foreach ($Site in Get-SPOSite) { 
        Write-host -foregroundcolor blue "Processing Site: "$Site.Url

        $ADGroups = Get-SPOUser -Site $Site.Url -Limit ALL | Where-Object { $_.LoginName -eq $User }

        # Iterate through each AD group
        foreach ($Group in $ADGroups) {
            # Get all direct permissions
            $Permissions = $Group.UserType
        
            # Get user groups that contain specified user as a user
            $SiteGroups = $Group.Groups

            # Add all data to object array
            $ADGroup = New-Object psobject
            # $ADGroup | add-member -type noteproperty -name "Site Collection" -value $Site.RootWeb.Title
            $ADGroup | add-member -type noteproperty -name "URL" -value $Site.Url
            $ADGroup | add-member -type noteproperty -name "Group Name" -value $Group.DisplayName
            $ADGroup | add-member -type noteproperty -name "Direct Permissions" -value ($Permissions -join ",")
            $ADGroup | add-member -type noteproperty -name "SharePoint Groups" -value ($SiteGroups -join ",")
            # Append object to Array
            $ADGroupCollection += $ADGroup
        }
    }
    
    $ReportResultsPath = Save-File
    # Write all the collected data to a CSV file in the selected location
    $ADGroupCollection | export-csv -path $ReportResultsPath -notypeinformation -Force

}



    

