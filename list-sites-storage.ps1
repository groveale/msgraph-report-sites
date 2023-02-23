##############################################
# Functions
##############################################



# Connect to MSGraph
function ConnectToMSGraph ($parameters) {
    Connect-MgGraph -ClientId $parameters.spStorageMetricsAppId.Value -TenantId $parameters.tenantId.Value -CertificateThumbprint $parameters.thumbprint.Value
}

# Get all Sites, check for PH, if PH get details
function GetPHDriveData {

    $output = @()
    $sites = Get-MgSite -Property "id, webUrl" -All

    foreach($site in $sites)
    {
        if ($site.WebUrl.Contains("-my.sharepoint.com"))
        {
            continue
        }

        $siteSize = 0
        $phSize = $false

        [array]$drives = Get-MgSiteDrive -SiteId $site.id -All

        if ($drives.Length -gt 0)
        {
            # # Has the potential to pull back multiple matches - ToDo investigate
            # $phDrive = $drives | where { $_.Name.Contains("Preservation Hold Library") }

            # if ($phDrive)
            # {
            #     # Get the root of the PH Drive (PH Library should not contain folders)
            #     $driveRoot = Get-MgDriveRoot -DriveId $phDrive.Id
            #     $phSize = $driveRoot.Size # Bytes
            #     $phSizeMB  = $phSize / 1024 / 1024
            # }

            # Get total site size
            # Use first item in drive list - seems to be a bug (or maybe a feature) each drive quota is the same across the site
            if ($drives)
            {
                try {
                    $drive = Get-MgDrive -DriveId $drives[0].Id
                    $siteSize = $drive.Quota.Used # Bytes
                    $siteSizeMB = $siteSize / 1024 / 1024
                }
                catch {
                    Write-Host "Error - $($_)"
                }
                
            }


            $item = New-Object PSObject
            $item | Add-Member -type NoteProperty -Name 'Url' -Value $site.webUrl
            $item | Add-Member -type NoteProperty -Name 'SiteSizeMB' -Value ([math]::Round($siteSizeMB,2))

            # if($pHSize)
            # {
            #     $item | Add-Member -type NoteProperty -Name 'PreservationSizeMB' -Value ([math]::Round($phSizeMB,2))
            #     $item | Add-Member -type NoteProperty -Name 'PreservationPercentage' -Value ([math]::Round((($phSizeMB/$siteSizeMB)*100)))
            # } else {
            #     $item | Add-Member -type NoteProperty -Name 'PreservationSizeMB' -Value "N/A"
            #     $item | Add-Member -type NoteProperty -Name 'PreservationPercentage' -Value "N/A"  
            # }

            Write-Host $item
        
            $output +=$item
        }
        else{

            # When sites don't have any drives (All teams sites should)
            # Some legacy calssic templates don't e.g. communities

            $item = New-Object PSObject
            $item | Add-Member -type NoteProperty -Name 'Url' -Value $site.webUrl
            $item | Add-Member -type NoteProperty -Name 'SiteSizeMB' -Value "Error - no files"
            # $item | Add-Member -type NoteProperty -Name 'PreservationSizeMB' -Value "N/A"
            # $item | Add-Member -type NoteProperty -Name 'PreservationPercentage' -Value "N/A" 

            $output +=$item
        }

    }
        

    return $output
}

##############################################
# Main
##############################################


# Load Parameters from json file
$parametersListContent = Get-Content '.\parameters.json' -ErrorAction Stop

$parameters = $parametersListContent | ConvertFrom-Json

## Connect
ConnectToMSGraph $parameters

$metrics = GetPHDriveData

$metrics | sort -Property "SiteSizeMB" -Descending | Export-Csv -Path $parameters.outputDir.Value -NoTypeInformation