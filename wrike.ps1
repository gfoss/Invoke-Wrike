#requires -version 3

  #====================================#
  # Wrike - PowerShell API Integration #
  # LogRhythm Security Operations      #
  # greg . foss @ logrhythm . com      #
  # v0.1  --  June, 2017               #
  #====================================#

# Copyright 2017 LogRhythm Inc.   
# Licensed under the MIT License. See LICENSE file in the project root for full license information.

function wrike {

<#

.SYNOPSIS

    Open Wrike in Application Mode in Chrome

    Create new tasks

    Integrate with automation tools and dynamically generate tasks

.USAGE

    Ensure that you have the Wrike functions imported

        PS C:\> Import-Module .\wrike.ps1

    Open Wrike in Chrome's Application Mode

        PS C:\> wrike

#>

    $wrike = "https://www.wrike.com/workspace.htm"
    Start-Process 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe' --app=$wrike
}

function Invoke-Wrike {

<#

.SYNOPSIS

    Query the Wrike API and Create Tasks

.INSTALL

    Obtain an API token from the Wrike and hardcode / append during function calls:

        https://www.wrike.com/frontend/apps/index.html#/api

.USAGE

    Ensure that you have the Wrike functions imported

        PS C:\> Import-Module .\wrike.ps1
    
    Interact with the Wrike API
        
        PS C:\> invoke-wrike -connectionTest -accessToken <api key>

        PS C:\> invoke-wrike -adminSearch -accessToken <api key>

        PS C:\> invoke-wrike -newTask <"Task Title"> -user <"assigned user (first name)"> -folder <"where to add that task?"> -accessToken <api key>

    Hardcode variables for quick access from the command line
        
        PS C:\> invoke-wrike -newTask <"Task Title">

#>

    [CmdLetBinding()]
    param( 
        [string]$accessToken,
        [string]$folder,
        [string]$user,
        [switch]$connectionTest,
        [string]$newTask,
        [string]$description = "API - Generated Task",
        [switch]$adminSearch
    )

    if ( -Not $accessToken ) {
        Write-Host "Wrike API Key Required..."
        Write-Host ""
        Exit 1
    }
    if ( -Not $folder ) {
        Write-Host "Wrike Folder Name Required..."
        Write-Host ""
        Exit 1
    }
    if ( -Not $user ) {
        Write-Host "Wrike User Required..."
        Write-Host ""
        Exit 1
    }

    $today = "{0:MM-dd-yyyy}" -f (Get-Date).ToUniversalTime()
    $twoDays = "{0:MM-dd-yyyy}" -f ((Get-Date).ToUniversalTime()).AddDays(2)

    # Test credentials and gather high-level details about your Wrike account
    if ( $connectionTest ) {

        $accountRequestData = Invoke-RestMethod -uri https://www.wrike.com/api/v3/accounts -Headers @{'Authorization' = ' bearer '+$accessToken}
        $result = $accountRequestData.data
        $result
    
    }

    # Create and assign a new task
    if ( $newTask ) {
        
        # Folder Query
        $folderQuery = Invoke-RestMethod -uri https://www.wrike.com/api/v3/folders -Headers @{'Authorization' = ' bearer '+$accessToken}
        $folderData = $folderQuery.data | Select-Object Id,title
        $wrikeID = @($folderData | findstr -i "$folder").split(" ")[0]

        # User Query
        $userQuery = Invoke-RestMethod -uri https://www.wrike.com/api/v3/contacts -Headers @{'Authorization' = ' bearer '+$accessToken}
        $userData = $userQuery.data | Select-Object id,firstName,lastName
        $ownerID = @($userData | findstr -i "$user").split(" ")[0]
        $ownerName = @($userData | findstr -i "$user").split(" ")[1]

        # Create Basic Task
        $payload = "title=$newTask&description=$description&responsibles=[`"$ownerID`"]" #&dates=@{'start':$today, 'due':$twoDays}"
        $output = Invoke-RestMethod -uri https://www.wrike.com/api/v3/folders/$wrikeID/tasks -Headers @{'Authorization' = ' bearer '+$accessToken} -Method Post -body $payload

        Write-Host ""
        Write-Host "Task: `"$newTask`" created, and assigned to $ownerName..."
        Write-Host ""

    }

    # Search for all Wrike admins within your organization
    if ( $adminSearch ) {

        $adminSearchData = Invoke-RestMethod -uri https://www.wrike.com/api/v3/contacts -Headers @{'Authorization' = ' bearer '+$accessToken}
        $result = $adminSearchData.data | Select-Object profiles | findstr -i "admin=true"
        Write-Host ""
        $result
        Write-Host ""

    }

# Clear assigned variables
Get-Variable | Remove-Variable -EA 0

}
