
  #====================================#
  # Wrike - PowerShell API Integration #
  # LogRhythm Security Operations      #
  # greg . foss @ logrhythm . com      #
  # v0.1  --  April, 2017              #
  #====================================#

# Copyright 2017 LogRhythm Inc.   
# Licensed under the MIT License. See LICENSE file in the project root for full license information.

<#

USAGE:

    Configure as a LogRhythm SmartResponse - works as contextual / standard SmartResponse

#>

[CmdLetBinding()]
param( 
    [string]$accessToken,
    [string]$folder,
    [string]$user,
    [string]$newTask,
    [string]$description = "API - Generated Task"
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

}
Exit 0