
        Wrike PowerShell API Integration
        greg . foss [at] logrhythm . com
        v0.1 -- July, 2017

## [About]

Automate all things Wrike! This script was designed to integrate with the LogRhythm SIEM, and automate the assignment of tasks based on SIEM alarms.

Blog Post => https://logrhythm.com/blog/automate-project-management-with-smartresponse/

![Overview](/images/wrike.png)

## [Install]

    Obtain an API token from the Wrike and hardcode / append during function calls:

        https://www.wrike.com/frontend/apps/index.html#/api

## [Synopsis]

    Open Wrike in Application Mode in Chrome

    Create new tasks

    Integrate with automation tools and dynamically generate tasks

## [Usage]

    Ensure that you have the Wrike functions imported

        PS C:\> Import-Module .\wrike.ps1

    Open Wrike in Chrome's Application Mode

        PS C:\> wrike
    
    Interact with the Wrike API
        
        PS C:\> invoke-wrike -connectionTest -accessToken <api key>

        PS C:\> invoke-wrike -adminSearch -accessToken <api key>

        PS C:\> invoke-wrike -newTask <"Task Title"> -user <"assigned user (first name)"> -folder <"where to add that task?"> -accessToken <api key>

    Hardcode variables for quick access from the command line
        
        PS C:\> invoke-wrike -newTask <"Task Title">

## [License]

Copyright 2017 LogRhythm Inc.   
Licensed under the MIT License. See LICENSE file in the project root for full license information.