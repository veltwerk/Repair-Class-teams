# Repair-Class-teams

Fixes Class Teams that were created using the deprecated teams teamplates.

## disclaimer

Please be carefull when using this script. It add permissions to the team's Sharepoint site for several educational apps. This script was created and tested in a very short time, so the regular caveats apply: use it at your own risk.

## What the script is for

Many educational organizations use (older) PowerShell modules to generate class teams in bulk each school year. Sometime before this summer, Microsoft changed the process of provisioning those teams. The previously recommended way of creating an M365 group and promoting it to a team using the Edu_Class template no longer provisions all the necessary apps correctly.

Veldwerk, like many other parties that use this method of creating class teams, missed the memo and were taken by surpise. The [Veldwerk tools for provisioning Azure / M365 for schools](https://www.veldwerk.nl/oplossingen/software-oplossingen/educonnector/) have been updated to make sure new class teams are created with all edu functions operational.

The powershell script uses the graph API to help you select the Teams that need fixing and grants the educational teams Apps the permissions needed to operate.

## How to use it

- Download the script and run it. If your administrator put a limiting powershell profile or execution policy in place, bypass those by using ``powershell.exe -Executionpolicy bypass -NoProfile``
- The script will ask for the userPrincipalname of an admin (we tested only with global admins)
