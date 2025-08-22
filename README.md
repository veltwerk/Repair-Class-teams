# Repair-Class-teams

Fixes Class Teams that were created using the deprecated teams templates. Ref: [Classwork, Assignment Import/upload attachment file issue.](https://learn.microsoft.com/en-us/answers/questions/5508367/classwork-assignment-import-upload-attachment-file)

## disclaimer

Please be carefull when using this script. It add permissions to the team's Sharepoint site for several educational apps. This script was created and tested in a very short time, without much error handling so the regular caveats apply: use it at your own risk.

## What the script is for

Many educational organizations use (older) PowerShell modules to generate class teams in bulk each school year. Sometime before this summer, Microsoft changed the process of provisioning those teams. The previously recommended way of creating an M365 group and promoting it to a team using the Edu_Class template no longer provisions all the necessary apps correctly.

Veldwerk, like many other parties that use this method of creating class teams, missed the memo and were taken by surpise. The [Veldwerk tools for provisioning Azure / M365 for schools](https://www.veldwerk.nl/oplossingen/software-oplossingen/educonnector/) have been updated to make sure new class teams are created with all edu functions operational.

The powershell script uses the graph API to help you select the Teams that need fixing and grants the educational teams Apps the permissions needed to operate.

## How to use it

- Download the script and run it. If your administrator put a limiting powershell profile or execution policy in place, bypass those by using ``powershell.exe -Executionpolicy bypass -NoProfile``
- The script will ask for the userPrincipalname of an admin; needed to approve Graph API permissions for the script
- It will check the installation of the MgGraph modules: Microsoft.Graph.Groups , Microsoft.Graph.Users
- All M365 groups that have a Team and that have ``Visibility = HiddenMembership`` to select class Teams
- It will ask you for a search term that will be used to filter on the ``mailNickname`` field
- A gridView is used to select the filtered teams you want fixed
- The admin account will be made owner of each team 
- the following apps will get ``Full Control`` permissions on the teams Sharepoint site


| Name                  | App Id                                |
| --------------------- | ------------------------------------- |
| EDU Teams Assignment  | 22d27567-b3f0-4dc2-9ec2-46ed368ba538  |
| EDU Teams Notebook    | c9a559d2-7aab-4f13-a6ed-e7e9c52aec87  |
| EDU OneNote           | 13291f5a-59ac-4c59-b0fa-d1632e8f3292  |
| EDU Teams Files       | 2d4d3d8e-2be3-4bef-9f87-7875a61c29de  |
| EDU Teams Calendar    | 8f348934-64be-4bb2-bc16-c54c96789f43  |

- The owner is removed after the team is repaired

When finished a CSV with the repaired teams is written and opened.
