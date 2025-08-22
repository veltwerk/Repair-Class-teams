# ============================================
# Bulk Repair Script for EDU Teams Notebooks V1.0
# ============================================

Write-Host "This script repairs the Notebook Classroom in Educational Teams groups in bulk." -ForegroundColor Yellow
$TeamsAdmin = Read-Host "Enter the UPN of a licensed Teams Admin"
# -------------------------------
# Functie: Module installeren en importeren
# -------------------------------
function Ensure-GraphModule {
    param (
        [string]$ModuleName
    )
    if (-not(Get-InstalledModule $ModuleName -ErrorAction SilentlyContinue)) {
        Write-Host "Installing module $ModuleName..." -ForegroundColor Yellow
        Install-Module $ModuleName -Confirm:$false -Force -Scope CurrentUser -AllowClobber
    }
    if (-not(Get-Module $ModuleName -ErrorAction SilentlyContinue)) {
        Write-Host "Importing module $ModuleName..." -ForegroundColor Yellow
        Import-Module $ModuleName
    }
    Write-Host "$ModuleName is ready to use." -ForegroundColor Green
}

# -------------------------------
# Modules controleren
# -------------------------------
$modules = @(
    "Microsoft.Graph.Groups",
    "Microsoft.Graph.Users"
)

foreach ($module in $modules) {
    Ensure-GraphModule -ModuleName $module
}

# -------------------------------
# Verbinding maken met Graph
# -------------------------------
Write-Host "Connecting to tenant..." -ForegroundColor Yellow

#Disconnect any existing sessions
Disconnect-MgGraph -ErrorAction SilentlyContinue > $null
Remove-Item "$env:USERPROFILE\.mgcontext" -Force -ErrorAction SilentlyContinue

# Connect to Microsoft Graph with required scopes
Connect-MgGraph -Scopes "Sites.FullControl.All","Group.ReadWrite.All","User.Read.All"
Write-Host "Connection successful." -ForegroundColor Green

# -------------------------------
# Teams Admin ophalen
# -------------------------------

$user = Get-MgUser -UserId $TeamsAdmin
$ownerId = $user.Id


# -------------------------------
# Teams ophalen en filteren
# -------------------------------
$AllTeams = Get-MgGroup -All | Where-Object {
    $_.Team -ne $null -and
    $_.Visibility -eq "HiddenMembership"
} | Select-Object -Property DisplayName, MailNickname, Id

Write-host ("{0} teams found." -f $AllTeams.Count)

$allDone = $false
while (!$allDone) {
    Write-Host "Enter a search term to filter the mailnickname of the teams; for example '2526-' or 'class'." -ForegroundColor Yellow
    Write-Host "Press <ENTER> to exit the script." -ForegroundColor Yellow
    $queryGroups = Read-Host "Search term: "
    if ($queryGroups -eq "") {
        Write-host "Script is ending."
        $allDone = $true
    } else {
        $queriedGroups = $AllTeams | Where-Object { $_.MailNickname -like "*$queryGroups*" }
        if ($queriedGroups.Count -eq 0) {
            Write-Host "No Teams found for the given query." -ForegroundColor Red
        } else {
            Write-Host "Select the Teams to repair via Out-GridView." -ForegroundColor Yellow

            $selectedGroups = $queriedGroups | Out-GridView -OutputMode Multiple
            $groupCount = $selectedGroups.Count

            # -------------------------------
            # Temporarily add owner
            # -------------------------------
            foreach ($team in $selectedGroups) {
                $progress = ($selectedGroups.IndexOf($team) / $groupCount) * 100
                Write-Progress -Activity "Adding Teams Owner" -Status "Adding to $($team.DisplayName)" -PercentComplete $progress

                $newGroupOwner = @{
                    "@odata.id" = "https://graph.microsoft.com/v1.0/users/$ownerId"
                }


                New-MgGroupOwnerByRef -GroupId $team.Id -BodyParameter $newGroupOwner -ErrorAction SilentlyContinue
            
            }

            # -------------------------------
            # Set EDU App permissions
            # -------------------------------
            $eduApps = @(
                @{ id = "22d27567-b3f0-4dc2-9ec2-46ed368ba538"; name = "EDU Teams Assignment" },
                @{ id = "c9a559d2-7aab-4f13-a6ed-e7e9c52aec87"; name = "EDU Teams Notebook" },
                @{ id = "13291f5a-59ac-4c59-b0fa-d1632e8f3292"; name = "EDU OneNote" },
                @{ id = "2d4d3d8e-2be3-4bef-9f87-7875a61c29de"; name = "EDU Teams Files" },
                @{ id = "8f348934-64be-4bb2-bc16-c54c96789f43"; name = "EDU Teams Calendar" }
            )

            foreach ($team in $selectedGroups) {
                $progress = ($selectedGroups.IndexOf($team) / $groupCount) * 100
                Write-Progress -Activity "Setting permissions" -Status "For $($team.DisplayName)" -PercentComplete $progress

                $site = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($team.Id)/sites/root"

                foreach ($app in $eduApps) {
                    $body = @{
                        roles               = @("fullcontrol")
                        grantedToIdentities = @(@{
                                application = @{
                                    id          = $app.id
                                    displayName = $app.name
                                }
                            })
                    }

                    Write-Host "Setting permissions for $($app.name) on $($team.DisplayName)" -ForegroundColor Cyan
                    $Result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/permissions" -Body $body
                }

                # Remove temporary owner
                Write-Host "Removing $($user.DisplayName) as owner of $($team.DisplayName)"
                Remove-MgGroupOwnerByRef -GroupId $team.Id -DirectoryObjectId $ownerId -ErrorAction SilentlyContinue
            }

            # -------------------------------
            # Export to CSV
            # -------------------------------
            $selectedGroups | Select-Object DisplayName, Id, MailNickname | Export-Csv -Path "$HOME\Repaired-Teams.csv" -NoTypeInformation
            Invoke-Item -Path "$HOME\Repaired-Teams.csv"
            $alldone = $true
        }
    }
}
Disconnect-MgGraph
