# ─── Parameters ───────────────────────────────────────────────────────────────
param(
    [string] $TfsServer = "http://your-tfs-server:8080/tfs",
    [switch] $SkipSize
)

# ─── Load required TFS assemblies ─────────────────────────────────────────────
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.TeamFoundation.Client")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.TeamFoundation.Common")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.TeamFoundation.VersionControl.Client")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.TeamFoundation.Build.Client")

# ─── Helper: format bytes into human-readable size ────────────────────────────
function Format-Bytes {
    param([long]$Bytes)
    if     ($Bytes -ge 1GB) { "{0:N2} GB" -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { "{0:N2} MB" -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { "{0:N2} KB" -f ($Bytes / 1KB) }
    else                    { "$Bytes B" }
}

# ─── Connect to TFS Configuration Server ──────────────────────────────────────
Write-Host "[$(Get-Date)] Connecting to TFS: $TfsServer ..." -ForegroundColor Cyan
$tfsConfigServer = New-Object Microsoft.TeamFoundation.Client.TfsConfigurationServer(
    (New-Object Uri($TfsServer))
)
$tfsConfigServer.Authenticate()

# ─── Get all collections ───────────────────────────────────────────────────────
$collectionService = $tfsConfigServer.GetService([Microsoft.TeamFoundation.Framework.Client.ITeamProjectCollectionService])
$collections = $collectionService.GetCollections()

if ($SkipSize) {
    Write-Host "[$(Get-Date)] [-SkipSize] Project size calculation will be skipped.`n" -ForegroundColor Yellow
}

$results = @()

foreach ($collection in $collections) {
    try {
        if ($collection.State -ne "Started") {
            $results += [PSCustomObject]@{
                Collection                = $collection.Name
                CollState                 = $collection.State
                Project                   = "(Collection not started)"
                ProjectUri                = "-"
                ProjectGuid               = "-"
                ProjectState              = "-"
                Description               = "-"
                LastCheckin               = "-"
                LastUser                  = "-"
                LastUserDisplayName       = "-"
                LastUserMailAddress       = "-"
                ChangesetCount            = "-"
                ChangesetUsers            = "-"
                ChangesetUsersDisplayName = "-"
                ChangesetUsersMailAddress = "-"
                Members                   = "-"
                MembersDisplayName        = "-"
                MembersMailAddress        = "-"
                BuildDefinitionCount      = "-"
                LatestBuildDate           = "-"
                LatestVersionSize         = "-"
                BiggestFileSize           = "-"
                BiggestFile               = "-"
            }
            continue
        }

        Write-Host "[$(Get-Date)] Processing collection: $($collection.Name)" -ForegroundColor Cyan

        $collectionUri = New-Object Uri("$TfsServer/$($collection.Name)")
        $tfsCollection = [Microsoft.TeamFoundation.Client.TfsTeamProjectCollectionFactory]::GetTeamProjectCollection($collectionUri)
        $tfsCollection.Authenticate()

        $cssService      = $tfsCollection.GetService([Microsoft.TeamFoundation.Server.ICommonStructureService])
        $projects        = $cssService.ListAllProjects()
        $vcs             = $tfsCollection.GetService([Microsoft.TeamFoundation.VersionControl.Client.VersionControlServer])
        $securityService = $tfsCollection.GetService([Microsoft.TeamFoundation.Server.IGroupSecurityService])
        $buildService    = $tfsCollection.GetService([Microsoft.TeamFoundation.Build.Client.IBuildServer])

        if ($projects.Count -eq 0) {
            $results += [PSCustomObject]@{
                Collection                = $collection.Name
                CollState                 = $collection.State
                Project                   = "(No projects found)"
                ProjectUri                = "-"
                ProjectGuid               = "-"
                ProjectState              = "-"
                Description               = "-"
                LastCheckin               = "-"
                LastUser                  = "-"
                LastUserDisplayName       = "-"
                LastUserMailAddress       = "-"
                ChangesetCount            = "-"
                ChangesetUsers            = "-"
                ChangesetUsersDisplayName = "-"
                ChangesetUsersMailAddress = "-"
                Members                   = "-"
                MembersDisplayName        = "-"
                MembersMailAddress        = "-"
                BuildDefinitionCount      = "-"
                LatestBuildDate           = "-"
                LatestVersionSize         = "-"
                BiggestFileSize           = "-"
                BiggestFile               = "-"
            }
            continue
        }

        foreach ($project in $projects) {
            $projectPath               = "$/" + $project.Name
            $projectUri                = $project.Uri
            $projectGuid               = $project.Uri.Split('/')[-1]
            $projectDescription        = if ($project.Description) { $project.Description } else { "(none)" }
            $lastCheckin               = "-"
            $lastUser                  = "-"
            $lastUserDisplayName       = "-"
            $lastUserMailAddress       = "-"
            $changesetCount            = "-"
            $changesetUsers            = "-"
            $changesetUsersDisplayName = "-"
            $changesetUsersMailAddress = "-"
            $members                   = "-"
            $membersDisplayName        = "-"
            $membersMailAddress        = "-"
            $buildDefinitionCount      = "-"
            $latestBuildDate           = "-"
            $latestVersionSize         = "-"
            $biggestFileName           = "-"
            $biggestFileSize           = "-"

            Write-Host "[$(Get-Date)]  -> $($project.Name)" -ForegroundColor Gray

            # ── Last check-in date + user info ─────────────────────────────────
            try {
                $history = $vcs.QueryHistory(
                    $projectPath,
                    [Microsoft.TeamFoundation.VersionControl.Client.VersionSpec]::Latest,
                    0,
                    [Microsoft.TeamFoundation.VersionControl.Client.RecursionType]::Full,
                    $null, $null, $null,
                    1,
                    $false,
                    $false
                )
                $latestChangeset = $history | Select-Object -First 1
                if ($latestChangeset) {
                    $lastCheckin         = $latestChangeset.CreationDate.ToString("yyyy-MM-dd HH:mm")
                    $lastUser            = $latestChangeset.Committer
                    $lastUserDisplayName = $latestChangeset.CommitterDisplayName

                    # Resolve mail address via identity lookup
                    $committerIdentity = $securityService.ReadIdentity(
                        [Microsoft.TeamFoundation.Server.SearchFactor]::AccountName,
                        $latestChangeset.Committer,
                        [Microsoft.TeamFoundation.Server.QueryMembership]::None
                    )
                    if ($committerIdentity) {
                        $lastUserMailAddress = $committerIdentity.MailAddress
                    }
                }
            }
            catch {
                $lastCheckin         = "Error: $($_.Exception.Message)"
                $lastUser            = "Error"
                $lastUserDisplayName = "Error"
                $lastUserMailAddress = "Error"
            }

            # ── Changeset count + unique committers ────────────────────────────
            try {
                $allChangesets = $vcs.QueryHistory(
                    $projectPath,
                    [Microsoft.TeamFoundation.VersionControl.Client.VersionSpec]::Latest,
                    0,
                    [Microsoft.TeamFoundation.VersionControl.Client.RecursionType]::Full,
                    $null, $null, $null,
                    [int]::MaxValue,
                    $false,
                    $false
                )

                # Materialize enumerator so we can reuse the list
                $changesetList  = $allChangesets | ForEach-Object { $_ }
                $changesetCount = ($changesetList | Measure-Object).Count

                # Unique committers — account name and display name are on the changeset directly
                $uniqueCommitters = $changesetList | Sort-Object Committer -Unique

                $changesetUsers            = ($uniqueCommitters | ForEach-Object { $_.Committer            } | Sort-Object -Unique) -join ", "
                $changesetUsersDisplayName = ($uniqueCommitters | ForEach-Object { $_.CommitterDisplayName } | Sort-Object -Unique) -join ", "

                # Mail address requires an identity lookup per unique committer
                $mailAddresses = @()
                foreach ($committer in $uniqueCommitters) {
                    $identity = $securityService.ReadIdentity(
                        [Microsoft.TeamFoundation.Server.SearchFactor]::AccountName,
                        $committer.Committer,
                        [Microsoft.TeamFoundation.Server.QueryMembership]::None
                    )
                    if ($identity -and $identity.MailAddress) {
                        $mailAddresses += $identity.MailAddress
                    }
                }
                $changesetUsersMailAddress = if ($mailAddresses.Count -gt 0) { ($mailAddresses | Sort-Object -Unique) -join ", " } else { "(none)" }
            }
            catch {
                $changesetCount            = "Error: $($_.Exception.Message)"
                $changesetUsers            = "Error"
                $changesetUsersDisplayName = "Error"
                $changesetUsersMailAddress = "Error"
            }

            # ── Project members ────────────────────────────────────────────────
            try {
                $appGroups        = $securityService.ListApplicationGroups($project.Uri)
                $allAccountNames  = @()
                $allDisplayNames  = @()
                $allMailAddresses = @()

                foreach ($group in $appGroups) {
                    $groupIdentity = $securityService.ReadIdentity(
                        [Microsoft.TeamFoundation.Server.SearchFactor]::Sid,
                        $group.Sid,
                        [Microsoft.TeamFoundation.Server.QueryMembership]::Direct
                    )

                    foreach ($memberSid in $groupIdentity.Members) {
                        $member = $securityService.ReadIdentity(
                            [Microsoft.TeamFoundation.Server.SearchFactor]::Sid,
                            $memberSid,
                            [Microsoft.TeamFoundation.Server.QueryMembership]::None
                        )
                        if ($member) {
                            $prefix = if ($member.SecurityGroup) { "[G]" } else { "[U]" }
                            $allAccountNames  += "$prefix $($member.AccountName)"
                            $allDisplayNames  += "$prefix $($member.DisplayName)"
                            $allMailAddresses += "$prefix $($member.MailAddress)"
                        }
                    }
                }

                $members            = if ($allAccountNames.Count  -gt 0) { ($allAccountNames  | Sort-Object -Unique) -join ", " } else { "(none)" }
                $membersDisplayName = if ($allDisplayNames.Count  -gt 0) { ($allDisplayNames  | Sort-Object -Unique) -join ", " } else { "(none)" }
                $membersMailAddress = if ($allMailAddresses.Count -gt 0) { ($allMailAddresses | Sort-Object -Unique) -join ", " } else { "(none)" }
            }
            catch {
                $members            = "Error: $($_.Exception.Message)"
                $membersDisplayName = "Error"
                $membersMailAddress = "Error"
            }

            # ── Build definition count + latest build date ─────────────────────
            try {
                $buildDefinitions     = $buildService.QueryBuildDefinitions($project.Name)
                $buildDefinitionCount = $buildDefinitions.Count

                if ($buildDefinitionCount -gt 0) {
                    $buildSpec                        = $buildService.CreateBuildDetailSpec($project.Name)
                    $buildSpec.MaxBuildsPerDefinition = 1
                    $buildSpec.QueryOrder             = [Microsoft.TeamFoundation.Build.Client.BuildQueryOrder]::FinishTimeDescending
                    $buildSpec.Status                 = [Microsoft.TeamFoundation.Build.Client.BuildStatus]::All

                    $buildResults = $buildService.QueryBuilds($buildSpec)
                    $latestBuild  = $buildResults.Builds | Sort-Object FinishTime -Descending | Select-Object -First 1

                    if ($latestBuild -and $latestBuild.FinishTime -gt [DateTime]::MinValue) {
                        $latestBuildDate = $latestBuild.FinishTime.ToString("yyyy-MM-dd HH:mm")
                    }
                }
            }
            catch {
                $buildDefinitionCount = "Error: $($_.Exception.Message)"
                $latestBuildDate      = "Error"
            }

            # ── Project size + biggest file ────────────────────────────────────
            if ($SkipSize) {
                $latestVersionSize = "Skipped"
                $biggestFileSize   = "Skipped"
                $biggestFileName   = "Skipped"
            }
            else {
                try {
                    $itemSet = $vcs.GetItems(
                        $projectPath,
                        [Microsoft.TeamFoundation.VersionControl.Client.RecursionType]::Full
                    )

                    $files = $itemSet.Items |
                        Where-Object { $_.ItemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::File }

                    # Total project size
                    $totalBytes        = ($files | Measure-Object -Property ContentLength -Sum).Sum
                    #$latestVersionSize = if ($totalBytes) { Format-Bytes $totalBytes } else { "0 B" }
                    $latestVersionSize = if ($totalBytes) { $totalBytes } else { 0 }

                    # Biggest file
                    $biggestFile = $files | Sort-Object ContentLength -Descending | Select-Object -First 1
                    if ($biggestFile) {
                        $biggestFileName = [System.IO.Path]::GetFileName($biggestFile.ServerItem)
                        #$biggestFileSize = Format-Bytes $biggestFile.ContentLength
                        $biggestFileSize = $biggestFile.ContentLength
                    }
                }
                catch {
                    $latestVersionSize = "Error: $($_.Exception.Message)"
                    $biggestFileSize   = "Error"
                    $biggestFileName   = "Error"
                }
            }

            $results += [PSCustomObject]@{
                Collection                = $collection.Name
                CollState                 = $collection.State
                Project                   = $project.Name
                ProjectUri                = $projectUri
                ProjectGuid               = $projectGuid
                ProjectState              = $project.Status
                Description               = $projectDescription
                LastCheckin               = $lastCheckin
                LastUser                  = $lastUser
                LastUserDisplayName       = $lastUserDisplayName
                LastUserMailAddress       = $lastUserMailAddress
                ChangesetCount            = $changesetCount
                ChangesetUsers            = $changesetUsers
                ChangesetUsersDisplayName = $changesetUsersDisplayName
                ChangesetUsersMailAddress = $changesetUsersMailAddress
                Members                   = $members
                MembersDisplayName        = $membersDisplayName
                MembersMailAddress        = $membersMailAddress
                BuildDefinitionCount      = $buildDefinitionCount
                LatestBuildDate           = $latestBuildDate
                LatestVersionSize         = $latestVersionSize
                BiggestFileSize           = $biggestFileSize
                BiggestFile               = $biggestFileName
            }
            Write-Host $($results[-1])
        }
    }
    catch {
        $results += [PSCustomObject]@{
            Collection                = $collection.Name
            CollState                 = $collection.State
            Project                   = "ERROR: $($_.Exception.Message)"
            ProjectUri                = "-"
            ProjectGuid               = "-"
            ProjectState              = "-"
            Description               = "-"
            LastCheckin               = "-"
            LastUser                  = "-"
            LastUserDisplayName       = "-"
            LastUserMailAddress       = "-"
            ChangesetCount            = "-"
            ChangesetUsers            = "-"
            ChangesetUsersDisplayName = "-"
            ChangesetUsersMailAddress = "-"
            Members                   = "-"
            MembersDisplayName        = "-"
            MembersMailAddress        = "-"
            BuildDefinitionCount      = "-"
            LatestBuildDate           = "-"
            LatestVersionSize         = "-"
            BiggestFileSize           = "-"
            BiggestFile               = "-"
        }
    }
}

Write-Host "`n[$(Get-Date)] Done.`n" -ForegroundColor Green

# ─── Display as table ──────────────────────────────────────────────────────────
$results | Format-Table -AutoSize -Property Collection, CollState, Project, ProjectUri, ProjectGuid, ProjectState, Description, LastCheckin, LastUser, LastUserDisplayName, LastUserMailAddress, ChangesetCount, ChangesetUsers, ChangesetUsersDisplayName, ChangesetUsersMailAddress, Members, MembersDisplayName, MembersMailAddress, BuildDefinitionCount, LatestBuildDate, LatestVersionSize, BiggestFileSize, BiggestFile

# Optional: export to CSV
# $results | Export-Csv -Path "C:\tfs_projects_report.csv" -NoTypeInformation -Encoding UTF8