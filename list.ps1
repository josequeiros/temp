# ─── Parameters ───────────────────────────────────────────────────────────────
param(
    [string] $TfsServer = "",
    [switch] $SkipSize
)

# ─── Load required TFS assemblies ─────────────────────────────────────────────
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.TeamFoundation.Client")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.TeamFoundation.Common")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.TeamFoundation.VersionControl.Client")

# ─── Helper: format bytes into human-readable size ────────────────────────────
function Format-Bytes {
    param([long]$Bytes)
    if     ($Bytes -ge 1GB) { "{0:N2} GB" -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { "{0:N2} MB" -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { "{0:N2} KB" -f ($Bytes / 1KB) }
    else                    { "$Bytes B" }
}

# ─── Connect to TFS Configuration Server ──────────────────────────────────────
Write-Host "Connecting to TFS: $TfsServer ..." -ForegroundColor Cyan
$tfsConfigServer = New-Object Microsoft.TeamFoundation.Client.TfsConfigurationServer(
    (New-Object Uri($TfsServer))
)
$tfsConfigServer.Authenticate()

# ─── Get all collections ───────────────────────────────────────────────────────
$collectionService = $tfsConfigServer.GetService([Microsoft.TeamFoundation.Framework.Client.ITeamProjectCollectionService])
$collections = $collectionService.GetCollections()

if ($SkipSize) {
    Write-Host "[-SkipSize] Project size calculation will be skipped.`n" -ForegroundColor Yellow
}

$results = @()

foreach ($collection in $collections) {
    try {
        if ($collection.State -ne "Started") {
            $results += [PSCustomObject]@{
                Collection        = $collection.Name
                CollState         = $collection.State
                Project           = "(Collection not started)"
                ProjectState      = "-"
                LastCheckin       = "-"
                LastUser          = "-"
                LatestVersionSize = "-"
                BiggestFileSize   = "-"
                BiggestFile       = "-"
            }
            continue
        }

        Write-Host "Processing collection: $($collection.Name)" -ForegroundColor Cyan

        $collectionUri = New-Object Uri("$TfsServer/$($collection.Name)")
        $tfsCollection = [Microsoft.TeamFoundation.Client.TfsTeamProjectCollectionFactory]::GetTeamProjectCollection($collectionUri)
        $tfsCollection.Authenticate()

        $cssService = $tfsCollection.GetService([Microsoft.TeamFoundation.Server.ICommonStructureService])
        $projects = $cssService.ListAllProjects()

        $vcs = $tfsCollection.GetService([Microsoft.TeamFoundation.VersionControl.Client.VersionControlServer])

        if ($projects.Count -eq 0) {
            $results += [PSCustomObject]@{
                Collection        = $collection.Name
                CollState         = $collection.State
                Project           = "(No projects found)"
                ProjectState      = "-"
                LastCheckin       = "-"
                LastUser          = "-"
                LatestVersionSize = "-"
                BiggestFileSize   = "-"
                BiggestFile       = "-"
            }
            continue
        }

        foreach ($project in $projects) {
            $projectPath       = "$/" + $project.Name
            $lastCheckin       = "-"
            $lastUser          = "-"
            $latestVersionSize = "-"
            $biggestFileName   = "-"
            $biggestFileSize   = "-"

            Write-Host "  -> $($project.Name)" -ForegroundColor Gray

            # ── Last check-in date + username ──────────────────────────────────
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
                    $lastCheckin = $latestChangeset.CreationDate.ToString("yyyy-MM-dd HH:mm")
                    #$lastUser    = $latestChangeset.CommitterDisplayName  # e.g. "John Doe"
                    # Alternative: $latestChangeset.Committer             # e.g. "DOMAIN\jdoe"
                    $lastUser    = $latestChangeset.Committer
                }
            }
            catch {
                $lastCheckin = "Error: $($_.Exception.Message)"
                $lastUser    = "Error"
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
                    $latestVersionSize = if ($totalBytes) { Format-Bytes $totalBytes } else { "0 B" }

                    # Biggest file
                    $biggestFile = $files | Sort-Object ContentLength -Descending | Select-Object -First 1
                    if ($biggestFile) {
                        $biggestFileName = [System.IO.Path]::GetFileName($biggestFile.ServerItem)
                        $biggestFileSize = Format-Bytes $biggestFile.ContentLength
                    }
                }
                catch {
                    $latestVersionSize = "Error: $($_.Exception.Message)"
                    $biggestFileSize   = "Error"
                    $biggestFileName   = "Error"
                }
            }

            $results += [PSCustomObject]@{
                Collection        = $collection.Name
                CollState         = $collection.State
                Project           = $project.Name
                ProjectState      = $project.Status
                LastCheckin       = $lastCheckin
                LastUser          = $lastUser
                LatestVersionSize = $latestVersionSize
                BiggestFileSize   = $biggestFileSize
                BiggestFile       = $biggestFileName
            }
        }
    }
    catch {
        $results += [PSCustomObject]@{
            Collection        = $collection.Name
            CollState         = $collection.State
            Project           = "ERROR: $($_.Exception.Message)"
            ProjectState      = "-"
            LastCheckin       = "-"
            LastUser          = "-"
            LatestVersionSize = "-"
            BiggestFileSize   = "-"
            BiggestFile       = "-"
        }
    }
}

Write-Host "`nDone.`n" -ForegroundColor Green

# ─── Display as table ──────────────────────────────────────────────────────────
$results | Format-Table -AutoSize -Property Collection, CollState, Project, ProjectState, LastCheckin, LastUser, LatestVersionSize, BiggestFileSize, BiggestFile

# Optional: export to CSV
#$results | Export-Csv -Path "C:\Users\José\Documents\tfs_projects_report.csv" -NoTypeInformation -Encoding UTF8