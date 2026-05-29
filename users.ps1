# função que resolve o FQDN de um domínio via DNS
function Get-DomainFQDN {
    param ($shortName)
    
    try {
        # tenta resolver via DNS SRV record do AD
        $srvRecord = Resolve-DnsName -Name "_ldap._tcp.$shortName" -Type SRV -ErrorAction Stop
        return $srvRecord[0].NameTarget -replace '^\w+\.', ''  # extrai o domínio do target
    } catch {
        try {
            # fallback: tenta resolver o nome diretamente
            $resolved = Resolve-DnsName -Name $shortName -Type A -ErrorAction Stop
            return $shortName
        } catch {
            return $null
        }
    }
}

# --- leitura do CSV e extração de utilizadores únicos ---
$csv = Import-Csv -Path "file.csv"

$users = @{}

foreach ($row in $csv) {
    $entries = $row.ChangesetUsers -split ",\s*"
    
    foreach ($entry in $entries) {
        $entry = $entry.Trim()
        if ($entry -match "^(.+)\\(.+)$") {
            $domain = $matches[1]
            $user   = $matches[2]
            $key    = "$domain\$user"
            
            if (-not $users.ContainsKey($key)) {
                $users[$key] = [PSCustomObject]@{
                    Domain = $domain
                    User   = $user
                }
            }
        }
    }
}

# --- resolve FQDNs únicos automaticamente ---
$fqdnCache = @{}

foreach ($domain in ($users.Values.Domain | Select-Object -Unique)) {
    $fqdn = Get-DomainFQDN -shortName $domain
    if ($fqdn) {
        $fqdnCache[$domain] = $fqdn
        Write-Host "✅ Domínio '$domain' resolvido: $fqdn" -ForegroundColor Cyan
    } else {
        Write-Warning "⚠️  Não foi possível resolver o FQDN para o domínio '$domain'"
    }
}

# --- consulta AD para cada utilizador ---
$results = foreach ($entry in ($users.Values | Sort-Object Domain, User)) {
    $fqdn = $fqdnCache[$entry.Domain]
    
    if (-not $fqdn) {
        [PSCustomObject]@{
            Domain      = $entry.Domain
            User        = $entry.User
            Nome        = "FQDN não resolvido"
            Email       = "-"
            Ativo       = "-"
            Bloqueado   = "-"
            UltimoLogin = "-"
        }
        continue
    }

    try {
        $root     = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$fqdn")
        $searcher = New-Object System.DirectoryServices.DirectorySearcher($root)
        $searcher.Filter = "(&(objectClass=user)(sAMAccountName=$($entry.User)))"
        $searcher.PropertiesToLoad.AddRange(@(
            "cn",
            "mail",
            "userAccountControl",
            "lockoutTime",
            "lastLogonTimestamp"
        ))

        $result = $searcher.FindOne()

        if ($result) {
            $uac      = $result.Properties["userAccountControl"][0]
            $disabled = [bool]($uac -band 2)

            $lockout  = $result.Properties["lockoutTime"][0]
            $isLocked = ($lockout -gt 0)

            $lastLogonRaw = $result.Properties["lastLogontimestamp"][0]
            $lastLogon    = if ($lastLogonRaw -and $lastLogonRaw -gt 0) {
                [datetime]::FromFileTime($lastLogonRaw)
            } else {
                "Nunca"
            }

            [PSCustomObject]@{
                Domain      = $entry.Domain
                User        = $entry.User
                Nome        = $result.Properties["cn"][0]
                Email       = $result.Properties["mail"][0]
                Ativo       = -not $disabled
                Bloqueado   = $isLocked
                UltimoLogin = $lastLogon
            }
        } else {
            [PSCustomObject]@{
                Domain      = $entry.Domain
                User        = $entry.User
                Nome        = "NÃO ENCONTRADO"
                Email       = "-"
                Ativo       = "-"
                Bloqueado   = "-"
                UltimoLogin = "-"
            }
        }
    } catch {
        [PSCustomObject]@{
            Domain      = $entry.Domain
            User        = $entry.User
            Nome        = "ERRO: $($_.Exception.Message)"
            Email       = "-"
            Ativo       = "-"
            Bloqueado   = "-"
            UltimoLogin = "-"
        }
    }
}

# --- output ---
$results | Format-Table -AutoSize
$results | Export-Csv -Path "users_ad_info.csv" -NoTypeInformation
Write-Host "`n✅ Resultado exportado para users_ad_info.csv" -ForegroundColor Green