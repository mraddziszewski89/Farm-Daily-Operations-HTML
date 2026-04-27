<#
.SYNOPSIS
    Farm Daily Operations HTML — SharePoint Subscription Edition
    Automatyczny codzienny raport operacyjny farmy SharePoint on-premise.

.DESCRIPTION
    Skrypt zbiera dane diagnostyczne i operacyjne z całej farmy SharePoint Subscription Edition,
    analizuje ich stan, przypisuje poziom zdrowia (OK / WARNING / CRITICAL) i generuje
    jeden czytelny raport HTML w formie dashboardu administracyjnego.

    Obsługiwane sekcje raportu:
      - Podsumowanie ogólne z krytycznymi problemami
      - Stan serwerów farmy i ich ról MinRole
      - Usługi SharePoint i IIS (w tym App Poole)
      - Aplikacje webowe SharePoint
      - Bazy danych (content + config)
      - Timer Jobs
      - Health Analyzer
      - Windows Event Log (ostatnie N godzin)
      - ULS Logs (ostatnie N godzin)
      - Certyfikaty IIS
      - Zasoby systemowe (dysk, RAM)
      - Spójność wersji farmy

.PARAMETER ReportOutputPath
    Katalog zapisu raportu HTML. Tworzony automatycznie jeśli nie istnieje.
    Domyślnie: C:\SharePoint\Reports

.PARAMETER ReportFileName
    Nazwa pliku HTML raportu (bez ścieżki).
    Domyślnie: FarmDailyOperations.html

.PARAMETER LogHours
    Liczba godzin wstecz do analizy Event Log i ULS.
    Domyślnie: 24

.PARAMETER DiskWarningThresholdGB
    Próg ostrzeżenia wolnego miejsca na dysku [GB].
    Domyślnie: 20

.PARAMETER DiskCriticalThresholdGB
    Próg krytyczny wolnego miejsca na dysku [GB].
    Domyślnie: 10

.PARAMETER RAMWarningThreshold
    Próg ostrzeżenia użycia RAM [%].
    Domyślnie: 80

.PARAMETER RAMCriticalThreshold
    Próg krytyczny użycia RAM [%].
    Domyślnie: 90

.PARAMETER ArchiveReports
    Przełącznik: jeśli podany, zapisuje również wersję archiwalną raportu
    z datą i godziną w nazwie pliku.

.PARAMETER SendEmail
    Przełącznik: jeśli podany, wysyła raport e-mailem po wygenerowaniu.

.PARAMETER SMTPServer
    Adres serwera SMTP.

.PARAMETER SMTPPort
    Port SMTP. Domyślnie: 25

.PARAMETER EmailFrom
    Adres nadawcy e-mail.

.PARAMETER EmailTo
    Tablica adresów odbiorców e-mail.

.PARAMETER EmailSubject
    Temat wiadomości e-mail.

.PARAMETER EmailBody
    Treść wiadomości e-mail (raport HTML dołączany jako załącznik).

.PARAMETER UseSSL
    Przełącznik: użyj SSL/TLS przy połączeniu SMTP.

.PARAMETER SMTPCredential
    Opcjonalne poświadczenia do uwierzytelnienia SMTP.

.EXAMPLE
    # Uruchomienie podstawowe z parametrami domyślnymi:
    .\FarmDailyOperations.ps1

.EXAMPLE
    # Pełne uruchomienie z archiwizacją i wysyłką e-mail przez SMTP z SSL:
    .\FarmDailyOperations.ps1 `
        -ReportOutputPath "D:\Reports\SharePoint" `
        -ReportFileName "FarmReport.html" `
        -LogHours 24 `
        -DiskWarningThresholdGB 25 `
        -DiskCriticalThresholdGB 10 `
        -RAMWarningThreshold 75 `
        -RAMCriticalThreshold 90 `
        -ArchiveReports `
        -SendEmail `
        -SMTPServer "mail.contoso.com" `
        -SMTPPort 587 `
        -EmailFrom "sharepoint-monitor@contoso.com" `
        -EmailTo @("admin@contoso.com","helpdesk@contoso.com") `
        -EmailSubject "SharePoint Farm Report - $(Get-Date -Format 'yyyy-MM-dd')" `
        -UseSSL `
        -SMTPCredential (Get-Credential)

.NOTES
    Wymagania:
      - SharePoint Subscription Edition (serwer z zainstalowanymi binariami SP)
      - PowerShell 5.1+
      - Uruchomienie jako Administrator lokalny i Farm Administrator w SharePoint
      - Moduł Microsoft.SharePoint.PowerShell (snapin lub autoload)
      - Moduł WebAdministration (dla danych IIS)

    Autor: Farm Daily Operations HTML Script
    Wersja: 1.0
    Platforma: SharePoint Subscription Edition on-premise
#>

[CmdletBinding()]
param(
    [string]$ReportOutputPath    = "C:\SharePoint\Reports",
    [string]$ReportFileName      = "FarmDailyOperations.html",
    [int]   $LogHours            = 24,
    [int]   $DiskWarningThresholdGB  = 20,
    [int]   $DiskCriticalThresholdGB = 10,
    [int]   $RAMWarningThreshold = 80,
    [int]   $RAMCriticalThreshold= 90,
    [switch]$ArchiveReports,

    # Parametry e-mail
    [switch]$SendEmail,
    [string]$SMTPServer          = "",
    [int]   $SMTPPort            = 25,
    [string]$EmailFrom           = "",
    [string[]]$EmailTo           = @(),
    [string]$EmailSubject        = "Farm Daily Operations - $(Get-Date -Format 'yyyy-MM-dd')",
    [string]$EmailBody           = "W zalaczeniu dzienny raport operacyjny farmy SharePoint.",
    [switch]$UseSSL,
    [System.Management.Automation.PSCredential]$SMTPCredential = $null
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

#region ─── KONFIGURACJA GLOBALNA ───────────────────────────────────────────────

$Script:StartTime    = Get-Date
$Script:ScriptServer = $env:COMPUTERNAME
$Script:LogFile      = Join-Path $ReportOutputPath "FarmDailyOperations_ScriptLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$Script:SPLoaded     = $false
$Script:FarmName     = "Nieznana farma"

# Zmienna akumulująca wszystkie wykryte problemy (dla sekcji Summary)
$Script:AllIssues    = [System.Collections.Generic.List[PSCustomObject]]::new()

# Minimalna liczba znaków ULS-a do wyświetlenia w raporcie
$Script:ULSMaxEntries = 200

#endregion

#region ─── FUNKCJA LOGOWANIA ───────────────────────────────────────────────────

function Write-Log {
    <#
    .SYNOPSIS
        Zapisuje wpis do pliku logu skryptu i opcjonalnie do konsoli.
    #>
    param(
        [string]$Message,
        [ValidateSet("INFO","WARNING","ERROR","DEBUG")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp][$Level] $Message"
    try {
        # Upewnij sie, ze katalog logu istnieje
        # Upewnij sie, ze katalog logu istnieje
        # Upewnij się, że katalog logu istnieje
        $logDir = Split-Path $Script:LogFile -Parent
        if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
        Add-Content -Path $Script:LogFile -Value $line -Encoding UTF8
    } catch { <# ignoruj błędy logowania #> }
    switch ($Level) {
        "ERROR"   { Write-Host $line -ForegroundColor Red }
        "WARNING" { Write-Host $line -ForegroundColor Yellow }
        "DEBUG"   { Write-Verbose $line }
        default   { Write-Host $line -ForegroundColor Cyan }
    }
}

#endregion

#region ─── HELPER: DODAJ PROBLEM DO LISTY ZBIORCZEJ ───────────────────────────

function Add-Issue {
    <#
    .SYNOPSIS
        Dodaje wykryty problem do centralnej listy problemów (używanej w Summary).
    #>
    param(
        [string]$Section,
        [string]$ObjectName,
        [string]$Description,
        [ValidateSet("WARNING","CRITICAL")]
        [string]$Severity,
        [string]$Server      = "",
        [string]$Recommendation = ""
    )
    $Script:AllIssues.Add([PSCustomObject]@{
        Section        = $Section
        ObjectName     = $ObjectName
        Description    = $Description
        Severity       = $Severity
        Server         = $Server
        Recommendation = $Recommendation
    })
}

#endregion

#region ─── WALIDACJA UPRAWNIEŃ ADMINISTRATORA ──────────────────────────────────

function Test-AdminPrivileges {
    <#
    .SYNOPSIS
        Sprawdza, czy skrypt uruchomiono z uprawnieniami lokalnego administratora.
        Ostrzega (nie przerywa), jeśli uprawnienia są niewystarczające.
    #>
    Write-Log "Sprawdzanie uprawnien administratora..."
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal   = New-Object Security.Principal.WindowsPrincipal($currentUser)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Log "OSTRZEZENIE: Skrypt nie jest uruchomiony jako Administrator. Niektore dane moga byc niedostepne." -Level WARNING
        return $false
    }
    Write-Log "Uprawnienia administratora: OK"
    return $true
}

#endregion

#region ─── INICJALIZACJA ŚRODOWISKA SHAREPOINT ─────────────────────────────────

function Initialize-SharePointEnvironment {
    <#
    .SYNOPSIS
        Ładuje snapin lub moduł SharePoint PowerShell.
        W SharePoint Subscription Edition moduł jest ładowany automatycznie
        przez pshell.exe; w normalnym PS trzeba go załadować ręcznie.
    #>
    Write-Log "Ladowanie srodowiska SharePoint PowerShell..."
    try {
        # Próba 1: użyj snapina (klasyczna ścieżka SP2019/SPSE)
        if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null) {
            Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
            Write-Log "Snapin Microsoft.SharePoint.PowerShell zaladowany pomyslnie."
        } else {
            Write-Log "Snapin Microsoft.SharePoint.PowerShell juz zaladowany."
        }
        $Script:SPLoaded = $true
    }
    catch {
        Write-Log "Nie udalo sie zaladowac snapina SP: $($_.Exception.Message)" -Level WARNING
        # Próba 2: moduł PowerShell (nowsze instalacje)
        try {
            Import-Module Microsoft.SharePoint.PowerShell -ErrorAction Stop
            $Script:SPLoaded = $true
            Write-Log "Modul Microsoft.SharePoint.PowerShell zaladowany pomyslnie."
        }
        catch {
            Write-Log "Nie udalo sie zaladowac modulu SP: $($_.Exception.Message)" -Level ERROR
            $Script:SPLoaded = $false
        }
    }

    # Pobierz nazwę farmy
    if ($Script:SPLoaded) {
        try {
            $farm = Get-SPFarm -ErrorAction Stop
            $Script:FarmName = if ($farm.Name) { $farm.Name } else { "SharePoint Farm" }
            Write-Log "Polaczono z farma: $($Script:FarmName)"
        }
        catch {
            Write-Log "Nie mozna pobrac obiektu farmy: $($_.Exception.Message)" -Level WARNING
            $Script:FarmName = "SharePoint Farm (brak dostepu)"
        }
    }
}

#endregion

#region ─── SEKCJA 1: SERWERY FARMY I ROLE MINROLE ──────────────────────────────

function Get-FarmServers {
    <#
    .SYNOPSIS
        Zbiera dane o wszystkich serwerach w farmie: rola MinRole, stan,
        wersja binariów SharePoint.
    .OUTPUTS
        Tablica PSCustomObject z polami: ServerName, Role, Status, BuildVersion,
        NeedsUpgrade, StatusLevel
    #>
    Write-Log "Pobieranie danych serwerow farmy..."
    $results = @()

    if (-not $Script:SPLoaded) {
        Write-Log "SP nie zaladowany - pomijam sekcje serwerow." -Level WARNING
        return @([PSCustomObject]@{ ServerName="N/A"; Role="Brak danych"; Status="Niedostepne";
            BuildVersion="N/A"; NeedsUpgrade="N/A"; StatusLevel="WARNING" })
    }

    try {
        $spServers = Get-SPServer -ErrorAction Stop | Where-Object { $_.Role -ne "Invalid" }
        foreach ($srv in $spServers) {
            $statusLevel = "OK"
            $needsUpgrade = "Nie"

            # Stan serwera
            $srvStatus = try { $srv.Status.ToString() } catch { "Nieznany" }
            if ($srvStatus -ne "Online") {
                $statusLevel = "CRITICAL"
                Add-Issue -Section "Serwery" -ObjectName $srv.DisplayName `
                    -Description "Serwer jest w stanie: $srvStatus" `
                    -Severity "CRITICAL" -Server $srv.DisplayName `
                    -Recommendation "Sprawdz stan serwera i uslug Windows. Zweryfikuj dzienniki zdarzen."
            }

            # Wersja i potrzeba aktualizacji - próba z SPServer.BuildVersion
            $buildVersion = try {
                $bv = $srv.BuildVersion
                if ($null -ne $bv -and $bv.ToString() -notin @("","0.0.0.0")) { $bv.ToString() } else { $null }
            } catch { $null }
            # Fallback: odczyt wersji z rejestru Windows na serwerze
            if (-not $buildVersion) {
                $regGetBuild = {
                    $regPath = "HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\16.0\WSS"
                    try { (Get-ItemProperty -Path $regPath -Name "Build" -ErrorAction Stop).Build } catch { "N/A" }
                }
                $buildVersion = try {
                    if ($srv.DisplayName -ieq $env:COMPUTERNAME) { & $regGetBuild }
                    else { Invoke-Command -ComputerName $srv.DisplayName -ScriptBlock $regGetBuild -ErrorAction Stop }
                } catch { "N/A" }
                if (-not $buildVersion) { $buildVersion = "N/A" }
            }
            try {
                if ($srv.NeedsUpgrade -eq $true) {
                    $needsUpgrade = "TAK"
                    $statusLevel  = if ($statusLevel -eq "OK") { "WARNING" } else { $statusLevel }
                    Add-Issue -Section "Serwery" -ObjectName $srv.DisplayName `
                        -Description "Serwer wymaga aktualizacji bazy danych SharePoint (NeedsUpgrade=True)" `
                        -Severity "WARNING" -Server $srv.DisplayName `
                        -Recommendation "Uruchom psconfig.exe -cmd upgrade -inplace b2b lub Start-SPConfigurationWizard."
                }
            } catch { }

            # Rola MinRole
            $role = try { $srv.Role.ToString() } catch { "Nieznana" }

            $results += [PSCustomObject]@{
                ServerName   = $srv.DisplayName
                Role         = $role
                Status       = $srvStatus
                BuildVersion = $buildVersion
                NeedsUpgrade = $needsUpgrade
                StatusLevel  = $statusLevel
            }
        }
        Write-Log "Pobrano dane $($results.Count) serwerow farmy."
    }
    catch {
        Write-Log "Blad pobierania serwerow farmy: $($_.Exception.Message)" -Level ERROR
        $results += [PSCustomObject]@{ ServerName="BLAD"; Role="Blad"; Status="Blad pobierania";
            BuildVersion="N/A"; NeedsUpgrade="N/A"; StatusLevel="CRITICAL" }
    }
    return $results
}

#endregion

#region ─── SEKCJA 2: USŁUGI SHAREPOINT ─────────────────────────────────────────

function Get-SharePointServices {
    <#
    .SYNOPSIS
        Pobiera stan wszystkich instancji usług SharePoint na wszystkich serwerach.
    .OUTPUTS
        Tablica PSCustomObject: ServerName, ServiceName, ServiceType, Status, StatusLevel
    #>
    Write-Log "Pobieranie stanu uslug SharePoint..."
    $results = @()

    if (-not $Script:SPLoaded) {
        return @([PSCustomObject]@{ ServerName="N/A"; ServiceName="Brak danych"; ServiceType="N/A";
            Status="Niedostepne"; StatusLevel="WARNING" })
    }

    try {
        $serviceInstances = Get-SPServiceInstance -ErrorAction Stop
        foreach ($si in $serviceInstances) {
            $statusLevel = "OK"
            $status      = try { $si.Status.ToString() } catch { "Nieznany" }
            $srvName     = try { $si.Server.DisplayName } catch { "Nieznany" }
            $svcName     = try { $si.TypeName } catch { $si.GetType().Name }

            # Usługi w stanie innym niż Online lub Disabled są podejrzane
            if ($status -notin @("Online","Disabled","Provisioning","Unprovisioning")) {
                $statusLevel = "WARNING"
                Add-Issue -Section "Uslugi SharePoint" -ObjectName $svcName `
                    -Description "Usluga '$svcName' na serwerze '$srvName' ma stan: $status" `
                    -Severity "WARNING" -Server $srvName `
                    -Recommendation "Sprawdz stan uslugi w Central Administration > System Settings > Manage services on server. Zrestartuj usluge jesli to mozliwe."
            }

            $results += [PSCustomObject]@{
                ServerName   = $srvName
                ServiceName  = $svcName
                ServiceType  = $(try { $si.GetType().Name } catch { "N/A" })
                Status       = $status
                StatusLevel  = $statusLevel
            }
        }
        Write-Log "Pobrano $($results.Count) instancji uslug SharePoint."
    }
    catch {
        Write-Log "Blad pobierania uslug SP: $($_.Exception.Message)" -Level ERROR
        $results += [PSCustomObject]@{ ServerName="BLAD"; ServiceName="Blad pobierania";
            ServiceType="N/A"; Status="Blad"; StatusLevel="CRITICAL" }
    }
    return $results
}

#endregion

#region ─── SEKCJA 3: IIS – WITRYNY I APP POOLE ─────────────────────────────────

function Get-IISStatus {
    <#
    .SYNOPSIS
        Sprawdza stan witryn IIS i Application Pools na lokalnym serwerze
        (i opcjonalnie zdalnie przez Invoke-Command na innych serwerach farmy).
    .OUTPUTS
        Hashtable: Sites (tablica) i AppPools (tablica)
    #>
    Write-Log "Pobieranie stanu IIS..."

    $allSites    = [System.Collections.Generic.List[PSCustomObject]]::new()
    $allAppPools = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Lista serwerów do sprawdzenia (zbieramy z farmy SP lub używamy lokalnego)
    $serversToCheck = @($env:COMPUTERNAME)
    if ($Script:SPLoaded) {
        try {
            $spSrvList = Get-SPServer -ErrorAction Stop | Where-Object { $_.Role -ne "Invalid" } |
                         Select-Object -ExpandProperty DisplayName
            $serversToCheck = $spSrvList
        } catch { }
    }

    foreach ($srvName in $serversToCheck) {
        $isLocal = ($srvName -ieq $env:COMPUTERNAME)
        $sb = {
            param($srvCtx)
            $r = @{ Sites = @(); AppPools = @() }
            try {
                Import-Module WebAdministration -ErrorAction Stop

                # Witryny IIS
                $sites = Get-WebSite -ErrorAction Stop
                foreach ($site in $sites) {
                    $r.Sites += [PSCustomObject]@{
                        ServerName  = $srvCtx
                        SiteName    = $site.Name
                        State       = $site.State
                        PhysPath    = $site.physicalPath
                        Bindings    = ($site.Bindings.Collection | ForEach-Object { "$($_.Protocol)://$($_.BindingInformation)" }) -join "; "
                        StatusLevel = if ($site.State -eq "Started") { "OK" } else { "WARNING" }
                    }
                }

                # App Poole
                $pools = Get-WebConfiguration -Filter "system.applicationHost/applicationPools/add" -ErrorAction Stop
                foreach ($pool in $pools) {
                    $state = (Get-WebAppPoolState -Name $pool.Name -ErrorAction SilentlyContinue).Value
                    $r.AppPools += [PSCustomObject]@{
                        ServerName   = $srvCtx
                        PoolName     = $pool.Name
                        State        = if ($state) { $state } else { "Nieznany" }
                        ManagedRuntime = $pool.ManagedRuntimeVersion
                        Identity     = $pool.ProcessModel.userName
                        StatusLevel  = if ($state -eq "Started") { "OK" } elseif ($state -eq "Stopped") { "WARNING" } else { "CRITICAL" }
                    }
                }
            }
            catch {
                $r.Sites   += [PSCustomObject]@{ ServerName=$srvCtx; SiteName="BLAD"; State="Blad IIS: $($_.Exception.Message)"; PhysPath=""; Bindings=""; StatusLevel="CRITICAL" }
                $r.AppPools += [PSCustomObject]@{ ServerName=$srvCtx; PoolName="BLAD"; State="Blad IIS"; ManagedRuntime=""; Identity=""; StatusLevel="CRITICAL" }
            }
            return $r
        }

        try {
            if ($isLocal) {
                $data = & $sb -srvCtx $srvName
            } else {
                $data = Invoke-Command -ComputerName $srvName -ScriptBlock $sb -ArgumentList $srvName `
                        -ErrorAction Stop
            }
            foreach ($s in $data.Sites)    { $allSites.Add($s) }
            foreach ($p in $data.AppPools) { $allAppPools.Add($p) }
        }
        catch {
            Write-Log "Blad pobierania IIS z $srvName : $($_.Exception.Message)" -Level WARNING
            $allSites.Add([PSCustomObject]@{ ServerName=$srvName; SiteName="Blad"; State="Nie mozna polaczyc";
                PhysPath=""; Bindings=""; StatusLevel="WARNING" })
        }
    }

    # Zgłoś problemy
    foreach ($s in $allSites) {
        if ($s.StatusLevel -ne "OK") {
            Add-Issue -Section "IIS Witryny" -ObjectName $s.SiteName `
                -Description "Witryna IIS '$($s.SiteName)' na '$($s.ServerName)' ma stan: $($s.State)" `
                -Severity "WARNING" -Server $s.ServerName `
                -Recommendation "Sprawdz stan witryny w IIS Manager lub wykonaj 'iisreset /status'. Zweryfikuj logi IIS."
        }
    }
    foreach ($p in $allAppPools) {
        if ($p.StatusLevel -eq "WARNING") {
            Add-Issue -Section "IIS App Poole" -ObjectName $p.PoolName `
                -Description "App Pool '$($p.PoolName)' na '$($p.ServerName)' jest zatrzymany." `
                -Severity "WARNING" -Server $p.ServerName `
                -Recommendation "Uruchom App Pool w IIS Manager lub: Start-WebAppPool -Name '$($p.PoolName)'."
        } elseif ($p.StatusLevel -eq "CRITICAL") {
            Add-Issue -Section "IIS App Poole" -ObjectName $p.PoolName `
                -Description "App Pool '$($p.PoolName)' na '$($p.ServerName)' ma nieznany/krytyczny stan." `
                -Severity "CRITICAL" -Server $p.ServerName `
                -Recommendation "Sprawdz stan App Poola w IIS. Zweryfikuj konto uslugowe i logi zdarzen aplikacji."
        }
    }

    Write-Log "IIS: $($allSites.Count) witryn, $($allAppPools.Count) app poolow."
    return @{ Sites = $allSites.ToArray(); AppPools = $allAppPools.ToArray() }
}

#endregion

#region ─── SEKCJA 4: APLIKACJE WEBOWE SHAREPOINT ──────────────────────────────

function Get-WebApplications {
    <#
    .SYNOPSIS
        Pobiera listę aplikacji webowych SharePoint wraz z ich URL-ami i stanem.
    .OUTPUTS
        Tablica PSCustomObject: Name, URL, ContentDBCount, AllowAnon, Status, StatusLevel
    #>
    Write-Log "Pobieranie aplikacji webowych SharePoint..."
    $results = @()

    if (-not $Script:SPLoaded) {
        return @([PSCustomObject]@{ Name="Brak danych"; URL="N/A"; ContentDBCount=0;
            AllowAnon="N/A"; Status="Niedostepne"; StatusLevel="WARNING" })
    }

    try {
        $webApps = Get-SPWebApplication -ErrorAction Stop
        foreach ($wa in $webApps) {
            $statusLevel = "OK"
            $status      = "Online"

            # Sprawdź dostępność HTTP (prosta próba połączenia)
            try {
                $uri = [System.Uri]$wa.Url
                $req = [System.Net.HttpWebRequest]::Create($wa.Url)
                $req.Timeout = 10000
                $req.AllowAutoRedirect = $false
                $resp = $req.GetResponse()
                $resp.Close()
            }
            catch [System.Net.WebException] {
                $httpResponse = $_.Exception.Response
                if ($httpResponse) {
                    $httpCode = [int]$httpResponse.StatusCode
                    if ($httpCode -ge 500) {
                        $status      = "Blad HTTP $httpCode"
                        $statusLevel = "CRITICAL"
                        Add-Issue -Section "Aplikacje webowe" -ObjectName $wa.Name `
                            -Description "Aplikacja webowa '$($wa.Name)' zwraca blad HTTP $httpCode" `
                            -Severity "CRITICAL" -Server $env:COMPUTERNAME `
                            -Recommendation "Sprawdz logi IIS i ULS dla aplikacji '$($wa.Name)'. Zweryfikuj App Pool i bazy danych."
                    } elseif ($httpCode -in @(302, 401, 403)) {
                        $status = "HTTP $httpCode (OK - przekierowanie/auth)"
                    } else {
                        $status      = "Niedostepna (HTTP $httpCode)"
                        $statusLevel = "WARNING"
                    }
                } else {
                    $status      = "Blad polaczenia HTTP: $($_.Exception.Message)"
                    $statusLevel = "WARNING"
                }
            }
            catch {
                $status      = "Blad polaczenia: $($_.Exception.Message.Substring(0, [Math]::Min(80,$_.Exception.Message.Length)))"
                $statusLevel = "WARNING"
                Add-Issue -Section "Aplikacje webowe" -ObjectName $wa.Name `
                    -Description "Nie mozna polaczyc sie z aplikacja webowa '$($wa.Name)': $status" `
                    -Severity "WARNING" -Server $env:COMPUTERNAME `
                    -Recommendation "Sprawdz czy IIS jest uruchomiony i czy App Pool jest aktywny. Zweryfikuj DNS i certyfikaty SSL."
            }

            $results += [PSCustomObject]@{
                Name          = $wa.Name
                URL           = $wa.Url
                ContentDBCount= $(try { $wa.ContentDatabases.Count } catch { 0 })
                AllowAnon     = $(try { $wa.AllowAnonymousAccess.ToString() } catch { "N/A" })
                Status        = $status
                StatusLevel   = $statusLevel
            }
        }
        Write-Log "Pobrano $($results.Count) aplikacji webowych."
    }
    catch {
        Write-Log "Blad pobierania aplikacji webowych: $($_.Exception.Message)" -Level ERROR
        $results += [PSCustomObject]@{ Name="BLAD"; URL="N/A"; ContentDBCount=0;
            AllowAnon="N/A"; Status="Blad pobierania"; StatusLevel="CRITICAL" }
    }
    return $results
}

#endregion

#region ─── SEKCJA 5: BAZY DANYCH ───────────────────────────────────────────────

function Get-Databases {
    <#
    .SYNOPSIS
        Pobiera stan baz danych SharePoint (content databases + config + admin).
        Sprawdza: stan, rozmiar, NeedsUpgrade, ReadOnly.
    .OUTPUTS
        Tablica PSCustomObject: DBName, Type, Server, Size_MB, Status, NeedsUpgrade,
        ReadOnly, StatusLevel
    #>
    Write-Log "Pobieranie stanu baz danych SharePoint..."
    $results = @()

    if (-not $Script:SPLoaded) {
        return @([PSCustomObject]@{ DBName="Brak danych"; Type="N/A"; Server="N/A";
            Size_MB=0; Status="Niedostepne"; NeedsUpgrade="N/A"; ReadOnly="N/A"; StatusLevel="WARNING" })
    }

    try {
        # Bazy danych treści
        if ($false) {
        $contentDBs = Get-SPContentDatabase -ErrorAction Stop
        foreach ($db in $contentDBs) {
            $statusLevel = "OK"
            $status      = try { $db.Status.ToString() } catch { "Nieznany" }
            $readOnly    = try { $db.IsReadOnly.ToString() } catch { "N/A" }
            $needsUpgrade= try { $db.NeedsUpgrade.ToString() } catch { "N/A" }
            $sizeBytes   = try { $db.DiskSizeRequired } catch { 0 }
            $sizeMB      = [Math]::Round($sizeBytes / 1MB, 1)

            if ($status -ne "Online") {
                $statusLevel = "CRITICAL"
                Add-Issue -Section "Bazy danych" -ObjectName $db.Name `
                    -Description "Baza danych treści '$($db.Name)' ma stan: $status" `
                    -Severity "CRITICAL" -Server $db.Server `
                    -Recommendation "Sprawdz dostepnosc SQL Server, state bazy danych w SSMS. Zweryfikuj logi SQL."
            }
            if ($needsUpgrade -eq "True") {
                $statusLevel = if ($statusLevel -eq "OK") { "WARNING" } else { $statusLevel }
                Add-Issue -Section "Bazy danych" -ObjectName $db.Name `
                    -Description "Baza danych '$($db.Name)' wymaga aktualizacji schematu." `
                    -Severity "WARNING" -Server $db.Server `
                    -Recommendation "Uruchom kreator konfiguracji: psconfig.exe -cmd upgrade -inplace b2b"
            }
            if ($readOnly -eq "True") {
                $statusLevel = if ($statusLevel -eq "OK") { "WARNING" } else { $statusLevel }
                Add-Issue -Section "Bazy danych" -ObjectName $db.Name `
                    -Description "Baza danych '$($db.Name)' jest w trybie tylko do odczytu." `
                    -Severity "WARNING" -Server $db.Server `
                    -Recommendation "Sprawdz i zmodyfikuj ustawienia bazy danych w SQL Server Management Studio."
            }

            $results += [PSCustomObject]@{
                DBName       = $db.Name
                Type         = "Content"
                Server       = $(try { $db.Server } catch { "N/A" })
                Size_MB      = $sizeMB
                Status       = $status
                NeedsUpgrade = $needsUpgrade
                ReadOnly     = $readOnly
                StatusLevel  = $statusLevel
            }
        }

        # Bazy administracyjne (config + central admin)
        $farm = Get-SPFarm -ErrorAction Stop
        $adminDBs = @($farm.Database)
        try {
            $caWA = Get-SPWebApplication -IncludeCentralAdministration -ErrorAction SilentlyContinue |
                    Where-Object { $_.IsAdministrationWebApplication }
            if ($caWA) { $adminDBs += $caWA.ContentDatabases }
        } catch { }

        foreach ($db in $adminDBs) {
            if ($null -eq $db) { continue }
            $statusLevel = "OK"
            $status      = try { $db.Status.ToString() } catch { "Nieznany" }
            if ($status -ne "Online") {
                $statusLevel = "CRITICAL"
                Add-Issue -Section "Bazy danych" -ObjectName $db.Name `
                    -Description "Baza konfiguracyjna '$($db.Name)' ma stan: $status" `
                    -Severity "CRITICAL" -Server $(try { $db.Server } catch { "N/A" }) `
                    -Recommendation "To jest krytyczna baza farmy. Sprawdz SQL Server natychmiast."
            }
            $results += [PSCustomObject]@{
                DBName       = $db.Name
                Type         = "Admin/Config"
                Server       = $(try { $db.Server } catch { "N/A" })
                Size_MB      = $(try { [Math]::Round($db.DiskSizeRequired / 1MB, 1) } catch { 0 })
                Status       = $status
                NeedsUpgrade = $(try { $db.NeedsUpgrade.ToString() } catch { "N/A" })
                ReadOnly     = $(try { $db.IsReadOnly.ToString() } catch { "N/A" })
                StatusLevel  = $statusLevel
            }
        }

        }
        $seenDbNames = @{}

        function Add-DatabaseRecord {
            param(
                [Parameter(Mandatory = $true)]$Database,
                [Parameter(Mandatory = $true)][string]$TypeLabel
            )

            if ($null -eq $Database) { return }

            $dbName = $(try { $Database.Name } catch { $null })
            if ([string]::IsNullOrWhiteSpace($dbName)) { return }
            if ($seenDbNames.ContainsKey($dbName)) { return }
            $seenDbNames[$dbName] = $true

            $statusLevel  = "OK"
            $status       = $(try { $Database.Status.ToString() } catch { "Nieznany" })
            $readOnly     = $(try { $Database.IsReadOnly.ToString() } catch { "N/A" })
            $needsUpgrade = $(try { $Database.NeedsUpgrade.ToString() } catch { "N/A" })
            $serverName   = $(try { $Database.Server } catch { "N/A" })
            $sizeMB       = $(try { [Math]::Round(($Database.DiskSizeRequired / 1MB), 1) } catch { 0 })

            if ($status -ne "Online") {
                $statusLevel = "CRITICAL"
                Add-Issue -Section "Bazy danych" -ObjectName $dbName `
                    -Description "Baza danych '$dbName' ($TypeLabel) ma stan: $status" `
                    -Severity "CRITICAL" -Server $serverName `
                    -Recommendation "Sprawdz dostepnosc SQL Server, stan bazy w SSMS oraz logi SQL i ULS."
            }
            if ($needsUpgrade -eq "True") {
                $statusLevel = if ($statusLevel -eq "OK") { "WARNING" } else { $statusLevel }
                Add-Issue -Section "Bazy danych" -ObjectName $dbName `
                    -Description "Baza danych '$dbName' ($TypeLabel) wymaga aktualizacji schematu." `
                    -Severity "WARNING" -Server $serverName `
                    -Recommendation "Uruchom psconfig.exe -cmd upgrade -inplace b2b lub Start-SPConfigurationWizard."
            }
            if ($readOnly -eq "True") {
                $statusLevel = if ($statusLevel -eq "OK") { "WARNING" } else { $statusLevel }
                Add-Issue -Section "Bazy danych" -ObjectName $dbName `
                    -Description "Baza danych '$dbName' ($TypeLabel) jest w trybie tylko do odczytu." `
                    -Severity "WARNING" -Server $serverName `
                    -Recommendation "Sprawdz i zmodyfikuj ustawienia bazy danych w SQL Server Management Studio."
            }

            return [PSCustomObject]@{
                DBName       = $dbName
                Type         = $TypeLabel
                Server       = $serverName
                Size_MB      = $sizeMB
                Status       = $status
                NeedsUpgrade = $needsUpgrade
                ReadOnly     = $readOnly
                StatusLevel  = $statusLevel
            }
        }

        foreach ($db in @(Get-SPContentDatabase -ErrorAction Stop)) {
            $rec = Add-DatabaseRecord -Database $db -TypeLabel "Content"
            if ($null -ne $rec) { $results += $rec }
        }

        $otherDBs = @()
        try {
            $otherDBs = @(Get-SPDatabase -ErrorAction Stop | Where-Object {
                $_ -and ($_ -isnot [Microsoft.SharePoint.Administration.SPContentDatabase])
            })
        }
        catch {
            Write-Log "Get-SPDatabase zwrocil blad, wlaczam tryb zgodnosci dla baz farmy: $($_.Exception.Message)" -Level WARNING
        }

        if ($otherDBs.Count -eq 0) {
            $farm = Get-SPFarm -ErrorAction Stop
            if ($farm.PSObject.Properties.Match("ConfigurationDatabase").Count -gt 0 -and $farm.ConfigurationDatabase) {
                $otherDBs += $farm.ConfigurationDatabase
            }
            elseif ($farm.PSObject.Properties.Match("Database").Count -gt 0 -and $farm.Database) {
                $otherDBs += $farm.Database
            }
        }

        try {
            $caWA = Get-SPWebApplication -IncludeCentralAdministration -ErrorAction SilentlyContinue |
                Where-Object { $_.IsAdministrationWebApplication } |
                Select-Object -First 1
            if ($caWA) {
                $otherDBs += @($caWA.ContentDatabases)
            }
        } catch { }

        foreach ($db in @($otherDBs)) {
            $typeLabel = $(try {
                if ($db -is [Microsoft.SharePoint.Administration.SPContentDatabase]) {
                    if ($(try { $db.WebApplication.IsAdministrationWebApplication } catch { $false })) {
                        "CentralAdmin Content"
                    } else {
                        "Content"
                    }
                } else {
                    $rawType = $(try { $db.TypeName } catch { $db.GetType().Name })
                    if ([string]::IsNullOrWhiteSpace($rawType)) { "Admin/Service/Config" } else { $rawType }
                }
            } catch { "Admin/Service/Config" })

            $rec = Add-DatabaseRecord -Database $db -TypeLabel $typeLabel
            if ($null -ne $rec) { $results += $rec }
        }

        if ($results.Count -eq 0) {
            $results += [PSCustomObject]@{
                DBName       = "Brak danych"
                Type         = "N/A"
                Server       = "N/A"
                Size_MB      = 0
                Status       = "Niedostepne"
                NeedsUpgrade = "N/A"
                ReadOnly     = "N/A"
                StatusLevel  = "WARNING"
            }
        }

        Write-Log "Pobrano dane $($results.Count) baz danych."
    }
    catch {
        Write-Log "Blad pobierania baz danych: $($_.Exception.Message)" -Level ERROR
        $results += [PSCustomObject]@{ DBName="BLAD"; Type="N/A"; Server="N/A";
            Size_MB=0; Status="Blad pobierania"; NeedsUpgrade="N/A"; ReadOnly="N/A"; StatusLevel="CRITICAL" }
    }
    return $results
}

#endregion

#region ─── SEKCJA 6: TIMER JOBS ────────────────────────────────────────────────

function Get-TimerJobStatus {
    <#
    .SYNOPSIS
        Pobiera ostatnie uruchomienia Timer Jobów SharePoint.
        Wykrywa zadania, które zakończyły się błędem lub nie uruchamiały się
        przez dłuższy czas niż oczekiwany interwał.
    .OUTPUTS
        Tablica PSCustomObject: JobName, Status, LastRunTime, Duration_s,
        Server, StatusLevel
    #>
    Write-Log "Pobieranie stanu Timer Jobs..."
    $results = @()

    if (-not $Script:SPLoaded) {
        return @([PSCustomObject]@{ JobName="Brak danych"; Status="Niedostepne";
            LastRunTime="N/A"; Duration_s=0; Server="N/A"; StatusLevel="WARNING" })
    }

    try {
        $jobHistory = Get-SPTimerJob -ErrorAction Stop
        foreach ($job in $jobHistory) {
            $statusLevel = "OK"
            $lastRun     = "Nigdy"
            $duration    = 0
            $lastStatus  = "Nieznany"

            try {
                $history = $job.HistoryEntries | Sort-Object StartTime -Descending | Select-Object -First 1
                if ($history) {
                    $lastRun    = $history.StartTime.ToString("yyyy-MM-dd HH:mm:ss")
                    $duration   = [Math]::Round(($history.EndTime - $history.StartTime).TotalSeconds, 1)
                    $lastStatus = $history.Status.ToString()

                    if ($lastStatus -eq "Failed") {
                        $statusLevel = "CRITICAL"
                        Add-Issue -Section "Timer Jobs" -ObjectName $job.DisplayName `
                            -Description "Timer Job '$($job.DisplayName)' zakonczyl sie blediem w: $lastRun" `
                            -Severity "CRITICAL" -Server $(try { $job.Server.DisplayName } catch { "N/A" }) `
                            -Recommendation "Sprawdz szczegoly bledow w CA > Monitoring > Timer Job Status. Zweryfikuj konto uslugowe."
                    } elseif ($lastStatus -eq "Succeeded" -and $history.StartTime -lt (Get-Date).AddHours(-48)) {
                        $statusLevel = "WARNING"
                        Add-Issue -Section "Timer Jobs" -ObjectName $job.DisplayName `
                            -Description "Timer Job '$($job.DisplayName)' nie byl uruchamiany od >48h (ostatnio: $lastRun)" `
                            -Severity "WARNING" -Server $(try { $job.Server.DisplayName } catch { "N/A" }) `
                            -Recommendation "Sprawdz schedule Timer Joba i stan uslugi SharePoint Timer Service."
                    }
                } else {
                    $lastRun    = "Brak historii"
                    $lastStatus = "Brak historii"
                }
            } catch { }

            $results += [PSCustomObject]@{
                JobName     = $job.DisplayName
                Status      = $lastStatus
                LastRunTime = $lastRun
                Duration_s  = $duration
                Schedule    = $(try { $job.Schedule.ToString() } catch { "N/A" })
                Server      = $(try { $job.Server.DisplayName } catch { "Wszystkie" })
                StatusLevel = $statusLevel
            }
        }
        Write-Log "Pobrano dane $($results.Count) Timer Jobs."
    }
    catch {
        Write-Log "Blad pobierania Timer Jobs: $($_.Exception.Message)" -Level ERROR
        $results += [PSCustomObject]@{ JobName="BLAD"; Status="Blad pobierania";
            LastRunTime="N/A"; Duration_s=0; Schedule="N/A"; Server="N/A"; StatusLevel="CRITICAL" }
    }
    return $results
}

#endregion

#region ─── SEKCJA 7: HEALTH ANALYZER ──────────────────────────────────────────

function Get-HealthAnalyzerRules {
    <#
    .SYNOPSIS
        Pobiera wyniki reguł Health Analyzer SharePoint.
        Zwraca tylko reguły z problemami (Failure lub Warning).
    .OUTPUTS
        Tablica PSCustomObject: RuleName, Category, Severity, Server,
        FailureImpact, Remedy, StatusLevel
    #>
    Write-Log "Pobieranie wynikow Health Analyzer..."
    $results = @()

    if (-not $Script:SPLoaded) {
        return @([PSCustomObject]@{ RuleName="Brak danych"; Category="N/A"; Severity="N/A";
            Server="N/A"; FailureImpact="N/A"; Remedy="N/A"; StatusLevel="WARNING" })
    }

    try {
        # Pobierz wyniki z Health Analyzer (SPHealthAnalyzerReportList)
        $caWA = Get-SPWebApplication -IncludeCentralAdministration -ErrorAction Stop |
                Where-Object { $_.IsAdministrationWebApplication } |
                Select-Object -First 1
        if ($caWA) {
            $caSite = Get-SPSite $caWA.Url -ErrorAction Stop
            $caWeb  = $caSite.RootWeb

            # Szukaj listy Health Analyzer wieloma metodami (różne wersje SP mogą różnie tworzyć listę)
            $reportList = $null
            # Metoda 1: po typie szablonu listy
            try {
                $htId = [int][Microsoft.SharePoint.SPListTemplateType]::HealthReports
                $reportList = $caWeb.Lists | Where-Object { $_.BaseTemplate -eq $htId } | Select-Object -First 1
            } catch { }
            # Metoda 2: po URL listy
            if (-not $reportList) {
                $reportList = try {
                    $caWeb.GetList("$($caWeb.ServerRelativeUrl.TrimEnd('/'))/Lists/HealthReports")
                } catch { $null }
            }
            # Metoda 3: po nazwie/tytule/folderze
            if (-not $reportList) {
                $reportList = $caWeb.Lists | Where-Object {
                    $_.RootFolder.Url -like "*HealthReports*" -or
                    $_.Title -in @("HealthReports","Health Analyzer Results","Health Reports",
                                   "Wyniki analizatora kondycji","Analizator kondycji")
                } | Select-Object -First 1
            }

            if ($reportList) {
                $query = New-Object Microsoft.SharePoint.SPQuery
                $query.RowLimit = 500
                $items = $reportList.GetItems($query)
                foreach ($item in $items) {
                    # Pobierz severity - SPSeverity: 0=Passed, 1=Failed/Error, 2=Warning
                    $sev = try {
                        $sevVal = $item["HealthRuleReportSeverity"]
                        if ($null -ne $sevVal) { $sevVal.ToString().Trim() } else { "0" }
                    } catch { "0" }
                    # Poprawne mapowanie severity (1=Blad/Critical, 2=Ostrzezenie/Warning)
                    $level = switch -Regex ($sev) {
                        "^1$|Error|Fail" { "CRITICAL" }
                        "^2$|Warn"       { "WARNING"  }
                        default          { "OK" }
                    }
                    if ($level -in @("WARNING","CRITICAL")) {
                        # Nazwa reguly - próba kilku nazw pola
                        $ruleName = try {
                            $n = $item["HealthRuleTitle"]
                            if ($null -ne $n -and $n.ToString().Trim() -ne "") { $n.ToString() }
                            else { $item["Title"].ToString() }
                        } catch { try { $item["Title"].ToString() } catch { "Nieznana regula" } }

                        $remedy   = try { $item["HealthRuleRemediation"].ToString() } catch {
                            try { $item["Remedy"].ToString() } catch { "Brak opisu" } }
                        $server   = try { $item["HealthRuleServer"].ToString() } catch { "N/A" }
                        $category = try { $item["HealthRuleCategory"].ToString() } catch { "N/A" }

                        Add-Issue -Section "Health Analyzer" -ObjectName $ruleName `
                            -Description "Health Analyzer: $ruleName (serwer: $server)" `
                            -Severity $level -Server $server `
                            -Recommendation $remedy

                        $results += [PSCustomObject]@{
                            RuleName      = $ruleName
                            Category      = $category
                            Severity      = switch ($sev) {
                                { $_ -match "^1$|Error|Fail" } { "Blad" }
                                { $_ -match "^2$|Warn" }       { "Ostrzezenie" }
                                default { "Info" }
                            }
                            Server        = $server
                            FailureImpact = $(try { $item["HealthRuleImpact"].ToString() } catch { "N/A" })
                            Remedy        = $remedy
                            StatusLevel   = $level
                        }
                    }
                }
            } else {
                Write-Log "Nie znaleziono listy Health Analyzer Reports w CA." -Level WARNING
            }
            $caWeb.Dispose(); $caSite.Dispose()
        }

        if ($results.Count -eq 0) {
            $results += [PSCustomObject]@{ RuleName="Brak problemow"; Category="N/A"; Severity="OK";
                Server="N/A"; FailureImpact="N/A"; Remedy="Brak"; StatusLevel="OK" }
        }
        Write-Log "Health Analyzer: $($results.Count) problemow."
    }
    catch {
        Write-Log "Blad pobierania Health Analyzer: $($_.Exception.Message)" -Level ERROR
        $results += [PSCustomObject]@{ RuleName="BLAD"; Category="Blad"; Severity="N/A";
            Server="N/A"; FailureImpact="N/A"; Remedy="Blad pobierania danych"; StatusLevel="WARNING" }
    }
    return $results
}

#endregion

#region ─── SEKCJA 8: WINDOWS EVENT LOG ─────────────────────────────────────────

function Get-EventLogErrors {
    <#
    .SYNOPSIS
        Pobiera błędy i ostrzeżenia z dzienników zdarzeń Windows (Application, System)
        z ostatnich N godzin z wszystkich serwerów farmy.
    .OUTPUTS
        Tablica PSCustomObject: ServerName, LogName, EventID, Source, Level,
        TimeCreated, Message, StatusLevel
    #>
    Write-Log "Pobieranie bledow Event Log (ostatnie $LogHours h)..."
    $results  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $cutoff   = (Get-Date).AddHours(-$LogHours)

    $serversToCheck = @($env:COMPUTERNAME)
    if ($Script:SPLoaded) {
        try {
            $serversToCheck = Get-SPServer -ErrorAction Stop |
                Where-Object { $_.Role -ne "Invalid" } |
                Select-Object -ExpandProperty DisplayName
        } catch { }
    }

    foreach ($srvName in $serversToCheck) {
        $isLocal = ($srvName -ieq $env:COMPUTERNAME)
        $sb = {
            param($cutoffDT, $srvCtx)
            $evts = @()
            foreach ($logName in @("Application","System")) {
                try {
                    $filter = @{
                        LogName   = $logName
                        Level     = @(1,2,3)   # Critical=1, Error=2, Warning=3
                        StartTime = $cutoffDT
                    }
                    $events = Get-WinEvent -FilterHashtable $filter -ErrorAction Stop |
                              Select-Object -First 100
                    foreach ($e in $events) {
                        $msg = $e.Message
                        if ($msg -and $msg.Length -gt 300) { $msg = $msg.Substring(0,300) + "..." }
                        $evts += [PSCustomObject]@{
                            ServerName  = $srvCtx
                            LogName     = $logName
                            EventID     = $e.Id
                            Source      = $e.ProviderName
                            Level       = $e.LevelDisplayName
                            TimeCreated = $e.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                            Message     = $msg
                            StatusLevel = if ($e.Level -le 2) { "CRITICAL" } else { "WARNING" }
                        }
                    }
                } catch { }
            }
            return $evts
        }

        try {
            $data = if ($isLocal) {
                & $sb -cutoffDT $cutoff -srvCtx $srvName
            } else {
                Invoke-Command -ComputerName $srvName -ScriptBlock $sb `
                    -ArgumentList $cutoff, $srvName -ErrorAction Stop
            }
            foreach ($d in $data) { $results.Add($d) }
        }
        catch {
            Write-Log "Blad Event Log z $srvName : $($_.Exception.Message)" -Level WARNING
            $results.Add([PSCustomObject]@{ ServerName=$srvName; LogName="N/A"; EventID=0;
                Source="Blad"; Level="Error"; TimeCreated=(Get-Date -Format "yyyy-MM-dd HH:mm:ss");
                Message="Nie mozna pobrac Event Log z $srvName : $($_.Exception.Message)"; StatusLevel="WARNING" })
        }
    }

    $critCount = @($results | Where-Object { $_.StatusLevel -eq "CRITICAL" }).Count
    $warnCount = @($results | Where-Object { $_.StatusLevel -eq "WARNING" }).Count
    if ($critCount -gt 0) {
        Add-Issue -Section "Event Log" -ObjectName "Windows Event Log" `
            -Description "Wykryto $critCount krytycznych bledow w Event Log (ostatnie $LogHours h)" `
            -Severity "CRITICAL" -Server "Farma" `
            -Recommendation "Przejrzyj sekcje Event Log w raporcie. Sprawdz zrodla bledow i podejmij dziania naprawcze."
    } elseif ($warnCount -gt 20) {
        Add-Issue -Section "Event Log" -ObjectName "Windows Event Log" `
            -Description "Wykryto $warnCount ostrzezen w Event Log (ostatnie $LogHours h)" `
            -Severity "WARNING" -Server "Farma" `
            -Recommendation "Przejrzyj sekcje Event Log. Duza liczba ostrzezen moze wskazywac na problemy konfiguracyjne."
    }

    Write-Log "Event Log: $($results.Count) zdarzen (krytyczne: $critCount, ostrzezenia: $warnCount)."
    return $results.ToArray()
}

#endregion

#region ─── SEKCJA 9: LOGI ULS ──────────────────────────────────────────────────

function Get-ULSErrors {
    <#
    .SYNOPSIS
        Analizuje logi ULS SharePoint z ostatnich N godzin.
        Zwraca wpisy o poziomie Unexpected, Exception, Critical, High.
    .OUTPUTS
        Tablica PSCustomObject: ServerName, Timestamp, Level, Area, Category,
        EventID, Message, StatusLevel
    #>
    Write-Log "Analizowanie logow ULS (ostatnie $LogHours h)..."
    $results  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $cutoff   = (Get-Date).AddHours(-$LogHours)

    $serversToCheck = @($env:COMPUTERNAME)
    if ($Script:SPLoaded) {
        try {
            $serversToCheck = Get-SPServer -ErrorAction Stop |
                Where-Object { $_.Role -ne "Invalid" } |
                Select-Object -ExpandProperty DisplayName
        } catch { }
    }

    # Standardowa ścieżka do logów ULS (fallback gdy SP niedostępny)
    $ulsRelPath = "Microsoft Shared\Web Server Extensions\16\LOGS"

    # Próba pobrania rzeczywistej ścieżki ULS z konfiguracji SharePoint
    $ulsActualPath = try {
        if ($Script:SPLoaded) {
            $diagCfg = Get-SPDiagnosticConfig -ErrorAction Stop
            if ($diagCfg.LogLocation -and (Test-Path $diagCfg.LogLocation)) { $diagCfg.LogLocation }
            else { $null }
        }
    } catch { $null }

    foreach ($srvName in $serversToCheck) {
        $isLocal = ($srvName -ieq $env:COMPUTERNAME)
        $sb = {
            param($cutoffDT, $srvCtx, $ulsRel, $maxEntries, $lookbackHours, $ulsOverridePath)

            $entries = [System.Collections.Generic.List[PSCustomObject]]::new()
            # Poziomy ULS uznawane za Error i Critical (bez High który jest WARNING)
            $errorCritLevels = @("Unexpected","Exception","Critical","Assert","Error")

            try {
                # Ustal ścieżkę ULS: priorytet override (z Get-SPDiagnosticConfig), potem rejestr, potem standardowa
                $ulsPath = $null
                if ($ulsOverridePath -and (Test-Path $ulsOverridePath)) {
                    $ulsPath = $ulsOverridePath
                }
                if (-not $ulsPath) {
                    # Próba odczytu z rejestru SharePoint
                    $ulsPath = try {
                        $regKey = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\16.0\WSS" -ErrorAction Stop
                        $logDir = $regKey.LogDir
                        if ($logDir -and (Test-Path $logDir)) { $logDir } else { $null }
                    } catch { $null }
                }
                if (-not $ulsPath) {
                    $ulsPath = Join-Path $env:CommonProgramFiles $ulsRel
                }

                if (-not (Test-Path $ulsPath)) {
                    $entries.Add([PSCustomObject]@{
                        ServerName="$srvCtx"; Timestamp="N/A"; Level="Warning";
                        Area="ULS"; Category="Dostep"; EventID="N/A";
                        Message="Katalog ULS nie istnieje: $ulsPath"; StatusLevel="WARNING"
                    })
                    return ,$entries.ToArray()
                }

                $ulsFiles = Get-ChildItem -Path $ulsPath -Filter "*.log" -ErrorAction Stop |
                    Where-Object { $_.LastWriteTime -ge $cutoffDT } |
                    Sort-Object LastWriteTime -Descending

                # Jeśli brak nowych plików — weź ostatni dostępny
                if (-not $ulsFiles) {
                    $ulsFiles = Get-ChildItem -Path $ulsPath -Filter "*.log" -ErrorAction Stop |
                        Sort-Object LastWriteTime -Descending |
                        Select-Object -First 2
                }

                foreach ($f in $ulsFiles) {
                    if ($entries.Count -ge $maxEntries) { break }
                    try {
                        # Otwieramy plik z dostępem współdzielonym (SP trzyma lock na aktywnym logu)
                        $stream = [System.IO.File]::Open(
                            $f.FullName,
                            [System.IO.FileMode]::Open,
                            [System.IO.FileAccess]::Read,
                            [System.IO.FileShare]::ReadWrite
                        )
                        $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8)
                        $fileLines = [System.Collections.Generic.List[string]]::new()
                        while (-not $reader.EndOfStream) { $fileLines.Add($reader.ReadLine()) }
                        $reader.Close(); $stream.Dispose()

                        [Array]$lines = $fileLines.ToArray()
                        [Array]::Reverse($lines)   # czytamy od najnowszych

                        foreach ($line in $lines) {
                            if ($entries.Count -ge $maxEntries) { break }
                            if ([string]::IsNullOrWhiteSpace($line) -or $line.Length -lt 50) { continue }

                            $parts = @($line -split '\t', 8)
                            if ($parts.Count -lt 7) { continue }

                            # Parsowanie czasu ULS (format: MM/dd/yyyy HH:mm:ss.ff — zawsze US)
                            $tsStr = $parts[0].Trim()
                            $ts    = [datetime]::MinValue
                            $parsed = [datetime]::TryParseExact(
                                $tsStr,
                                [string[]]@(
                                    "MM/dd/yyyy HH:mm:ss.ff",
                                    "MM/dd/yyyy HH:mm:ss.fff",
                                    "MM/dd/yyyy HH:mm:ss",
                                    "M/d/yyyy H:mm:ss.ff",
                                    "M/d/yyyy H:mm:ss.fff",
                                    "M/d/yyyy H:mm:ss"
                                ),
                                [System.Globalization.CultureInfo]::InvariantCulture,
                                [System.Globalization.DateTimeStyles]::None,
                                [ref]$ts
                            )
                            if (-not $parsed) {
                                # Fallback: ogólne parsowanie
                                $parsed = [datetime]::TryParse($tsStr, [ref]$ts)
                            }
                            if (-not $parsed -or $ts -lt $cutoffDT) { continue }

                            $level = $parts[6].Trim()

                            # Filtruj — pokazuj tylko Error i Critical
                            if ($level -notin $errorCritLevels) { continue }

                            $msg = if ($parts.Count -ge 8) { $parts[7].Trim() } else { "N/A" }
                            if ($msg.Length -gt 400) { $msg = $msg.Substring(0,400) + "..." }

                            $entries.Add([PSCustomObject]@{
                                ServerName  = $srvCtx
                                Timestamp   = $ts.ToString("yyyy-MM-dd HH:mm:ss")
                                Level       = $level
                                Area        = if ($parts.Count -ge 4) { $parts[3].Trim() } else { "N/A" }
                                Category    = if ($parts.Count -ge 5) { $parts[4].Trim() } else { "N/A" }
                                EventID     = if ($parts.Count -ge 6) { $parts[5].Trim() } else { "N/A" }
                                Message     = $msg
                                StatusLevel = "CRITICAL"
                            })
                        }
                    } catch { }
                }

                if ($entries.Count -eq 0) {
                    $entries.Add([PSCustomObject]@{
                        ServerName="$srvCtx"; Timestamp=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss");
                        Level="Info"; Area="ULS"; Category="Informacja"; EventID="N/A";
                        Message="Brak wpisow Error/Critical w logach ULS z ostatnich $lookbackHours godzin (sciezka: $ulsPath).";
                        StatusLevel="OK"
                    })
                }
            } catch {
                $entries.Add([PSCustomObject]@{
                    ServerName="$srvCtx"; Timestamp=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss");
                    Level="Error"; Area="ULS"; Category="Dostep"; EventID="N/A";
                    Message="Blad odczytu logow ULS: $($_.Exception.Message)"; StatusLevel="WARNING"
                })
            }

            return ,$entries.ToArray()
        }

        # Przekaż ścieżkę ULS: dla lokalnego serwera użyj override, dla zdalnych przekaż ścieżkę relative
        $overridePath = if ($isLocal -and $ulsActualPath) { $ulsActualPath } else { $null }

        try {
            $data = if ($isLocal) {
                & $sb -cutoffDT $cutoff -srvCtx $srvName -ulsRel $ulsRelPath `
                    -maxEntries $Script:ULSMaxEntries -lookbackHours $LogHours -ulsOverridePath $overridePath
            } else {
                Invoke-Command -ComputerName $srvName -ScriptBlock $sb `
                    -ArgumentList $cutoff, $srvName, $ulsRelPath, $Script:ULSMaxEntries, $LogHours, $null -ErrorAction Stop
            }
            foreach ($d in $data) { $results.Add($d) }
        }
        catch {
            Write-Log "Blad ULS z $srvName : $($_.Exception.Message)" -Level WARNING
            $results.Add([PSCustomObject]@{ ServerName=$srvName; Timestamp=(Get-Date -Format "yyyy-MM-dd HH:mm:ss");
                Level="Error"; Area="N/A"; Category="N/A"; EventID="N/A";
                Message="Nie mozna pobrac logow ULS z $srvName : $($_.Exception.Message)"; StatusLevel="WARNING" })
        }
    }

    if ($results.Count -eq 0) {
        $results.Add([PSCustomObject]@{
            ServerName="Farma"; Timestamp=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss");
            Level="Info"; Area="ULS"; Category="Informacja"; EventID="N/A";
            Message="Brak wpisow Error/Critical w logach ULS z ostatnich $LogHours godzin."; StatusLevel="OK"
        })
    }

    $critCount = @($results | Where-Object { $_.StatusLevel -eq "CRITICAL" }).Count
    if ($critCount -gt 0) {
        Add-Issue -Section "ULS Logs" -ObjectName "ULS" `
            -Description "Wykryto $critCount krytycznych wpisow w logach ULS (ostatnie $LogHours h)" `
            -Severity "CRITICAL" -Server "Farma" `
            -Recommendation "Przejrzyj sekcje ULS w raporcie. Skopiuj Correlation ID i przeszukaj logi poleceniem Merge-SPLogFile."
    }

    Write-Log "ULS: $($results.Count) wpisow w sekcji ULS."
    return ,$results.ToArray()
}

#endregion

#region ─── SEKCJA 10: CERTYFIKATY IIS ──────────────────────────────────────────

function Get-CertificateStatus {
    <#
    .SYNOPSIS
        Sprawdza certyfikaty SSL/TLS przypisane do wiązań IIS SharePoint.
        Wykrywa certyfikaty wygasłe lub bliskie wygaśnięcia (≤30 i ≤14 dni).
    .OUTPUTS
        Tablica PSCustomObject: ServerName, SiteName, Binding, Subject, Thumbprint,
        ExpiryDate, DaysLeft, IsExpired, StatusLevel
    #>
    Write-Log "Sprawdzanie certyfikatow IIS..."
    $results  = [System.Collections.Generic.List[PSCustomObject]]::new()

    $serversToCheck = @($env:COMPUTERNAME)
    if ($Script:SPLoaded) {
        try {
            $serversToCheck = Get-SPServer -ErrorAction Stop |
                Where-Object { $_.Role -ne "Invalid" } |
                Select-Object -ExpandProperty DisplayName
        } catch { }
    }

    foreach ($srvName in $serversToCheck) {
        $isLocal = ($srvName -ieq $env:COMPUTERNAME)
        $sb = {
            param($srvCtx)
            $r = @()
            try {
                Import-Module WebAdministration -ErrorAction Stop
                $sites = Get-WebSite -ErrorAction Stop
                foreach ($site in $sites) {
                    $bindings = $site.Bindings.Collection | Where-Object { $_.Protocol -eq "https" }
                    foreach ($binding in $bindings) {
                        $thumbprint = $binding.CertificateHash
                        if (-not $thumbprint) { continue }
                        $certPath = "Cert:\LocalMachine\My\$thumbprint"
                        $cert = try { Get-Item $certPath -ErrorAction Stop } catch { $null }
                        if ($cert) {
                            $daysLeft = [Math]::Round(($cert.NotAfter - (Get-Date)).TotalDays, 0)
                            $r += [PSCustomObject]@{
                                ServerName  = $srvCtx
                                SiteName    = $site.Name
                                Binding     = $binding.BindingInformation
                                Subject     = $cert.Subject
                                Thumbprint  = $thumbprint
                                ExpiryDate  = $cert.NotAfter.ToString("yyyy-MM-dd")
                                DaysLeft    = $daysLeft
                                IsExpired   = ($daysLeft -le 0)
                                StatusLevel = if ($daysLeft -le 0) { "CRITICAL" } elseif ($daysLeft -le 14) { "CRITICAL" } elseif ($daysLeft -le 30) { "WARNING" } else { "OK" }
                            }
                        } else {
                            $r += [PSCustomObject]@{
                                ServerName="$srvCtx"; SiteName=$site.Name;
                                Binding=$binding.BindingInformation; Subject="Nie znaleziono certyfikatu";
                                Thumbprint=$thumbprint; ExpiryDate="N/A"; DaysLeft=-1; IsExpired=$true; StatusLevel="CRITICAL"
                            }
                        }
                    }
                }
            } catch {
                $r += [PSCustomObject]@{
                    ServerName="$srvCtx"; SiteName="BLAD"; Binding="N/A";
                    Subject="Blad: $($_.Exception.Message)"; Thumbprint="N/A";
                    ExpiryDate="N/A"; DaysLeft=-1; IsExpired=$false; StatusLevel="WARNING"
                }
            }
            return $r
        }

        try {
            $data = if ($isLocal) {
                & $sb -srvCtx $srvName
            } else {
                Invoke-Command -ComputerName $srvName -ScriptBlock $sb -ArgumentList $srvName -ErrorAction Stop
            }
            foreach ($d in $data) { $results.Add($d) }
        }
        catch {
            Write-Log "Blad certyfikatow z $srvName : $($_.Exception.Message)" -Level WARNING
        }
    }

    # Zgłoś problemy certyfikatów
    foreach ($c in $results) {
        if ($c.StatusLevel -eq "CRITICAL") {
            $desc = if ($c.DaysLeft -le 0) {
                "Certyfikat wygasl! ($($c.Subject)) na witrynie '$($c.SiteName)'"
            } else {
                "Certyfikat wygasa za $($c.DaysLeft) dni ($($c.Subject)) na witrynie '$($c.SiteName)'"
            }
            Add-Issue -Section "Certyfikaty" -ObjectName $c.SiteName `
                -Description $desc -Severity "CRITICAL" -Server $c.ServerName `
                -Recommendation "Odnow certyfikat SSL przed data $($c.ExpiryDate). Zaktualizuj powiazania IIS po odnowieniu."
        } elseif ($c.StatusLevel -eq "WARNING") {
            Add-Issue -Section "Certyfikaty" -ObjectName $c.SiteName `
                -Description "Certyfikat wygasa za $($c.DaysLeft) dni ($($c.Subject))" `
                -Severity "WARNING" -Server $c.ServerName `
                -Recommendation "Zaplanuj odnowienie certyfikatu. Data wygasniecia: $($c.ExpiryDate)"
        }
    }

    if ($results.Count -eq 0) {
        $results.Add([PSCustomObject]@{ ServerName="N/A"; SiteName="Brak HTTPS"; Binding="N/A";
            Subject="Brak certyfikatow HTTPS w IIS lub brak dostepu"; Thumbprint="N/A";
            ExpiryDate="N/A"; DaysLeft=999; IsExpired=$false; StatusLevel="OK" })
    }

    Write-Log "Certyfikaty: $($results.Count) rekordow."
    return $results.ToArray()
}

#endregion

#region ─── SEKCJA 11: ZASOBY SYSTEMOWE ─────────────────────────────────────────

function Get-SystemResources {
    <#
    .SYNOPSIS
        Pobiera dane o dyskach, pamięci RAM i CPU ze wszystkich serwerów farmy.
    .OUTPUTS
        Hashtable: Disks (tablica) i Memory (tablica)
    #>
    Write-Log "Pobieranie zasobow systemowych..."
    $allDisks  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $allMemory = [System.Collections.Generic.List[PSCustomObject]]::new()

    $serversToCheck = @($env:COMPUTERNAME)
    if ($Script:SPLoaded) {
        try {
            $serversToCheck = Get-SPServer -ErrorAction Stop |
                Where-Object { $_.Role -ne "Invalid" } |
                Select-Object -ExpandProperty DisplayName
        } catch { }
    }

    foreach ($srvName in $serversToCheck) {
        $isLocal = ($srvName -ieq $env:COMPUTERNAME)
        $sb = {
            param($srvCtx, $diskWarnGB, $diskCritGB, $ramWarnPct, $ramCritPct)
            $r = @{ Disks = @(); Memory = @() }
            try {
                # Dyski
                $disks = Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
                foreach ($d in $disks) {
                    $freeGB  = [Math]::Round($d.FreeSpace / 1GB, 2)
                    $totalGB = [Math]::Round($d.Size / 1GB, 2)
                    $usedPct = if ($d.Size -gt 0) { [Math]::Round(100 - ($d.FreeSpace / $d.Size * 100), 1) } else { 0 }
                    $level   = if ($freeGB -le $diskCritGB) { "CRITICAL" } elseif ($freeGB -le $diskWarnGB) { "WARNING" } else { "OK" }
                    $r.Disks += [PSCustomObject]@{
                        ServerName  = $srvCtx
                        Drive       = $d.DeviceID
                        Label       = $d.VolumeName
                        FreeGB      = $freeGB
                        TotalGB     = $totalGB
                        UsedPct     = $usedPct
                        StatusLevel = $level
                    }
                }
                # Pamięć RAM
                $os    = Get-WmiObject Win32_OperatingSystem -ErrorAction Stop
                $cs    = Get-WmiObject Win32_ComputerSystem  -ErrorAction Stop
                $totalRAMGB = [Math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
                $freeRAMGB  = [Math]::Round($os.FreePhysicalMemory  / 1MB, 2)  # FreePhysicalMemory jest w KB
                $freeRAMGB  = [Math]::Round($os.FreePhysicalMemory * 1KB / 1GB, 2)
                $usedRAMPct = if ($totalRAMGB -gt 0) { [Math]::Round(100 - ($freeRAMGB / $totalRAMGB * 100), 1) } else { 0 }
                $ramLevel   = if ($usedRAMPct -ge $ramCritPct) { "CRITICAL" } elseif ($usedRAMPct -ge $ramWarnPct) { "WARNING" } else { "OK" }
                $r.Memory += [PSCustomObject]@{
                    ServerName  = $srvCtx
                    TotalRAM_GB = $totalRAMGB
                    FreeRAM_GB  = $freeRAMGB
                    UsedPct     = $usedRAMPct
                    StatusLevel = $ramLevel
                }
            } catch {
                $r.Disks  += [PSCustomObject]@{ ServerName=$srvCtx; Drive="BLAD"; Label="Blad";
                    FreeGB=0; TotalGB=0; UsedPct=0; StatusLevel="WARNING" }
                $r.Memory += [PSCustomObject]@{ ServerName=$srvCtx; TotalRAM_GB=0; FreeRAM_GB=0;
                    UsedPct=0; StatusLevel="WARNING" }
            }
            return $r
        }

        try {
            $data = if ($isLocal) {
                & $sb -srvCtx $srvName -diskWarnGB $DiskWarningThresholdGB `
                    -diskCritGB $DiskCriticalThresholdGB -ramWarnPct $RAMWarningThreshold `
                    -ramCritPct $RAMCriticalThreshold
            } else {
                Invoke-Command -ComputerName $srvName -ScriptBlock $sb `
                    -ArgumentList $srvName,$DiskWarningThresholdGB,$DiskCriticalThresholdGB,`
                                  $RAMWarningThreshold,$RAMCriticalThreshold -ErrorAction Stop
            }
            foreach ($d in $data.Disks)  { $allDisks.Add($d) }
            foreach ($m in $data.Memory) { $allMemory.Add($m) }
        }
        catch {
            Write-Log "Blad zasobow systemowych z $srvName : $($_.Exception.Message)" -Level WARNING
            $allDisks.Add([PSCustomObject]@{ ServerName=$srvName; Drive="Blad"; Label="Blad";
                FreeGB=0; TotalGB=0; UsedPct=0; StatusLevel="WARNING" })
            $allMemory.Add([PSCustomObject]@{ ServerName=$srvName; TotalRAM_GB=0; FreeRAM_GB=0;
                UsedPct=0; StatusLevel="WARNING" })
        }
    }

    # Zgłoś problemy dyskowe i pamięci
    foreach ($d in $allDisks) {
        if ($d.StatusLevel -eq "CRITICAL") {
            Add-Issue -Section "Zasoby systemowe" -ObjectName "$($d.ServerName):$($d.Drive)" `
                -Description "Krytycznie malo miejsca na dysku $($d.Drive) na '$($d.ServerName)': $($d.FreeGB) GB wolne" `
                -Severity "CRITICAL" -Server $d.ServerName `
                -Recommendation "Natychmiast zwolnij miejsce na dysku lub rozszerz wolumin. Sprawdz logi IIS, ULS, bazy SQL."
        } elseif ($d.StatusLevel -eq "WARNING") {
            Add-Issue -Section "Zasoby systemowe" -ObjectName "$($d.ServerName):$($d.Drive)" `
                -Description "Malo miejsca na dysku $($d.Drive) na '$($d.ServerName)': $($d.FreeGB) GB wolne" `
                -Severity "WARNING" -Server $d.ServerName `
                -Recommendation "Zaplanuj zwolnienie miejsca lub rozszerzenie woluminu."
        }
    }
    foreach ($m in $allMemory) {
        if ($m.StatusLevel -eq "CRITICAL") {
            Add-Issue -Section "Zasoby systemowe" -ObjectName "$($m.ServerName) RAM" `
                -Description "Krytyczne uzycie pamieci RAM na '$($m.ServerName)': $($m.UsedPct)%" `
                -Severity "CRITICAL" -Server $m.ServerName `
                -Recommendation "Sprawdz procesy zuzyajace RAM (Task Manager, Get-Process). Rozwaz dodanie pamieci."
        } elseif ($m.StatusLevel -eq "WARNING") {
            Add-Issue -Section "Zasoby systemowe" -ObjectName "$($m.ServerName) RAM" `
                -Description "Wysokie uzycie pamieci RAM na '$($m.ServerName)': $($m.UsedPct)%" `
                -Severity "WARNING" -Server $m.ServerName `
                -Recommendation "Monitoruj wykorzystanie RAM. Sprawdz mozliwosc dodania pamieci fizycznej."
        }
    }

    Write-Log "Zasoby: $($allDisks.Count) dyskow, $($allMemory.Count) serwerow RAM."
    return @{ Disks = $allDisks.ToArray(); Memory = $allMemory.ToArray() }
}

#endregion

#region ─── SEKCJA 12: SPÓJNOŚĆ WERSJI FARMY ────────────────────────────────────

function Get-FarmVersionConsistency {
    <#
    .SYNOPSIS
        Sprawdza spójność wersji SharePoint na wszystkich serwerach farmy.
        Wykrywa rozbieżności wersji build, które mogą powodować problemy.
    .OUTPUTS
        Tablica PSCustomObject: ServerName, BuildVersion, PatchLevel, IsConsistent,
        StatusLevel
    #>
    Write-Log "Sprawdzanie spojnosci wersji farmy..."
    $results = @()

    if (-not $Script:SPLoaded) {
        return @([PSCustomObject]@{ ServerName="N/A"; BuildVersion="Brak danych";
            PatchLevel="N/A"; IsConsistent="N/A"; StatusLevel="WARNING" })
    }

    try {
        $servers = Get-SPServer -ErrorAction Stop | Where-Object { $_.Role -ne "Invalid" }
        $versions = @{}

        foreach ($srv in $servers) {
            # Próba z SPServer.BuildVersion, fallback do rejestru Windows
            $build = try {
                $bv = $srv.BuildVersion
                if ($null -ne $bv -and $bv.ToString() -notin @("","0.0.0.0")) { $bv.ToString() } else { $null }
            } catch { $null }
            if (-not $build) {
                $regGetBuild3 = {
                    $regPath = "HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\16.0\WSS"
                    try { (Get-ItemProperty -Path $regPath -Name "Build" -ErrorAction Stop).Build } catch { "Nieznana" }
                }
                $build = try {
                    if ($srv.DisplayName -ieq $env:COMPUTERNAME) { & $regGetBuild3 }
                    else { Invoke-Command -ComputerName $srv.DisplayName -ScriptBlock $regGetBuild3 -ErrorAction Stop }
                } catch { "Nieznana" }
                if (-not $build) { $build = "Nieznana" }
            }
            $versions[$srv.DisplayName] = $build
            $results += [PSCustomObject]@{
                ServerName   = $srv.DisplayName
                BuildVersion = $build
                PatchLevel   = $(try { $srv.ProductVersions | Select-Object -ExpandProperty PatchLevel -First 1 } catch { "N/A" })
                IsConsistent = "Sprawdzam..."
                StatusLevel  = "OK"
            }
        }

        # Sprawdź spójność
        $uniqueVersions = @($versions.Values | Select-Object -Unique)
        if ($uniqueVersions.Count -gt 1) {
            $versionList = $uniqueVersions -join ", "
            foreach ($r in $results) {
                $r.IsConsistent = "NIE - Roznice: $versionList"
                $r.StatusLevel  = "WARNING"
            }
            Add-Issue -Section "Wersje farmy" -ObjectName "Spójność wersji" `
                -Description "Rozne wersje build SharePoint na serwerach farmy: $versionList" `
                -Severity "WARNING" -Server "Farma" `
                -Recommendation "Zainstaluj brakujace aktualizacje na serwerach z nizsza wersja. Uruchom psconfig po aktualizacji."
        } else {
            foreach ($r in $results) {
                $r.IsConsistent = "TAK"
            }
        }

        Write-Log "Wersje: $($uniqueVersions.Count) unikalnych wersji na $($results.Count) serwerach."
    }
    catch {
        Write-Log "Blad sprawdzania wersji farmy: $($_.Exception.Message)" -Level ERROR
        $results += [PSCustomObject]@{ ServerName="BLAD"; BuildVersion="Blad";
            PatchLevel="N/A"; IsConsistent="N/A"; StatusLevel="WARNING" }
    }
    return $results
}

#endregion

#region ─── SEKCJA 13: SERVICE APPLICATIONS ─────────────────────────────────────

function Get-ServiceApplications {
    <#
    .SYNOPSIS
        Pobiera stan kluczowych Service Applications SharePoint.
    .OUTPUTS
        Tablica PSCustomObject: Name, Type, Status, StatusLevel
    #>
    Write-Log "Pobieranie Service Applications..."
    $results = @()

    if (-not $Script:SPLoaded) {
        return @([PSCustomObject]@{ Name="Brak danych"; Type="N/A"; Status="Niedostepne"; StatusLevel="WARNING" })
    }

    try {
        $svcApps = Get-SPServiceApplication -ErrorAction Stop
        foreach ($sa in $svcApps) {
            $status = try { $sa.Status.ToString() } catch { "Nieznany" }
            $level  = if ($status -eq "Online") { "OK" } elseif ($status -eq "Disabled") { "WARNING" } else { "CRITICAL" }
            if ($level -ne "OK") {
                Add-Issue -Section "Service Applications" -ObjectName $sa.Name `
                    -Description "Service Application '$($sa.Name)' ma stan: $status" `
                    -Severity $level -Server "Farma" `
                    -Recommendation "Sprawdz stan Service Application w CA > Application Management > Manage service applications."
            }
            $results += [PSCustomObject]@{
                Name        = $sa.Name
                Type        = $(try { $sa.TypeName } catch { $sa.GetType().Name })
                Status      = $status
                StatusLevel = $level
            }
        }
        Write-Log "Service Applications: $($results.Count) rekordow."
    }
    catch {
        Write-Log "Blad pobierania Service Applications: $($_.Exception.Message)" -Level ERROR
        $results += [PSCustomObject]@{ Name="BLAD"; Type="Blad"; Status="Blad pobierania"; StatusLevel="CRITICAL" }
    }
    return $results
}

#endregion

#region ─── BUDOWANIE RAPORTU HTML ──────────────────────────────────────────────

function Get-StatusBadge {
    param([string]$Level, [string]$Text = "")
    $label = if ($Text) { $Text } else { $Level }
    $class = switch ($Level) {
        "OK"       { "badge-ok" }
        "WARNING"  { "badge-warn" }
        "CRITICAL" { "badge-crit" }
        default    { "badge-info" }
    }
    return "<span class='badge $class'>$label</span>"
}

function Get-StatusIcon {
    param([string]$Level)
    return $(switch ($Level) {
        "OK"       { "<span class='icon-ok'>&#10003;</span>" }
        "WARNING"  { "<span class='icon-warn'>&#9888;</span>" }
        "CRITICAL" { "<span class='icon-crit'>&#10007;</span>" }
        default    { "<span class='icon-info'>&#8505;</span>" }
    })
}

function ConvertTo-HtmlTable {
    <#
    .SYNOPSIS
        Generuje tabelę HTML z tablicy obiektów PowerShell.
        Kolumna 'StatusLevel' jest automatycznie zamieniana na kolorową odznakę.
    #>
    param(
        [object[]]$Data,
        [string]$TableId,
        [string[]]$Columns,       # Lista kolumn do pokazania (opcjonalna)
        [string[]]$ColumnHeaders  # Opcjonalne nagłówki kolumn
    )
    if (-not $Data -or $Data.Count -eq 0) {
        return "<p class='no-data'>Brak danych do wyswietlenia.</p>"
    }

    $first = $Data[0]
    if (-not $Columns) {
        $Columns = $first.PSObject.Properties.Name
    }
    if (-not $ColumnHeaders) {
        $ColumnHeaders = $Columns
    }

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("<div class='table-wrapper'>")
    [void]$sb.AppendLine("<table id='$TableId' class='data-table'>")
    [void]$sb.AppendLine("<thead><tr>")
    for ($i = 0; $i -lt $Columns.Count; $i++) {
        if ($Columns[$i] -eq "StatusLevel") {
            [void]$sb.AppendLine("<th>$($ColumnHeaders[$i])</th>")
            continue
        }
        [void]$sb.AppendLine("<th>$($ColumnHeaders[$i])</th>")
    }
    [void]$sb.AppendLine("</tr></thead>")
    [void]$sb.AppendLine("<tbody>")

    foreach ($row in $Data) {
        $rowLevel = try { $row.StatusLevel } catch { "OK" }
        $rowClass = switch ($rowLevel) {
            "CRITICAL" { "row-crit" }
            "WARNING"  { "row-warn" }
            default    { "" }
        }
        [void]$sb.Append("<tr class='$rowClass'>")
        foreach ($col in $Columns) {
            if ($col -eq "StatusLevel") { continue }
            $val = try { $row.$col } catch { "N/A" }
            if ($null -eq $val) { $val = "" }
            $valStr = $val.ToString()
            # Zabezpieczenie przed XSS
            $valStr = $valStr.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;")
            # Oznacz status w dedykowanej kolumnie jeśli ona poprzedza
            if ($col -eq ($Columns | Where-Object { $_ -ne "StatusLevel" } | Select-Object -Last 1)) {
                [void]$sb.Append("<td>$valStr</td>")
                [void]$sb.Append("<td>$(Get-StatusBadge -Level $rowLevel)</td>")
            } else {
                [void]$sb.Append("<td>$valStr</td>")
            }
        }
        [void]$sb.AppendLine("</tr>")
    }

    [void]$sb.AppendLine("</tbody>")
    [void]$sb.AppendLine("</table>")
    [void]$sb.AppendLine("</div>")
    return $sb.ToString()
}

function Build-SectionHTML {
    <#
    .SYNOPSIS
        Buduje sekcję HTML z nagłówkiem, licznikami statusów i tabelą danych.
    #>
    param(
        [string]$SectionId,
        [string]$SectionTitle,
        [string]$Icon,
        [string]$TableHTML,
        [int]$OKCount,
        [int]$WarnCount,
        [int]$CritCount,
        [string]$ExtraInfo = ""
    )

    $headerLevel = if ($CritCount -gt 0) { "CRITICAL" } elseif ($WarnCount -gt 0) { "WARNING" } else { "OK" }
    $headerClass = switch ($headerLevel) {
        "CRITICAL" { "section-header-crit" }
        "WARNING"  { "section-header-warn" }
        default    { "section-header-ok" }
    }

    return @"
<div class="report-section" id="section-$SectionId">
  <div class="section-header $headerClass" onclick="toggleSection('$SectionId')">
    <span class="section-icon">$Icon</span>
    <span class="section-title">$SectionTitle</span>
    <div class="section-badges">
      <span class="badge badge-ok">OK: $OKCount</span>
      <span class="badge badge-warn">OSTRZEZENIA: $WarnCount</span>
      <span class="badge badge-crit">KRYTYCZNE: $CritCount</span>
    </div>
    <span class="section-toggle" id="toggle-$SectionId">&#9660;</span>
  </div>
  <div class="section-body" id="body-$SectionId">
    $ExtraInfo
    $TableHTML
  </div>
</div>
"@
}

function Build-HTMLReport {
    <#
    .SYNOPSIS
        Buduje kompletny raport HTML ze wszystkich zebranych danych.
    #>
    param(
        [object[]]$Servers,
        [object[]]$SPServices,
        [hashtable]$IISData,
        [object[]]$WebApps,
        [object[]]$Databases,
        [object[]]$TimerJobs,
        [object[]]$HealthAnalyzer,
        [object[]]$EventLog,
        [object[]]$ULSLog,
        [object[]]$Certificates,
        [hashtable]$SystemResources,
        [object[]]$FarmVersions,
        [object[]]$ServiceApps
    )

    $now       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $farmName  = $Script:FarmName
    $scriptSrv = $Script:ScriptServer
    $duration  = [Math]::Round(((Get-Date) - $Script:StartTime).TotalSeconds, 1)

    function Normalize-DataItems {
        param([object[]]$Data)

        if ($null -eq $Data) { return @() }
        return @($Data | Where-Object { $null -ne $_ })
    }

    $Servers        = Normalize-DataItems -Data $Servers
    $SPServices     = Normalize-DataItems -Data $SPServices
    $WebApps        = Normalize-DataItems -Data $WebApps
    $Databases      = Normalize-DataItems -Data $Databases
    $TimerJobs      = Normalize-DataItems -Data $TimerJobs
    $HealthAnalyzer = Normalize-DataItems -Data $HealthAnalyzer
    $EventLog       = Normalize-DataItems -Data $EventLog
    $ULSLog         = Normalize-DataItems -Data $ULSLog
    $Certificates   = Normalize-DataItems -Data $Certificates
    $FarmVersions   = Normalize-DataItems -Data $FarmVersions
    $ServiceApps    = Normalize-DataItems -Data $ServiceApps

    if (-not $IISData) { $IISData = @{} }
    $IISData = @{
        Sites    = if ($IISData.ContainsKey("Sites")) { Normalize-DataItems -Data $IISData["Sites"] } else { @() }
        AppPools = if ($IISData.ContainsKey("AppPools")) { Normalize-DataItems -Data $IISData["AppPools"] } else { @() }
    }

    if (-not $SystemResources) { $SystemResources = @{} }
    $SystemResources = @{
        Disks  = if ($SystemResources.ContainsKey("Disks")) { Normalize-DataItems -Data $SystemResources["Disks"] } else { @() }
        Memory = if ($SystemResources.ContainsKey("Memory")) { Normalize-DataItems -Data $SystemResources["Memory"] } else { @() }
    }

    # Oblicz globalny status farmy
    $totalCrit = $Script:AllIssues | Where-Object { $_.Severity -eq "CRITICAL" } | Measure-Object | Select-Object -ExpandProperty Count
    $totalWarn = $Script:AllIssues | Where-Object { $_.Severity -eq "WARNING"  } | Measure-Object | Select-Object -ExpandProperty Count
    $globalStatus = if ($totalCrit -gt 0) { "CRITICAL" } elseif ($totalWarn -gt 0) { "WARNING" } else { "OK" }
    $globalClass  = switch ($globalStatus) { "CRITICAL" { "status-critical" } "WARNING" { "status-warning" } default { "status-ok" } }
    $globalLabel  = switch ($globalStatus) { "CRITICAL" { "KRYTYCZNY" } "WARNING" { "OSTRZEZENIE" } default { "ZDROWA" } }

    # ─── CSS ────────────────────────────────────────────────────────────────────
    $css = @'
:root {
  --bg: #0f1117;
  --bg2: #1a1d27;
  --bg3: #23263a;
  --border: #2e3150;
  --text: #d0d4e8;
  --text-dim: #8892b0;
  --ok: #2ecc71;
  --warn: #f39c12;
  --crit: #e74c3c;
  --info: #3498db;
  --accent: #5865f2;
  --radius: 8px;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); font-size: 13px; }

/* ── HEADER ── */
.header { background: linear-gradient(135deg,#1a1d35 0%,#2d3170 50%,#1a1d35 100%);
  padding: 24px 32px; border-bottom: 2px solid var(--accent); }
.header h1 { font-size: 22px; color: #fff; margin-bottom: 6px; letter-spacing: 0.5px; }
.header-meta { color: var(--text-dim); font-size: 12px; display: flex; gap: 24px; flex-wrap: wrap; margin-top: 8px; }
.header-meta span { display: flex; align-items: center; gap: 5px; }

/* ── GLOBAL STATUS BANNER ── */
.status-banner { padding: 14px 32px; display: flex; align-items: center; gap: 16px;
  border-bottom: 1px solid var(--border); }
.status-ok       { background: rgba(46,204,113,.12); }
.status-warning  { background: rgba(243,156,18,.12); }
.status-critical { background: rgba(231,76,60,.15); }
.status-dot { width: 14px; height: 14px; border-radius: 50%; flex-shrink: 0; }
.status-ok       .status-dot { background: var(--ok); box-shadow: 0 0 8px var(--ok); }
.status-warning  .status-dot { background: var(--warn); box-shadow: 0 0 8px var(--warn); }
.status-critical .status-dot { background: var(--crit); box-shadow: 0 0 8px var(--crit); animation: pulse 1.5s infinite; }
@keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.5} }
.status-text { font-size: 15px; font-weight: 600; }
.status-ok       .status-text { color: var(--ok); }
.status-warning  .status-text { color: var(--warn); }
.status-critical .status-text { color: var(--crit); }
.status-counts { margin-left: auto; display: flex; gap: 12px; }

/* ── LAYOUT ── */
.container { display: flex; min-height: calc(100vh - 130px); }
.sidebar { width: 240px; background: var(--bg2); border-right: 1px solid var(--border);
  padding: 16px 0; position: sticky; top: 0; height: 100vh; overflow-y: auto; flex-shrink: 0; }
.sidebar-title { color: var(--text-dim); font-size: 10px; text-transform: uppercase;
  letter-spacing: 1px; padding: 8px 16px; margin-bottom: 4px; }
.nav-item { display: block; padding: 9px 16px 9px 20px; color: var(--text-dim);
  text-decoration: none; font-size: 12px; border-left: 3px solid transparent;
  transition: all .15s; cursor: pointer; }
.nav-item:hover { color: var(--text); background: var(--bg3); }
.nav-item.active { color: var(--accent); border-left-color: var(--accent); background: rgba(88,101,242,.1); }
.nav-item .nav-badge { float: right; }
.main-content { flex: 1; padding: 20px 28px; overflow-x: auto; }

/* ── SEARCH / FILTER BAR ── */
.toolbar { display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; align-items: center; }
.search-box { padding: 8px 14px; background: var(--bg2); border: 1px solid var(--border);
  border-radius: var(--radius); color: var(--text); font-size: 12px; width: 280px; outline: none; }
.search-box:focus { border-color: var(--accent); }
.filter-select { padding: 8px 12px; background: var(--bg2); border: 1px solid var(--border);
  border-radius: var(--radius); color: var(--text); font-size: 12px; outline: none; cursor: pointer; }
.btn { padding: 7px 16px; border-radius: var(--radius); border: none; font-size: 12px;
  cursor: pointer; font-weight: 500; transition: opacity .15s; }
.btn-primary { background: var(--accent); color: #fff; }
.btn-secondary { background: var(--bg3); color: var(--text); border: 1px solid var(--border); }
.btn:hover { opacity: .85; }

/* ── SUMMARY CARDS ── */
.summary-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(160px,1fr)); gap: 12px; margin-bottom: 24px; }
.stat-card { background: var(--bg2); border: 1px solid var(--border); border-radius: var(--radius);
  padding: 16px; text-align: center; }
.stat-card .stat-num  { font-size: 32px; font-weight: 700; line-height: 1; }
.stat-card .stat-label{ font-size: 11px; color: var(--text-dim); margin-top: 4px; text-transform: uppercase; letter-spacing: .5px; }
.stat-card.card-crit  { border-color: var(--crit); }
.stat-card.card-warn  { border-color: var(--warn); }
.stat-card.card-ok    { border-color: var(--ok); }
.stat-card .stat-num.num-crit { color: var(--crit); }
.stat-card .stat-num.num-warn { color: var(--warn); }
.stat-card .stat-num.num-ok   { color: var(--ok); }
.stat-card .stat-num.num-info { color: var(--info); }

/* ── ISSUES TABLE (Summary) ── */
.issues-table { width: 100%; border-collapse: collapse; margin-bottom: 16px; }
.issues-table th { background: var(--bg3); color: var(--text-dim); font-size: 11px;
  text-transform: uppercase; letter-spacing: .5px; padding: 8px 12px; text-align: left; }
.issues-table td { padding: 8px 12px; border-bottom: 1px solid var(--border); vertical-align: top; }
.issues-table tr:hover td { background: rgba(88,101,242,.05); }

/* ── REPORT SECTIONS ── */
.report-section { background: var(--bg2); border: 1px solid var(--border);
  border-radius: var(--radius); margin-bottom: 16px; overflow: hidden; }
.section-header { display: flex; align-items: center; gap: 10px; padding: 13px 18px;
  cursor: pointer; user-select: none; transition: background .15s; }
.section-header:hover { background: rgba(255,255,255,.04); }
.section-header-ok   { border-left: 4px solid var(--ok); }
.section-header-warn { border-left: 4px solid var(--warn); }
.section-header-crit { border-left: 4px solid var(--crit); }
.section-icon  { font-size: 16px; }
.section-title { font-weight: 600; font-size: 14px; color: #fff; }
.section-badges { display: flex; gap: 6px; margin-left: auto; }
.section-toggle { font-size: 12px; color: var(--text-dim); transition: transform .2s; }
.section-toggle.collapsed { transform: rotate(-90deg); }
.section-body { padding: 16px 18px; display: block; }
.section-body.hidden { display: none; }

/* ── BADGES ── */
.badge { display: inline-block; padding: 2px 9px; border-radius: 12px; font-size: 11px;
  font-weight: 600; white-space: nowrap; }
.badge-ok   { background: rgba(46,204,113,.15); color: var(--ok); border: 1px solid rgba(46,204,113,.3); }
.badge-warn { background: rgba(243,156,18,.15); color: var(--warn); border: 1px solid rgba(243,156,18,.3); }
.badge-crit { background: rgba(231,76,60,.15); color: var(--crit); border: 1px solid rgba(231,76,60,.3); }
.badge-info { background: rgba(52,152,219,.15); color: var(--info); border: 1px solid rgba(52,152,219,.3); }

/* ── ICONS (inline) ── */
.icon-ok   { color: var(--ok); }
.icon-warn { color: var(--warn); }
.icon-crit { color: var(--crit); }

/* ── DATA TABLES ── */
.table-wrapper { overflow-x: auto; }
.data-table { width: 100%; border-collapse: collapse; font-size: 12px; }
.data-table th { background: var(--bg3); color: var(--text-dim); font-size: 11px;
  text-transform: uppercase; letter-spacing: .4px; padding: 9px 12px; text-align: left;
  white-space: nowrap; position: sticky; top: 0; }
.data-table td { padding: 8px 12px; border-bottom: 1px solid var(--border);
  vertical-align: top; word-break: break-word; max-width: 380px; }
.data-table tr:hover td { background: rgba(88,101,242,.04); }
.data-table tr.row-warn td { background: rgba(243,156,18,.04); }
.data-table tr.row-crit td { background: rgba(231,76,60,.06); }
.data-table tr.filtered-out { display: none; }

/* ── MISC ── */
.no-data { color: var(--text-dim); padding: 16px 0; font-style: italic; }
.reco-box { background: rgba(88,101,242,.08); border: 1px solid rgba(88,101,242,.25);
  border-radius: var(--radius); padding: 10px 14px; margin-top: 8px; font-size: 12px;
  color: #a0a8d4; }
.reco-box::before { content: "Rekomendacja: "; font-weight: 600; color: var(--accent); }
.tag { display: inline-block; padding: 1px 7px; border-radius: 4px; font-size: 10px;
  background: var(--bg3); color: var(--text-dim); margin-right: 4px; }
.disk-bar-wrap { background: var(--bg3); border-radius: 4px; height: 8px; min-width: 80px; display: inline-block; vertical-align: middle; width: 90px; }
.disk-bar { height: 8px; border-radius: 4px; }
.disk-bar-ok   { background: var(--ok); }
.disk-bar-warn { background: var(--warn); }
.disk-bar-crit { background: var(--crit); }
.footer { text-align: center; padding: 16px; color: var(--text-dim); font-size: 11px;
  border-top: 1px solid var(--border); }
@media(max-width:900px){ .sidebar{display:none;} }
'@

    # ─── JAVASCRIPT ─────────────────────────────────────────────────────────────
    $js = @'
// Zwijanie/rozwijanie sekcji
function toggleSection(id){
  var body   = document.getElementById('body-'+id);
  var toggle = document.getElementById('toggle-'+id);
  if(!body) return;
  body.classList.toggle('hidden');
  toggle.classList.toggle('collapsed');
}

// Rozwiń wszystkie
function expandAll(){
  document.querySelectorAll('.section-body').forEach(function(b){
    b.classList.remove('hidden');
  });
  document.querySelectorAll('.section-toggle').forEach(function(t){
    t.classList.remove('collapsed');
  });
}
// Zwiń wszystkie
function collapseAll(){
  document.querySelectorAll('.section-body').forEach(function(b){
    b.classList.add('hidden');
  });
  document.querySelectorAll('.section-toggle').forEach(function(t){
    t.classList.add('collapsed');
  });
}

// Globalne wyszukiwanie tekstu w widocznych tabelach
function globalSearch(){
  var q = document.getElementById('globalSearch').value.toLowerCase();
  document.querySelectorAll('.data-table tbody tr').forEach(function(row){
    var txt = row.textContent.toLowerCase();
    row.classList.toggle('filtered-out', q.length > 0 && txt.indexOf(q) === -1);
  });
}

// Filtr po statusie
function filterStatus(){
  var sel = document.getElementById('statusFilter').value;
  document.querySelectorAll('.data-table tbody tr').forEach(function(row){
    if(!sel){ row.classList.remove('filtered-out'); return; }
    var match = (sel==='CRITICAL' && row.classList.contains('row-crit')) ||
                (sel==='WARNING'  && row.classList.contains('row-warn'))  ||
                (sel==='OK'       && !row.classList.contains('row-crit') && !row.classList.contains('row-warn'));
    row.classList.toggle('filtered-out', !match);
  });
}

// Filtr po serwerze
function filterServer(){
  var sel = document.getElementById('serverFilter').value.toLowerCase();
  document.querySelectorAll('.data-table tbody tr').forEach(function(row){
    if(!sel){ row.classList.remove('filtered-out'); return; }
    var txt = row.textContent.toLowerCase();
    row.classList.toggle('filtered-out', txt.indexOf(sel) === -1);
  });
}

// Reset filtrów
function resetFilters(){
  document.getElementById('globalSearch').value  = '';
  document.getElementById('statusFilter').value  = '';
  document.getElementById('serverFilter').value  = '';
  document.querySelectorAll('.data-table tbody tr').forEach(function(r){
    r.classList.remove('filtered-out');
  });
}

// Nawigacja boczna — smooth scroll
document.querySelectorAll('.nav-item').forEach(function(item){
  item.addEventListener('click', function(){
    var target = document.getElementById(item.dataset.target);
    if(target){ target.scrollIntoView({behavior:'smooth', block:'start'}); }
    document.querySelectorAll('.nav-item').forEach(function(i){ i.classList.remove('active'); });
    item.classList.add('active');
  });
});

// Podświetlaj aktywny element nawigacji przy scrollowaniu
window.addEventListener('scroll', function(){
  var sections = document.querySelectorAll('.report-section');
  var mid = window.scrollY + window.innerHeight/2;
  sections.forEach(function(s){
    var navItem = document.querySelector('.nav-item[data-target="'+s.id+'"]');
    if(!navItem) return;
    if(s.offsetTop <= mid && s.offsetTop + s.offsetHeight > mid){
      document.querySelectorAll('.nav-item').forEach(function(i){ i.classList.remove('active'); });
      navItem.classList.add('active');
    }
  });
});
'@

    # ─── BUDOWANIE HTML SEKCJI ───────────────────────────────────────────────────

    # Oblicz statystyki dla każdej sekcji
    function Get-SectionStats {
        param([object[]]$Data)
        $items = @($Data | Where-Object {
            $null -ne $_ -and $_.PSObject.Properties.Match("StatusLevel").Count -gt 0
        })
        $c = @($items | Where-Object { $_.StatusLevel -eq "CRITICAL" }).Count
        $w = @($items | Where-Object { $_.StatusLevel -eq "WARNING"  }).Count
        $o = @($items | Where-Object { $_.StatusLevel -eq "OK"       }).Count
        return @{ C=$c; W=$w; O=$o }
    }

    # 1. Serwery
    $srvStats = Get-SectionStats -Data $Servers
    $srvTable  = ConvertTo-HtmlTable -Data $Servers -TableId "tbl-servers" `
        -Columns @("ServerName","Role","Status","BuildVersion","NeedsUpgrade","StatusLevel") `
        -ColumnHeaders @("Serwer","Rola MinRole","Stan","Wersja Build","Wymaga aktualizacji","Status")
    $html_Servers = Build-SectionHTML -SectionId "servers" -SectionTitle "Serwery farmy" `
        -Icon "&#128268;" -TableHTML $srvTable -OKCount $srvStats.O -WarnCount $srvStats.W -CritCount $srvStats.C

    # 2. Service Applications
    $saStats = Get-SectionStats -Data $ServiceApps
    $saTable  = ConvertTo-HtmlTable -Data $ServiceApps -TableId "tbl-svcapps" `
        -Columns @("Name","Type","Status","StatusLevel") `
        -ColumnHeaders @("Nazwa","Typ","Stan","Status")
    $html_ServiceApps = Build-SectionHTML -SectionId "serviceapps" -SectionTitle "Service Applications" `
        -Icon "&#9881;" -TableHTML $saTable -OKCount $saStats.O -WarnCount $saStats.W -CritCount $saStats.C

    # 3. Usługi SharePoint
    $spSvcStats = Get-SectionStats -Data $SPServices
    $spSvcTable  = ConvertTo-HtmlTable -Data $SPServices -TableId "tbl-spsvcs" `
        -Columns @("ServerName","ServiceName","Status","StatusLevel") `
        -ColumnHeaders @("Serwer","Nazwa uslugi","Stan","Status")
    $html_SPServices = Build-SectionHTML -SectionId "spservices" -SectionTitle "Uslugi SharePoint" `
        -Icon "&#128295;" -TableHTML $spSvcTable -OKCount $spSvcStats.O -WarnCount $spSvcStats.W -CritCount $spSvcStats.C

    # 4. IIS — Witryny
    $siteStats = Get-SectionStats -Data $IISData.Sites
    $siteTable  = ConvertTo-HtmlTable -Data $IISData.Sites -TableId "tbl-iissites" `
        -Columns @("ServerName","SiteName","State","Bindings","StatusLevel") `
        -ColumnHeaders @("Serwer","Witryna IIS","Stan","Powiazania","Status")
    $html_IISSites = Build-SectionHTML -SectionId "iissites" -SectionTitle "Witryny IIS" `
        -Icon "&#127760;" -TableHTML $siteTable -OKCount $siteStats.O -WarnCount $siteStats.W -CritCount $siteStats.C

    # 5. App Poole
    $poolStats = Get-SectionStats -Data $IISData.AppPools
    $poolTable  = ConvertTo-HtmlTable -Data $IISData.AppPools -TableId "tbl-apppools" `
        -Columns @("ServerName","PoolName","State","ManagedRuntime","Identity","StatusLevel") `
        -ColumnHeaders @("Serwer","App Pool","Stan",".NET Runtime","Konto tozsamosci","Status")
    $html_AppPools = Build-SectionHTML -SectionId "apppools" -SectionTitle "Application Pools IIS" `
        -Icon "&#9878;" -TableHTML $poolTable -OKCount $poolStats.O -WarnCount $poolStats.W -CritCount $poolStats.C

    # 6. Aplikacje webowe
    $waStats = Get-SectionStats -Data $WebApps
    $waTable  = ConvertTo-HtmlTable -Data $WebApps -TableId "tbl-webapps" `
        -Columns @("Name","URL","ContentDBCount","AllowAnon","Status","StatusLevel") `
        -ColumnHeaders @("Nazwa","URL","Bazy content","Anonimowy","Dostepnosc","Status")
    $html_WebApps = Build-SectionHTML -SectionId "webapps" -SectionTitle "Aplikacje webowe SharePoint" `
        -Icon "&#127968;" -TableHTML $waTable -OKCount $waStats.O -WarnCount $waStats.W -CritCount $waStats.C

    # 7. Bazy danych
    $dbStats = Get-SectionStats -Data $Databases
    $dbTable  = ConvertTo-HtmlTable -Data $Databases -TableId "tbl-databases" `
        -Columns @("DBName","Type","Server","Size_MB","Status","NeedsUpgrade","ReadOnly","StatusLevel") `
        -ColumnHeaders @("Baza danych","Typ","SQL Server","Rozmiar [MB]","Stan","Wymaga upgr.","ReadOnly","Status")
    $html_Databases = Build-SectionHTML -SectionId "databases" -SectionTitle "Bazy danych SharePoint" `
        -Icon "&#128190;" -TableHTML $dbTable -OKCount $dbStats.O -WarnCount $dbStats.W -CritCount $dbStats.C

    # 8. Timer Jobs
    $tjStats = Get-SectionStats -Data $TimerJobs
    $tjTable  = ConvertTo-HtmlTable -Data $TimerJobs -TableId "tbl-timerjobs" `
        -Columns @("JobName","Status","LastRunTime","Duration_s","Schedule","Server","StatusLevel") `
        -ColumnHeaders @("Nazwa zadania","Ostatni status","Ostatnie uruchomienie","Czas [s]","Harmonogram","Serwer","Status")
    $html_TimerJobs = Build-SectionHTML -SectionId "timerjobs" -SectionTitle "Timer Jobs" `
        -Icon "&#8987;" -TableHTML $tjTable -OKCount $tjStats.O -WarnCount $tjStats.W -CritCount $tjStats.C

    # 9. Health Analyzer
    $haStats = Get-SectionStats -Data $HealthAnalyzer
    $haTable  = ConvertTo-HtmlTable -Data $HealthAnalyzer -TableId "tbl-health" `
        -Columns @("RuleName","Category","Severity","Server","FailureImpact","Remedy","StatusLevel") `
        -ColumnHeaders @("Nazwa reguly","Kategoria","Poziom","Serwer","Wplyw","Rekomendacja","Status")
    $html_Health = Build-SectionHTML -SectionId "health" -SectionTitle "Health Analyzer" `
        -Icon "&#10084;" -TableHTML $haTable -OKCount $haStats.O -WarnCount $haStats.W -CritCount $haStats.C

    # 10. Event Log
    $elStats = Get-SectionStats -Data $EventLog
    $elTable  = ConvertTo-HtmlTable -Data $EventLog -TableId "tbl-eventlog" `
        -Columns @("ServerName","LogName","EventID","Source","Level","TimeCreated","Message","StatusLevel") `
        -ColumnHeaders @("Serwer","Dziennik","ID","Zrodlo","Poziom","Czas","Tresc","Status")
    $elInfo   = "<p style='color:var(--text-dim);margin-bottom:8px;'>Pokazano max $($EventLog.Count) zdarzen z ostatnich $LogHours godzin.</p>"
    $html_EventLog = Build-SectionHTML -SectionId "eventlog" -SectionTitle "Windows Event Log" `
        -Icon "&#128220;" -TableHTML $elTable -OKCount $elStats.O -WarnCount $elStats.W -CritCount $elStats.C `
        -ExtraInfo $elInfo

    # 11. ULS
    $ulsStats = Get-SectionStats -Data $ULSLog
    $ulsTable  = ConvertTo-HtmlTable -Data $ULSLog -TableId "tbl-uls" `
        -Columns @("ServerName","Timestamp","Level","Area","Category","EventID","Message","StatusLevel") `
        -ColumnHeaders @("Serwer","Znacznik czasu","Poziom","Obszar","Kategoria","ID zdarzenia","Tresc","Status")
    $ulsProblemCount = @($ULSLog | Where-Object { $_.StatusLevel -in @("CRITICAL","WARNING") }).Count
    $ulsInfo  = if ($ULSLog.Count -eq 0) {
        "<p style='color:var(--text-dim);margin-bottom:8px;'>Brak wpisow Error/Critical w logach ULS z ostatnich $LogHours godzin.</p>"
    } elseif ($ulsProblemCount -eq 0) {
        "<p style='color:var(--text-dim);margin-bottom:8px;'>Brak wpisow Error/Critical (Unexpected, Exception, Critical, Assert, Error) z ostatnich $LogHours godzin.</p>"
    } else {
        "<p style='color:var(--text-dim);margin-bottom:8px;'>Pokazano $($ULSLog.Count) wpisow ULS (Error/Critical) z ostatnich $LogHours godzin.</p>"
    }
    $html_ULS = Build-SectionHTML -SectionId "uls" -SectionTitle "Logi ULS SharePoint" `
        -Icon "&#128196;" -TableHTML $ulsTable -OKCount $ulsStats.O -WarnCount $ulsStats.W -CritCount $ulsStats.C `
        -ExtraInfo $ulsInfo

    # 12. Certyfikaty
    $certStats = Get-SectionStats -Data $Certificates
    $certTable  = ConvertTo-HtmlTable -Data $Certificates -TableId "tbl-certs" `
        -Columns @("ServerName","SiteName","Subject","ExpiryDate","DaysLeft","IsExpired","StatusLevel") `
        -ColumnHeaders @("Serwer","Witryna IIS","Podmiot","Data wygasniecia","Dni pozostale","Wygasl","Status")
    $html_Certs = Build-SectionHTML -SectionId "certs" -SectionTitle "Certyfikaty SSL/TLS" `
        -Icon "&#128274;" -TableHTML $certTable -OKCount $certStats.O -WarnCount $certStats.W -CritCount $certStats.C

    # 13. Zasoby — Dyski
    $diskStats = Get-SectionStats -Data $SystemResources.Disks
    $diskTableHTML = "<div class='table-wrapper'><table id='tbl-disks' class='data-table'><thead><tr><th>Serwer</th><th>Dysk</th><th>Etykieta</th><th>Wolne [GB]</th><th>Lacznie [GB]</th><th>Uzyte [%]</th><th>Wykres</th><th>Status</th></tr></thead><tbody>"
    foreach ($d in $SystemResources.Disks) {
        $cls   = if ($d.StatusLevel -eq "CRITICAL") {"row-crit"} elseif ($d.StatusLevel -eq "WARNING") {"row-warn"} else {""}
        $bCls  = if ($d.StatusLevel -eq "CRITICAL") {"disk-bar-crit"} elseif ($d.StatusLevel -eq "WARNING") {"disk-bar-warn"} else {"disk-bar-ok"}
        $pct   = [Math]::Min($d.UsedPct, 100)
        $badge = Get-StatusBadge -Level $d.StatusLevel
        $diskTableHTML += "<tr class='$cls'><td>$($d.ServerName)</td><td>$($d.Drive)</td><td>$($d.Label)</td><td>$($d.FreeGB)</td><td>$($d.TotalGB)</td><td>$($d.UsedPct)%</td><td><div class='disk-bar-wrap'><div class='disk-bar $bCls' style='width:$pct%'></div></div></td><td>$badge</td></tr>"
    }
    $diskTableHTML += "</tbody></table></div>"
    $html_Disks = Build-SectionHTML -SectionId "disks" -SectionTitle "Dyski serwera" `
        -Icon "&#128190;" -TableHTML $diskTableHTML -OKCount $diskStats.O -WarnCount $diskStats.W -CritCount $diskStats.C

    # 14. Zasoby — Pamięć RAM
    $ramStats = Get-SectionStats -Data $SystemResources.Memory
    $ramTableHTML = "<div class='table-wrapper'><table id='tbl-ram' class='data-table'><thead><tr><th>Serwer</th><th>RAM lacznie [GB]</th><th>RAM wolne [GB]</th><th>Uzyte [%]</th><th>Wykres</th><th>Status</th></tr></thead><tbody>"
    foreach ($m in $SystemResources.Memory) {
        $cls  = if ($m.StatusLevel -eq "CRITICAL") {"row-crit"} elseif ($m.StatusLevel -eq "WARNING") {"row-warn"} else {""}
        $bCls = if ($m.StatusLevel -eq "CRITICAL") {"disk-bar-crit"} elseif ($m.StatusLevel -eq "WARNING") {"disk-bar-warn"} else {"disk-bar-ok"}
        $pct  = [Math]::Min($m.UsedPct, 100)
        $badge = Get-StatusBadge -Level $m.StatusLevel
        $ramTableHTML += "<tr class='$cls'><td>$($m.ServerName)</td><td>$($m.TotalRAM_GB)</td><td>$($m.FreeRAM_GB)</td><td>$($m.UsedPct)%</td><td><div class='disk-bar-wrap'><div class='disk-bar $bCls' style='width:$pct%'></div></div></td><td>$badge</td></tr>"
    }
    $ramTableHTML += "</tbody></table></div>"
    $html_RAM = Build-SectionHTML -SectionId "ram" -SectionTitle "Pamiec RAM" `
        -Icon "&#128293;" -TableHTML $ramTableHTML -OKCount $ramStats.O -WarnCount $ramStats.W -CritCount $ramStats.C

    # 15. Wersje farmy
    $verStats = Get-SectionStats -Data $FarmVersions
    $verTable  = ConvertTo-HtmlTable -Data $FarmVersions -TableId "tbl-versions" `
        -Columns @("ServerName","BuildVersion","IsConsistent","StatusLevel") `
        -ColumnHeaders @("Serwer","Wersja Build","Spojnosc","Status")
    $html_Versions = Build-SectionHTML -SectionId "versions" -SectionTitle "Spojnosc wersji farmy" `
        -Icon "&#128290;" -TableHTML $verTable -OKCount $verStats.O -WarnCount $verStats.W -CritCount $verStats.C

    # ─── SEKCJA PODSUMOWANIA ─────────────────────────────────────────────────────
    $critIssues = @($Script:AllIssues | Where-Object { $_.Severity -eq "CRITICAL" })
    $warnIssues = @($Script:AllIssues | Where-Object { $_.Severity -eq "WARNING"  })

    $summaryIssuesHTML = ""
    if ($Script:AllIssues.Count -eq 0) {
        $summaryIssuesHTML = "<p style='color:var(--ok);padding:16px 0;'>&#10003; Brak wykrytych problemow. Farma dziala poprawnie.</p>"
    } else {
        $summaryIssuesHTML = "<table class='issues-table'><thead><tr><th>Status</th><th>Sekcja</th><th>Obiekt</th><th>Opis</th><th>Serwer</th><th>Rekomendacja</th></tr></thead><tbody>"
        foreach ($issue in ($critIssues + $warnIssues)) {
            $badgeHTML = Get-StatusBadge -Level $issue.Severity
            $desc  = $issue.Description.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;")
            $obj   = $issue.ObjectName.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;")
            $reco  = $issue.Recommendation.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;")
            $srv   = $issue.Server.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;")
            $summaryIssuesHTML += "<tr><td>$badgeHTML</td><td><span class='tag'>$($issue.Section)</span></td><td>$obj</td><td>$desc</td><td>$srv</td><td><small>$reco</small></td></tr>"
        }
        $summaryIssuesHTML += "</tbody></table>"
    }

    $html_Summary = @"
<div class="report-section" id="section-summary">
  <div class="section-header section-header-$($globalStatus.ToLower().Replace('ok','ok').Replace('warning','warn').Replace('critical','crit'))" onclick="toggleSection('summary')">
    <span class="section-icon">&#128203;</span>
    <span class="section-title">Podsumowanie ogolne farmy</span>
    <div class="section-badges">
      <span class="badge badge-crit">KRYTYCZNE: $($critIssues.Count)</span>
      <span class="badge badge-warn">OSTRZEZENIA: $($warnIssues.Count)</span>
    </div>
    <span class="section-toggle" id="toggle-summary">&#9660;</span>
  </div>
  <div class="section-body" id="body-summary">
    <div class="summary-grid">
      <div class="stat-card card-crit"><div class="stat-num num-crit">$($critIssues.Count)</div><div class="stat-label">Problemy krytyczne</div></div>
      <div class="stat-card card-warn"><div class="stat-num num-warn">$($warnIssues.Count)</div><div class="stat-label">Ostrzezenia</div></div>
      <div class="stat-card card-ok"><div class="stat-num num-info">$($Servers.Count)</div><div class="stat-label">Serwery farmy</div></div>
      <div class="stat-card card-ok"><div class="stat-num num-info">$($WebApps.Count)</div><div class="stat-label">Aplikacje webowe</div></div>
      <div class="stat-card card-ok"><div class="stat-num num-info">$($Databases.Count)</div><div class="stat-label">Bazy danych</div></div>
      <div class="stat-card card-ok"><div class="stat-num num-info">$($ServiceApps.Count)</div><div class="stat-label">Service Applications</div></div>
      <div class="stat-card"><div class="stat-num num-info">$($TimerJobs.Count)</div><div class="stat-label">Timer Jobs</div></div>
      <div class="stat-card"><div class="stat-num num-info">$($Certificates.Count)</div><div class="stat-label">Certyfikaty SSL</div></div>
    </div>
    <h3 style="margin-bottom:10px;font-size:14px;color:var(--text-dim);">Wykryte problemy wymagajace uwagi:</h3>
    $summaryIssuesHTML
  </div>
</div>
"@

    # ─── LISTA SERWERÓW DLA FILTRA ────────────────────────────────────────────────
    $allServerNames = @()
    if ($Servers) { $allServerNames += $Servers | Select-Object -ExpandProperty ServerName -ErrorAction SilentlyContinue }
    $allServerNames = $allServerNames | Select-Object -Unique | Sort-Object
    $serverOptions  = ($allServerNames | ForEach-Object { "<option value='$_'>$_</option>" }) -join ""

    # ─── NAWIGACJA BOCZNA ────────────────────────────────────────────────────────
    $navItems = @(
        @{ target="section-summary";    label="Podsumowanie";         icon="&#128203;" }
        @{ target="section-servers";    label="Serwery farmy";        icon="&#128268;" }
        @{ target="section-serviceapps";label="Service Applications"; icon="&#9881;" }
        @{ target="section-spservices"; label="Uslugi SharePoint";    icon="&#128295;" }
        @{ target="section-iissites";   label="Witryny IIS";          icon="&#127760;" }
        @{ target="section-apppools";   label="App Poole IIS";        icon="&#9878;" }
        @{ target="section-webapps";    label="Aplikacje webowe";     icon="&#127968;" }
        @{ target="section-databases";  label="Bazy danych";          icon="&#128190;" }
        @{ target="section-timerjobs";  label="Timer Jobs";           icon="&#8987;" }
        @{ target="section-health";     label="Health Analyzer";      icon="&#10084;" }
        @{ target="section-eventlog";   label="Event Log";            icon="&#128220;" }
        @{ target="section-uls";        label="Logi ULS";             icon="&#128196;" }
        @{ target="section-certs";      label="Certyfikaty SSL";      icon="&#128274;" }
        @{ target="section-disks";      label="Dyski";                icon="&#128190;" }
        @{ target="section-ram";        label="Pamiec RAM";           icon="&#128293;" }
        @{ target="section-versions";   label="Wersje farmy";         icon="&#128290;" }
    )
    $navHTML = ($navItems | ForEach-Object {
        "<span class='nav-item' data-target='$($_.target)'>$($_.icon) $($_.label)</span>"
    }) -join "`n"

    # ─── SKŁADANIE PEŁNEGO HTML ───────────────────────────────────────────────────
    <#
    $globalStatusLabel = switch ($globalStatus) {
        "CRITICAL" { "KRYTYCZNY — Farma wymaga natychmiastowej uwagi!" }
        "WARNING"  { "OSTRZEZENIE — Wykryto problemy wymagajace uwagi" }
        default    { "ZDROWA — Farma dziala poprawnie" }
    }

    #>
    $globalStatusLabel = switch ($globalStatus) {
        "CRITICAL" { "KRYTYCZNY - Farma wymaga natychmiastowej uwagi!" }
        "WARNING"  { "OSTRZEZENIE - Wykryto problemy wymagajace uwagi" }
        default    { "ZDROWA - Farma dziala poprawnie" }
    }

    $htmlDoc = @"
<!DOCTYPE html>
<html lang="pl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Farm Daily Operations — $farmName — $now</title>
<style>$css</style>
</head>
<body>

<!-- HEADER -->
<div class="header">
  <h1>&#127968; Farm Daily Operations — $farmName</h1>
  <div class="header-meta">
    <span>&#128197; Data generowania: <strong>$now</strong></span>
    <span>&#128268; Serwer: <strong>$scriptSrv</strong></span>
    <span>&#8987; Czas wykonania: <strong>${duration}s</strong></span>
    <span>&#128203; Analizowany okres: <strong>ostatnie $LogHours godzin</strong></span>
  </div>
</div>

<!-- GLOBAL STATUS BANNER -->
<div class="status-banner $globalClass">
  <div class="status-dot"></div>
  <span class="status-text">Stan farmy: $globalStatusLabel</span>
  <div class="status-counts">
    <span class="badge badge-crit">$($critIssues.Count) krytycznych</span>
    <span class="badge badge-warn">$($warnIssues.Count) ostrzezen</span>
  </div>
</div>

<!-- LAYOUT -->
<div class="container">

  <!-- SIDEBAR NAV -->
  <nav class="sidebar">
    <div class="sidebar-title">Nawigacja</div>
    $navHTML
  </nav>

  <!-- MAIN CONTENT -->
  <main class="main-content">

    <!-- TOOLBAR -->
    <div class="toolbar">
      <input type="text" id="globalSearch" class="search-box" placeholder="&#128269; Szukaj w tabelach..." oninput="globalSearch()">
      <select id="statusFilter" class="filter-select" onchange="filterStatus()">
        <option value="">Wszystkie statusy</option>
        <option value="CRITICAL">CRITICAL</option>
        <option value="WARNING">WARNING</option>
        <option value="OK">OK</option>
      </select>
      <select id="serverFilter" class="filter-select" onchange="filterServer()">
        <option value="">Wszystkie serwery</option>
        $serverOptions
      </select>
      <button class="btn btn-secondary" onclick="resetFilters()">&#10006; Reset</button>
      <button class="btn btn-secondary" onclick="expandAll()">&#9660; Rozwin wszystko</button>
      <button class="btn btn-secondary" onclick="collapseAll()">&#9650; Zwij wszystko</button>
    </div>

    <!-- SEKCJE RAPORTU -->
    $html_Summary
    $html_Servers
    $html_ServiceApps
    $html_SPServices
    $html_IISSites
    $html_AppPools
    $html_WebApps
    $html_Databases
    $html_TimerJobs
    $html_Health
    $html_EventLog
    $html_ULS
    $html_Certs
    $html_Disks
    $html_RAM
    $html_Versions

  </main>
</div>

<div class="footer">
  Farm Daily Operations HTML &mdash; SharePoint Subscription Edition &mdash;
  Wygenerowano: $now &mdash; Serwer: $scriptSrv
</div>

<script>$js</script>
</body>
</html>
"@

    return $htmlDoc
}

#endregion

#region ─── ZAPIS RAPORTU ────────────────────────────────────────────────────────

function Save-Report {
    param([string]$HtmlContent)

    Write-Log "Zapisywanie raportu HTML..."

    # Utwórz katalog jeśli nie istnieje
    if (-not (Test-Path $ReportOutputPath)) {
        try {
            New-Item -ItemType Directory -Path $ReportOutputPath -Force | Out-Null
            Write-Log "Utworzono katalog raportu: $ReportOutputPath"
        }
        catch {
            Write-Log "Nie mozna utworzyc katalogu '$ReportOutputPath': $($_.Exception.Message)" -Level ERROR
            $ReportOutputPath = $env:TEMP
            Write-Log "Uzywam katalogu tymczasowego: $ReportOutputPath" -Level WARNING
        }
    }

    # Zapis głównego raportu
    $reportPath = Join-Path $ReportOutputPath $ReportFileName
    try {
        [System.IO.File]::WriteAllText($reportPath, $HtmlContent, [System.Text.Encoding]::UTF8)
        Write-Log "Raport zapisany: $reportPath"
    }
    catch {
        Write-Log "Blad zapisu raportu: $($_.Exception.Message)" -Level ERROR
    }

    # Opcjonalny archiwum z datą i godziną w nazwie
    if ($ArchiveReports) {
        $archiveName = [System.IO.Path]::GetFileNameWithoutExtension($ReportFileName) +
                       "_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".html"
        $archivePath = Join-Path $ReportOutputPath $archiveName
        try {
            [System.IO.File]::WriteAllText($archivePath, $HtmlContent, [System.Text.Encoding]::UTF8)
            Write-Log "Kopia archiwalna zapisana: $archivePath"
        }
        catch {
            Write-Log "Blad zapisu archiwum: $($_.Exception.Message)" -Level WARNING
        }
    }

    return $reportPath
}

#endregion

#region ─── WYSYŁKA E-MAIL ───────────────────────────────────────────────────────

function Send-ReportEmail {
    param([string]$ReportFilePath)

    if (-not $SendEmail) { return }
    <#
    if (-not $SMTPServer) { Write-Log "Nie podano SMTPServer — pomijam wysylke." -Level WARNING; return }
    if (-not $EmailFrom)  { Write-Log "Nie podano EmailFrom — pomijam wysylke."  -Level WARNING; return }
    if ($EmailTo.Count -eq 0) { Write-Log "Nie podano adresow EmailTo — pomijam wysylke." -Level WARNING; return }

    #>
    if (-not $SMTPServer) { Write-Log "Nie podano SMTPServer - pomijam wysylke." -Level WARNING; return }
    if (-not $EmailFrom)  { Write-Log "Nie podano EmailFrom - pomijam wysylke."  -Level WARNING; return }
    if ($EmailTo.Count -eq 0) { Write-Log "Nie podano adresow EmailTo - pomijam wysylke." -Level WARNING; return }

    Write-Log "Wysylanie raportu e-mailem do: $($EmailTo -join ', ')..."

    try {
        $mailParams = @{
            SmtpServer  = $SMTPServer
            Port        = $SMTPPort
            From        = $EmailFrom
            To          = $EmailTo
            Subject     = $EmailSubject
            Body        = $EmailBody
            Attachments = $ReportFilePath
            Encoding    = [System.Text.Encoding]::UTF8
        }
        if ($UseSSL)            { $mailParams.UseSsl     = $true }
        if ($SMTPCredential)    { $mailParams.Credential = $SMTPCredential }

        Send-MailMessage @mailParams -ErrorAction Stop
        Write-Log "E-mail wyslany pomyslnie."
    }
    catch {
        Write-Log "Blad wysylki e-mail: $($_.Exception.Message)" -Level ERROR
        Write-Log "Raport dostepny lokalnie pod sciezka: $ReportFilePath" -Level WARNING
    }
}

#endregion

#region ─── SEKCJA GŁÓWNA — URUCHOMIENIE WSZYSTKICH FUNKCJI ─────────────────────

<#
Write-Log "=============================================="
Write-Log " Farm Daily Operations HTML — START"
Write-Log " Data: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Log " Serwer: $Script:ScriptServer"
Write-Log "=============================================="

#>
Write-Log "=============================================="
Write-Log " Farm Daily Operations HTML - START"
Write-Log " Data: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Log " Serwer: $Script:ScriptServer"
Write-Log "=============================================="

# Krok 0: Walidacja
$isAdmin = Test-AdminPrivileges

# Krok 1: Inicjalizacja SharePoint
Initialize-SharePointEnvironment

# Krok 2: Zbieranie danych — każda sekcja w osobnym bloku try/catch
# Jeśli jedna sekcja nie zadziała, pozostałe nadal są wykonywane.

Write-Log "--- Zbieranie danych z farmy ---"

$data_Servers = try { Get-FarmServers }         catch { Write-Log "BLAD sekcji Servers: $($_.Exception.Message)" -Level ERROR
    @([PSCustomObject]@{ ServerName="BLAD"; Role="N/A"; Status="Blad"; BuildVersion="N/A"; NeedsUpgrade="N/A"; StatusLevel="CRITICAL" }) }

$data_ServiceApps = try { Get-ServiceApplications } catch { Write-Log "BLAD sekcji ServiceApps: $($_.Exception.Message)" -Level ERROR
    @([PSCustomObject]@{ Name="BLAD"; Type="N/A"; Status="Blad"; StatusLevel="CRITICAL" }) }

$data_SPServices = try { Get-SharePointServices }  catch { Write-Log "BLAD sekcji SPServices: $($_.Exception.Message)" -Level ERROR
    @([PSCustomObject]@{ ServerName="BLAD"; ServiceName="Blad"; ServiceType="N/A"; Status="Blad"; StatusLevel="CRITICAL" }) }

$data_IIS = try { Get-IISStatus } catch { Write-Log "BLAD sekcji IIS: $($_.Exception.Message)" -Level ERROR
    @{ Sites = @([PSCustomObject]@{ ServerName="BLAD"; SiteName="Blad"; State="Blad"; PhysPath=""; Bindings=""; StatusLevel="CRITICAL" })
       AppPools = @([PSCustomObject]@{ ServerName="BLAD"; PoolName="Blad"; State="Blad"; ManagedRuntime=""; Identity=""; StatusLevel="CRITICAL" }) } }

$data_WebApps = try { Get-WebApplications }       catch { Write-Log "BLAD sekcji WebApps: $($_.Exception.Message)" -Level ERROR
    @([PSCustomObject]@{ Name="BLAD"; URL="N/A"; ContentDBCount=0; AllowAnon="N/A"; Status="Blad"; StatusLevel="CRITICAL" }) }

$data_Databases = try { Get-Databases }           catch { Write-Log "BLAD sekcji Databases: $($_.Exception.Message)" -Level ERROR
    @([PSCustomObject]@{ DBName="BLAD"; Type="N/A"; Server="N/A"; Size_MB=0; Status="Blad"; NeedsUpgrade="N/A"; ReadOnly="N/A"; StatusLevel="CRITICAL" }) }

$data_TimerJobs = try { Get-TimerJobStatus }      catch { Write-Log "BLAD sekcji TimerJobs: $($_.Exception.Message)" -Level ERROR
    @([PSCustomObject]@{ JobName="BLAD"; Status="Blad"; LastRunTime="N/A"; Duration_s=0; Schedule="N/A"; Server="N/A"; StatusLevel="CRITICAL" }) }

$data_Health = try { Get-HealthAnalyzerRules }    catch { Write-Log "BLAD sekcji Health: $($_.Exception.Message)" -Level ERROR
    @([PSCustomObject]@{ RuleName="BLAD"; Category="N/A"; Severity="N/A"; Server="N/A"; FailureImpact="N/A"; Remedy="Blad pobierania"; StatusLevel="WARNING" }) }

$data_EventLog = try { Get-EventLogErrors }       catch { Write-Log "BLAD sekcji EventLog: $($_.Exception.Message)" -Level ERROR
    @([PSCustomObject]@{ ServerName="BLAD"; LogName="N/A"; EventID=0; Source="Blad"; Level="Error"; TimeCreated="N/A"; Message="Blad pobierania Event Log"; StatusLevel="WARNING" }) }

$data_ULS = try { Get-ULSErrors }                 catch { Write-Log "BLAD sekcji ULS: $($_.Exception.Message)" -Level ERROR
    @([PSCustomObject]@{ ServerName="BLAD"; Timestamp="N/A"; Level="Error"; Area="N/A"; Category="N/A"; EventID="N/A"; Message="Blad pobierania ULS"; StatusLevel="WARNING" }) }

$data_Certs = try { Get-CertificateStatus }       catch { Write-Log "BLAD sekcji Certs: $($_.Exception.Message)" -Level ERROR
    @([PSCustomObject]@{ ServerName="BLAD"; SiteName="Blad"; Binding="N/A"; Subject="Blad"; Thumbprint="N/A"; ExpiryDate="N/A"; DaysLeft=0; IsExpired=$false; StatusLevel="WARNING" }) }

$data_Resources = try { Get-SystemResources }     catch { Write-Log "BLAD sekcji Resources: $($_.Exception.Message)" -Level ERROR
    @{ Disks  = @([PSCustomObject]@{ ServerName="BLAD"; Drive="Blad"; Label="N/A"; FreeGB=0; TotalGB=0; UsedPct=0; StatusLevel="WARNING" })
       Memory = @([PSCustomObject]@{ ServerName="BLAD"; TotalRAM_GB=0; FreeRAM_GB=0; UsedPct=0; StatusLevel="WARNING" }) } }

$data_Versions = try { Get-FarmVersionConsistency } catch { Write-Log "BLAD sekcji Versions: $($_.Exception.Message)" -Level ERROR
    @([PSCustomObject]@{ ServerName="BLAD"; BuildVersion="N/A"; IsConsistent="N/A"; StatusLevel="WARNING" }) }

# Krok 3: Budowanie raportu HTML
Write-Log "--- Budowanie raportu HTML ---"
$htmlContent = Build-HTMLReport `
    -Servers         $data_Servers `
    -SPServices      $data_SPServices `
    -IISData         $data_IIS `
    -WebApps         $data_WebApps `
    -Databases       $data_Databases `
    -TimerJobs       $data_TimerJobs `
    -HealthAnalyzer  $data_Health `
    -EventLog        $data_EventLog `
    -ULSLog          $data_ULS `
    -Certificates    $data_Certs `
    -SystemResources $data_Resources `
    -FarmVersions    $data_Versions `
    -ServiceApps     $data_ServiceApps

# Krok 4: Zapis raportu
$savedPath = Save-Report -HtmlContent $htmlContent

# Krok 5: Opcjonalna wysyłka e-mail
Send-ReportEmail -ReportFilePath $savedPath

# Krok 6: Podsumowanie końcowe
$endTime      = Get-Date
$totalSeconds = [Math]::Round(($endTime - $Script:StartTime).TotalSeconds, 1)
$critCount    = @($Script:AllIssues | Where-Object { $_.Severity -eq "CRITICAL" }).Count
$warnCount    = @($Script:AllIssues | Where-Object { $_.Severity -eq "WARNING"  }).Count

<#
Write-Log "=============================================="
Write-Log " Farm Daily Operations HTML — KONIEC"
Write-Log " Czas wykonania    : ${totalSeconds}s"
Write-Log " Raport zapisany   : $savedPath"
Write-Log " Log skryptu       : $Script:LogFile"
Write-Log " Problemy krytyczne: $critCount"
Write-Log " Ostrzezenia       : $warnCount"
Write-Log " Status globalny   : $(if($critCount -gt 0){'CRITICAL'}elseif($warnCount -gt 0){'WARNING'}else{'OK'})"
Write-Log "=============================================="

# Zwróć ścieżkę raportu (przydatne przy wywoływaniu z innych skryptów)
#>
Write-Log "=============================================="
Write-Log " Farm Daily Operations HTML - KONIEC"
Write-Log " Czas wykonania    : ${totalSeconds}s"
Write-Log " Raport zapisany   : $savedPath"
Write-Log " Log skryptu       : $Script:LogFile"
Write-Log " Problemy krytyczne: $critCount"
Write-Log " Ostrzezenia       : $warnCount"
Write-Log " Status globalny   : $(if($critCount -gt 0){'CRITICAL'}elseif($warnCount -gt 0){'WARNING'}else{'OK'})"
Write-Log "=============================================="

return $savedPath

#endregion
