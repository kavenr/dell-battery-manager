<#
.SYNOPSIS
    Dell Battery Manager : consulte et modifie les réglages de batterie du BIOS d'un portable Dell.

.DESCRIPTION
    - Mode de charge, seuils personnalisés, interrupteurs Charge avancée / Peak Shift, seuil Peak Shift :
      provider WMI BIOS Dell (root\dcim\sysman\biosattributes), réponse immédiate.
    - Plannings par jour (Charge avancée, Peak Shift) : cctk.exe de Dell Command | Configure déjà installé
      (~20 s par appel), lus seulement à l'ouverture de l'onglet concerné.
    - Infos batterie : WMI Windows (root\wmi) + rapport powercfg pour la capacité d'origine.
    Requiert Windows PowerShell 5.1 et les droits administrateur (élévation demandée automatiquement).

.PARAMETER ShowConsole
    Garde la fenêtre console visible (diagnostic).

.PARAMETER Demo
    Lance l'interface avec des réglages BIOS fictifs, sans droits admin et sans rien écrire.

.PARAMETER Snapshot
    Avec -Demo : enregistre une capture PNG de chaque onglet (<Snapshot>-1.png, ...), puis quitte.

.PARAMETER Language
    Langue de l'interface pour cette exécution (fr ou en), sans modifier la préférence enregistrée.
    Sinon : choix fait dans le menu Langue, ou à défaut l'anglais.
#>
param(
    [switch]$ShowConsole,
    [switch]$Demo,
    [string]$Snapshot,
    [ValidateSet('', 'fr', 'en')][string]$Language = ''
)

# =====================================================================================
#  Cœur : accès BIOS et batterie, sans interface (aussi chargé dans les tâches de fond)
# =====================================================================================
$Core = {
    # ---------------------------------------------------------------- Traductions
    # Chaque clé donne le texte français puis anglais. La langue active est $global:DbmShared.Lang ('fr' ou 'en') :
    # une hashtable synchronisée, la même pour l'interface et les tâches de fond.
    function Get-TextTable {
        @{
            # Menus
            'Menu.File'           = '_Fichier', '_File'
            'Menu.Refresh'        = '_Actualiser', '_Refresh'
            'Menu.Quit'           = '_Quitter', '_Quit'
            'Menu.Language'       = '_Langue', '_Language'

            # Bannière : Dell Command | Configure absent
            'Banner.Title'        = 'Dell Command | Configure est introuvable.', 'Dell Command | Configure was not found.'
            'Banner.Text'         = "Sans cet outil Dell (cctk.exe), les plannings Charge avancée et Peak Shift ne peuvent être ni lus ni modifiés.",
                                    'Without this Dell tool (cctk.exe), the Advanced Charge and Peak Shift schedules can be neither read nor changed.'
            'Banner.Link'         = "Télécharger l'outil sur le site de Dell", 'Download the tool from the Dell website'
            'Banner.Url'          = 'https://www.dell.com/support/kbdoc/fr-fr/000178000/dell-command-configure',
                                    'https://www.dell.com/support/kbdoc/en-us/000178000/dell-command-configure'
            'Banner.Recheck'      = 'Vérifier à nouveau', 'Check again'
            'Banner.StillMissing' = 'Dell Command | Configure est toujours introuvable.', 'Dell Command | Configure is still not found.'
            'Banner.Found'        = 'Dell Command | Configure détecté.', 'Dell Command | Configure detected.'

            # Barre du bas, onglets, boutons communs
            'App.DemoSuffix'      = ' (démo : réglages fictifs)', ' (demo: simulated settings)'
            'Btn.Refresh'         = 'Actualiser', 'Refresh'
            'Btn.Revert'          = 'Annuler les modifications', 'Discard changes'
            'Btn.Apply'           = 'Appliquer', 'Apply'
            'Btn.ReadBios'        = 'Relire le BIOS', 'Reload from BIOS'
            'Btn.CopyMonday'      = 'Copier lundi sur mar.–ven.', 'Copy Monday to Tue–Fri'
            'Pwd.Label'           = 'Mot de passe BIOS :', 'BIOS password:'
            'Tab.Battery'         = 'Batterie', 'Battery'
            'Tab.Mode'            = 'Mode de charge', 'Charge mode'
            'Tab.Advanced'        = 'Charge avancée', 'Advanced Charge'
            'Tab.PeakShift'       = 'Peak Shift', 'Peak Shift'

            # Onglet Batterie
            'Bat.Mode'            = 'Mode de charge du BIOS', 'BIOS charge mode'
            'Bat.Remaining'       = 'Énergie restante', 'Remaining energy'
            'Bat.Full'            = 'Capacité à pleine charge', 'Full charge capacity'
            'Bat.Design'          = "Capacité d'origine", 'Design capacity'
            'Bat.Health'          = 'Santé', 'Health'
            'Bat.Cycles'          = 'Cycles de charge', 'Charge cycles'
            'Bat.Rate'            = 'Puissance', 'Power'
            'Bat.Voltage'         = 'Tension', 'Voltage'
            'Bat.Hint'            = "Mis à jour toutes les 15 secondes. La santé compare la capacité à pleine charge à la capacité d'origine.",
                                    'Updated every 15 seconds. Health compares the full charge capacity with the design capacity.'
            'Bat.Charging'        = 'Branché, en charge', 'Plugged in, charging'
            'Bat.NotCharging'     = 'Branché, pas en charge', 'Plugged in, not charging'
            'Bat.Paused'          = 'Charge en pause : batterie pleine ou seuil du mode de charge atteint.',
                                    'Charging paused: battery full or charge mode threshold reached.'
            'Bat.OnBattery'       = 'Sur batterie', 'On battery'
            'Bat.RunTime'         = 'Autonomie estimée : {0} h {1:D2}', 'Estimated time remaining: {0} h {1:D2} min'
            'Bat.RateCharging'    = '{0:N1} W en charge', '{0:N1} W charging'
            'Bat.RateDraw'        = '{0:N1} W consommés', '{0:N1} W drawn'
            'Bat.NoCycles'        = 'non fourni par la batterie', 'not reported by the battery'
            'Bat.HealthValue'     = "{0:N0} % de la capacité d'origine", '{0:N0}% of design capacity'
            'Bat.ModeCustom'      = ' : charge de {0} % à {1} %', ': charges from {0}% to {1}%'
            'Bat.AdvOn'           = '  ·  Charge avancée activée', '  ·  Advanced Charge on'
            'Bat.PeakOn'          = '  ·  Peak Shift activé', '  ·  Peak Shift on'
            'Fmt.Wh'              = '{0:N1} Wh', '{0:N1} Wh'
            'Fmt.V'               = '{0:N2} V', '{0:N2} V'
            'Fmt.Pct'             = '{0} %', '{0}%'

            # Onglet Mode de charge
            'Mode.Intro'          = 'Choisissez comment le BIOS charge la batterie.', 'Choose how the BIOS charges the battery.'
            'Mode.Start'          = 'Début de la charge', 'Start charging at'
            'Mode.Stop'           = 'Arrêt de la charge', 'Stop charging at'
            'Mode.CustomHint'     = "La charge reprend quand la batterie descend sous le seuil de début et s'arrête au seuil d'arrêt. Écart minimal : 5 %.",
                                    'Charging resumes when the battery drops below the start threshold and stops at the stop threshold. Minimum gap: 5%.'
            'Mode.Standard.Name'  = 'Standard', 'Standard'
            'Mode.Standard.Desc'  = 'Charge complète, sur une durée plus longue.', 'Full charge, over a longer time.'
            'Mode.Express.Name'   = 'ExpressCharge', 'ExpressCharge'
            'Mode.Express.Desc'   = 'Charge rapide (technologie Dell). Sollicite davantage la batterie.',
                                    'Fast charging (Dell technology). Puts more strain on the battery.'
            'Mode.PrimAcUse.Name' = 'Utilisation principale sur secteur', 'Primarily AC use'
            'Mode.PrimAcUse.Desc' = 'Pour un portable presque toujours branché : le BIOS abaisse le seuil de charge pour ménager la batterie.',
                                    'For a laptop that is almost always plugged in: the BIOS lowers the charge threshold to preserve the battery.'
            'Mode.Adaptive.Name'  = 'Adaptatif', 'Adaptive'
            'Mode.Adaptive.Desc'  = 'Le BIOS ajuste la charge selon votre utilisation habituelle. Pas de seuil fixe.',
                                    'The BIOS adjusts charging to your usual usage. No fixed threshold.'
            'Mode.Custom.Name'    = 'Personnalisé', 'Custom'
            'Mode.Custom.Desc'    = "La charge démarre et s'arrête aux seuils choisis ci-dessous.", 'Charging starts and stops at the thresholds chosen below.'
            'Mode.GapProblem'     = "Le seuil d'arrêt doit dépasser le seuil de début d'au moins 5 %.",
                                    'The stop threshold must be at least 5% above the start threshold.'
            'Change.Mode'         = 'Mode de charge : {0} → {1}', 'Charge mode: {0} → {1}'
            'Change.Start'        = 'Début de la charge : {0} % → {1} %', 'Start charging at: {0}% → {1}%'
            'Change.Stop'         = 'Arrêt de la charge : {0} % → {1} %', 'Stop charging at: {0}% → {1}%'

            # Onglets de planning
            'Adv.Hint'            = 'Ménage la batterie en dehors des heures de travail et la charge rapidement avant et pendant la période de travail de chaque jour.',
                                    "Preserves the battery outside working hours and charges it quickly before and during each day's work period."
            'Adv.Enable'          = 'Activer la charge avancée', 'Enable Advanced Charge'
            'Adv.Col1'            = 'Début de journée', 'Start of day'
            'Adv.Col2'            = 'Durée de travail', 'Work period'
            'Peak.Hint'           = "Entre Début et Fin, le portable fonctionne sur batterie tant qu'elle reste au-dessus du seuil, même branché. La charge ne reprend qu'à l'heure de reprise.",
                                    'Between Start and End, the laptop runs on battery while it stays above the threshold, even when plugged in. Charging only resumes at the resume time.'
            'Peak.Enable'         = 'Activer Peak Shift', 'Enable Peak Shift'
            'Peak.Threshold'      = 'Seuil de batterie', 'Battery threshold'
            'Peak.Col1'           = 'Début', 'Start'
            'Peak.Col2'           = 'Fin', 'End'
            'Peak.Col3'           = 'Reprise de la charge', 'Charge resume'
            'Day.Mon'             = 'Lundi', 'Monday'
            'Day.Tue'             = 'Mardi', 'Tuesday'
            'Day.Wed'             = 'Mercredi', 'Wednesday'
            'Day.Thu'             = 'Jeudi', 'Thursday'
            'Day.Fri'             = 'Vendredi', 'Friday'
            'Day.Sat'             = 'Samedi', 'Saturday'
            'Day.Sun'             = 'Dimanche', 'Sunday'
            'Sched.Raw'           = 'Valeur lue par cctk : {0}', 'Value read by cctk: {0}'
            'Sched.NoteLastKnown' = "Le BIOS ne renvoie pas de planning : dernier planning connu de l'app affiché.",
                                    'The BIOS returns no schedule: showing the last schedule known to the app.'
            'Sched.NoteDefault'   = 'Le BIOS ne renvoie pas de planning : valeurs par défaut affichées.',
                                    'The BIOS returns no schedule: showing default values.'
            'Sched.NoCctk'        = 'Dell Command | Configure (cctk) est introuvable : les plannings ne peuvent être ni lus ni modifiés.',
                                    'Dell Command | Configure (cctk) was not found: schedules can be neither read nor changed.'
            'Sched.ErrMinutes'    = '{0} : les minutes doivent valoir 00, 15, 30 ou 45.', '{0}: minutes must be 00, 15, 30 or 45.'
            'Sched.ErrOrder'      = '{0} : il faut Début ≤ Fin ≤ Reprise de la charge.', '{0}: Start ≤ End ≤ Charge resume is required.'
            'Sched.SumAdv'        = '      {0} : début {1}, durée {2}', '      {0}: start {1}, duration {2}'
            'Sched.SumPeak'       = '      {0} : {1} → {2}, reprise {3}', '      {0}: {1} → {2}, resume {3}'
            'Sched.ChangeOn'      = '•  {0} : activé avec ce planning', '•  {0}: enabled with this schedule'
            'Sched.ChangeOff'     = '•  {0} : désactivé', '•  {0}: disabled'
            'Sched.ChangeThresh'  = '•  Seuil Peak Shift : {0} % → {1} %', '•  Peak Shift threshold: {0}% → {1}%'

            # Statut, dialogues, erreurs
            'Status.Starting'     = 'Démarrage…', 'Starting…'
            'Status.Error'        = 'Erreur : {0}', 'Error: {0}'
            'Status.ReadingBios'  = 'Lecture des réglages du BIOS…', 'Reading BIOS settings…'
            'Status.BiosRead'     = 'Réglages lus à {0:HH:mm}.', 'Settings read at {0:HH:mm}.'
            'Status.ReadingSched' = 'Lecture des plannings via cctk (~20 s)…', 'Reading schedules with cctk (~20 s)…'
            'Status.SchedRead'    = 'Plannings lus à {0:HH:mm}.', 'Schedules read at {0:HH:mm}.'
            'Status.Writing'      = 'Écriture dans le BIOS…', 'Writing to the BIOS…'
            'Status.WritingSched' = 'Écriture via cctk puis relecture (~40 s)…', 'Writing with cctk, then reading back (~40 s)…'
            'Status.Applied'      = 'Réglages appliqués.', 'Settings applied.'
            'Status.SchedApplied' = 'Planning appliqué.', 'Schedule applied.'
            'Status.Pending'      = "Écrit, mais le BIOS renvoie encore l'ancienne valeur pour : {0}", 'Written, but the BIOS still returns the old value for: {0}'
            'Status.SchedDiffers' = 'cctk a accepté la commande, mais le BIOS renvoie un état différent.', 'cctk accepted the command, but the BIOS returns a different state.'
            'Status.DemoNoWrite'  = "Mode démo : rien n'a été écrit.", 'Demo mode: nothing was written.'
            'Status.Demo'         = 'Mode démo : réglages BIOS fictifs, aucune écriture.', 'Demo mode: simulated BIOS settings, nothing is written.'
            'Dlg.Confirm'         = "Écrire ces réglages dans le BIOS ?`n`n{0}", "Write these settings to the BIOS?`n`n{0}"
            'Dlg.ConfirmTitle'    = 'Confirmer', 'Confirm'
            'Dlg.WriteInProgress' = "Une écriture dans le BIOS est en cours. Attendez qu'elle se termine.", 'A BIOS write is in progress. Wait for it to finish.'
            'Err.PasswordNeeded'  = "Un mot de passe administrateur BIOS est défini : saisissez-le en bas de la fenêtre avant d'appliquer.",
                                    'A BIOS administrator password is set: enter it at the bottom of the window before applying.'
            'Err.BiosRefused'     = 'Le BIOS a refusé {0}={1} : {2} (code {3}).', 'The BIOS rejected {0}={1}: {2} (code {3}).'
            'Err.Bios1'           = 'échec', 'failed'
            'Err.Bios2'           = 'paramètre invalide', 'invalid parameter'
            'Err.Bios3'           = 'accès refusé (mot de passe BIOS ?)', 'access denied (BIOS password?)'
            'Err.Bios4'           = 'non supporté', 'not supported'
            'Err.Bios5'           = 'erreur mémoire', 'memory error'
            'Err.Bios6'           = 'erreur de protocole', 'protocol error'
            'Err.BiosUnknown'     = 'erreur inconnue', 'unknown error'
            'Err.CctkMissing'     = "cctk.exe introuvable : Dell Command | Configure n'est pas installé.", 'cctk.exe not found: Dell Command | Configure is not installed.'
            'Err.CctkRead'        = "Lecture de {0} par cctk impossible (code {1}) :`n{2}", "cctk could not read {0} (code {1}):`n{2}"
            'Err.CctkWrite'       = "cctk a refusé la modification (code {0}) :`n{1}", "cctk rejected the change (code {0}):`n{1}"
        }
    }

    function Tr {
        # Tr 'Clé' [arguments de mise en forme...] : texte dans la langue active, nombres au format de cette langue.
        param([string]$Key)
        if (-not $global:DbmText) { $global:DbmText = Get-TextTable }
        $en = -not $global:DbmShared -or $global:DbmShared.Lang -ne 'fr'   # anglais par défaut
        $pair = $global:DbmText[$Key]
        $text = if (-not $pair) { $Key } elseif ($en) { $pair[1] } else { $pair[0] }
        if (-not $args.Count) { return $text }
        $culture = [Globalization.CultureInfo]::GetCultureInfo($(if ($en) { 'en-US' } else { 'fr-FR' }))
        [string]::Format($culture, $text, [object[]]$args)
    }

    function Get-DellBatteryAttributes {
        $names = 'PrimaryBattChargeCfg', 'CustomChargeStart', 'CustomChargeStop',
                 'AdvBatteryChargeCfg', 'PeakShiftCfg', 'PeakShiftBatteryThreshold'
        $result = @{}
        foreach ($class in 'EnumerationAttribute', 'IntegerAttribute') {
            foreach ($a in Get-CimInstance -Namespace 'root\dcim\sysman\biosattributes' -ClassName $class) {
                if ($names -notcontains $a.AttributeName) { continue }
                $result[$a.AttributeName] = [pscustomobject]@{
                    Name     = $a.AttributeName
                    Current  = [string](@($a.CurrentValue)[0])
                    Possible = @($a.PossibleValue | Where-Object { $_ })
                    Lower    = $a.LowerBound
                    Upper    = $a.UpperBound
                }
            }
        }
        $result
    }

    function Test-DellAdminPassword {
        $admin = Get-CimInstance -Namespace 'root\dcim\sysman\wmisecurity' -ClassName PasswordObject |
            Where-Object { $_.NameId -eq 'Admin' }
        [bool]($admin -and $admin.IsPasswordSet -eq 1)
    }

    function Set-DellBiosAttribute {
        param([string]$Name, [string]$Value, [string]$Password)
        $bai = @(Get-WmiObject -Namespace 'root\dcim\sysman\biosattributes' -Class BIOSAttributeInterface)[0]
        if ($Password) {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($Password)
            $r = $bai.SetAttribute(1, $bytes.Length, $bytes, $Name, $Value)
        } else {
            $r = $bai.SetAttribute(0, 0, 0, $Name, $Value)
        }
        if ($r.Status -ne 0) {
            $code = [int]$r.Status
            $reason = if ($code -ge 1 -and $code -le 6) { Tr "Err.Bios$code" } else { Tr 'Err.BiosUnknown' }
            throw (Tr 'Err.BiosRefused' $Name $Value $reason $code)
        }
    }

    function Find-Cctk {
        foreach ($base in ${env:ProgramFiles(x86)}, $env:ProgramFiles) {
            if (-not $base) { continue }
            $path = Join-Path $base 'Dell\Command Configure\X86_64\cctk.exe'
            if (Test-Path -LiteralPath $path) { return $path }
        }
        $null
    }

    function Invoke-Cctk {
        param([string[]]$Arguments)
        $exe = Find-Cctk
        if (-not $exe) { throw (Tr 'Err.CctkMissing') }
        $output = & $exe @Arguments 2>&1 | Out-String
        [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $output }
    }

    function Format-Time {
        # "8:0" -> "08:00" ; rien si l'heure est invalide.
        param([string]$Text)
        if ($Text -match '^\s*(\d{1,2}):(\d{1,2})\s*$') {
            $h = [int]$Matches[1]; $m = [int]$Matches[2]
            if ($h -le 23 -and $m -le 59) { '{0:D2}:{1:D2}' -f $h, $m }
        }
    }

    function ConvertFrom-CctkSchedule {
        # "Enabled,Mon-10:00/08:00,Tue-..." -> Enabled + jours ; "Disabled" -> aucun jour.
        param([string]$Value, [int]$FieldCount)
        $parts = @($Value -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $days = [ordered]@{}
        foreach ($part in ($parts | Select-Object -Skip 1)) {
            if ($part -notmatch '^([A-Za-z]{3})[A-Za-z]*-(.+)$') { continue }
            $day = (Get-Culture).TextInfo.ToTitleCase($Matches[1].ToLower())
            $times = @($Matches[2] -split '/' | ForEach-Object { Format-Time $_ })
            if ($times.Count -eq $FieldCount) { $days[$day] = $times }
        }
        [pscustomobject]@{
            Enabled = ($parts.Count -gt 0 -and $parts[0] -eq 'Enabled')
            Days    = $days
            Raw     = $Value
        }
    }

    function ConvertTo-CctkSchedule {
        param([bool]$Enabled, $Days)
        if (-not $Enabled) { return 'Disabled' }
        $items = foreach ($d in 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun') {
            if ($Days.Contains($d)) { '{0}-{1}' -f $d, ($Days[$d] -join '/') }
        }
        (@('Enabled') + @($items)) -join ','
    }

    function Get-ScheduleErrors {
        param($Days, [string]$Kind)
        foreach ($d in @($Days.Keys)) {
            $minutes = @($Days[$d] | ForEach-Object { $p = $_ -split ':'; [int]$p[0] * 60 + [int]$p[1] })
            if (@($minutes | Where-Object { $_ % 15 -ne 0 }).Count) {
                Tr 'Sched.ErrMinutes' (Tr "Day.$d")
            }
            if ($Kind -eq 'PeakShift' -and -not ($minutes[0] -le $minutes[1] -and $minutes[1] -le $minutes[2])) {
                Tr 'Sched.ErrOrder' (Tr "Day.$d")
            }
        }
    }

    function Get-ChargeWriteOrder {
        # Ordre d'écriture de CustomChargeStart/Stop qui respecte l'écart minimal à chaque étape :
        # on monte d'abord l'arrêt si le début monte, sinon on descend d'abord le début.
        param([int]$OldStart, [int]$NewStart, [int]$OldStop, [int]$NewStop)
        $start = if ($NewStart -ne $OldStart) { @{ Name = 'CustomChargeStart'; Value = $NewStart } }
        $stop = if ($NewStop -ne $OldStop) { @{ Name = 'CustomChargeStop'; Value = $NewStop } }
        if ($NewStart -gt $OldStart) { $stop, $start | Where-Object { $_ } } else { $start, $stop | Where-Object { $_ } }
    }

    function Read-CctkSchedules {
        $options = 'AdvBatteryChargeCfg', 'PeakShiftCfg'
        $values = @{}
        $r = Invoke-Cctk ($options | ForEach-Object { "--$_" })
        foreach ($line in $r.Output -split "`r?`n") {
            if ($line -match '^\s*(AdvBatteryChargeCfg|PeakShiftCfg)=(.*)$') { $values[$Matches[1]] = $Matches[2].Trim() }
        }
        foreach ($opt in $options) {
            # Repli : un appel par option si l'appel groupé n'a pas tout renvoyé.
            if ($values.ContainsKey($opt)) { continue }
            $single = Invoke-Cctk @("--$opt")
            if ($single.Output -match "(?m)^\s*$opt=(.*)$") { $values[$opt] = $Matches[1].Trim() }
            else { throw (Tr 'Err.CctkRead' $opt $single.ExitCode $single.Output.Trim()) }
        }
        [pscustomobject]@{
            Advanced  = ConvertFrom-CctkSchedule $values['AdvBatteryChargeCfg'] 2
            PeakShift = ConvertFrom-CctkSchedule $values['PeakShiftCfg'] 3
        }
    }

    function Write-CctkSchedules {
        param([string]$Advanced, [string]$PeakShift, [string]$Password)
        $cctkArgs = @()
        if ($Advanced) { $cctkArgs += "--AdvBatteryChargeCfg=$Advanced" }
        if ($PeakShift) { $cctkArgs += "--PeakShiftCfg=$PeakShift" }
        if (-not $cctkArgs) { return }
        if ($Password) { $cctkArgs += "--ValSetupPwd=$Password" }
        $r = Invoke-Cctk $cctkArgs
        if ($r.ExitCode -ne 0) {
            # Le message part à l'écran et dans erreurs.log : on n'y laisse jamais le mot de passe.
            $output = $r.Output.Trim()
            if ($Password) { $output = $output.Replace($Password, '***') }
            throw (Tr 'Err.CctkWrite' $r.ExitCode $output)
        }
    }

    function Get-BatteryLive {
        $status = @(Get-CimInstance -Namespace 'root\wmi' -ClassName BatteryStatus -ErrorAction SilentlyContinue)[0]
        $full = @(Get-CimInstance -Namespace 'root\wmi' -ClassName BatteryFullChargedCapacity -ErrorAction SilentlyContinue)[0]
        $cycles = @(Get-CimInstance -Namespace 'root\wmi' -ClassName BatteryCycleCount -ErrorAction SilentlyContinue)[0]
        $win = @(Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue)[0]
        [pscustomobject]@{
            Percent       = $win.EstimatedChargeRemaining
            RunTimeMin    = $win.EstimatedRunTime
            Remaining     = $status.RemainingCapacity
            FullCharge    = $full.FullChargedCapacity
            PowerOnline   = [bool]$status.PowerOnline
            Charging      = [bool]$status.Charging
            Discharging   = [bool]$status.Discharging
            ChargeRate    = $status.ChargeRate
            DischargeRate = $status.DischargeRate
            Voltage       = $status.Voltage
            CycleCount    = $cycles.CycleCount
        }
    }

    function Get-BatteryReport {
        $file = Join-Path $env:TEMP ('dbm-battery-report-{0}.xml' -f $PID)
        try {
            & powercfg.exe /batteryreport /xml /output $file | Out-Null
            [xml]$xml = Get-Content -Raw -LiteralPath $file
            $b = @($xml.BatteryReport.Batteries.Battery)[0]
            [pscustomobject]@{
                Design     = [int64]$b.DesignCapacity
                FullCharge = [int64]$b.FullChargeCapacity
                Cycles     = [int]$b.CycleCount
            }
        } finally {
            Remove-Item -LiteralPath $file -ErrorAction SilentlyContinue
        }
    }
}

$functionsBefore = @(Get-ChildItem function: | ForEach-Object Name)
. $Core
$CoreFunctionNames = @(Get-ChildItem function: | Where-Object { $functionsBefore -notcontains $_.Name } | ForEach-Object Name)

# =====================================================================================
#  Interface
# =====================================================================================
$AppXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Dell Battery Manager" Width="800" Height="660" MinWidth="680" MinHeight="560"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" FontSize="13">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="14,5"/>
      <Setter Property="MinWidth" Value="90"/>
      <Setter Property="Margin" Value="8,0,0,0"/>
    </Style>
    <Style x:Key="Primary" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="#0067C0"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderBrush" Value="#0067C0"/>
    </Style>
    <Style x:Key="Hint" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#5F5F5F"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="Problem" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#C42B1C"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="Label" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#5F5F5F"/>
      <Setter Property="Margin" Value="0,5,24,5"/>
    </Style>
    <Style x:Key="Value" TargetType="TextBlock">
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,5,0,5"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="White"/>
      <Setter Property="BorderBrush" Value="#DDDDDD"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="4"/>
      <Setter Property="Padding" Value="14"/>
    </Style>
  </Window.Resources>

  <!-- Les textes "@Clé" sont remplacés au démarrage par la traduction de la clé (voir Get-TextTable). -->
  <DockPanel Background="#F3F3F3">
    <Menu DockPanel.Dock="Top" Background="#F3F3F3" Padding="4,2">
      <MenuItem Header="@Menu.File">
        <MenuItem x:Name="MenuRefresh" Header="@Menu.Refresh" InputGestureText="F5"/>
        <Separator/>
        <MenuItem x:Name="MenuQuit" Header="@Menu.Quit"/>
      </MenuItem>
      <MenuItem Header="@Menu.Language">
        <MenuItem x:Name="MenuLangFr" Header="Français"/>
        <MenuItem x:Name="MenuLangEn" Header="English"/>
      </MenuItem>
    </Menu>

    <Border x:Name="CctkBanner" DockPanel.Dock="Top" Background="#FFF4CE" BorderBrush="#E3C567" BorderThickness="0,1,0,1"
            Padding="14,8" Visibility="Collapsed">
      <DockPanel>
        <Button x:Name="BannerRecheck" DockPanel.Dock="Right" Content="@Banner.Recheck" VerticalAlignment="Center" Margin="12,0,0,0"/>
        <TextBlock DockPanel.Dock="Left" Text="&#xE7BA;" FontFamily="Segoe MDL2 Assets" FontSize="16" Foreground="#9D5D00"
                   VerticalAlignment="Center" Margin="0,0,10,0"/>
        <TextBlock TextWrapping="Wrap" VerticalAlignment="Center">
          <Run FontWeight="SemiBold" Text="@Banner.Title"/>
          <Run Text="@Banner.Text"/>
          <Hyperlink x:Name="BannerLink"><Run Text="@Banner.Link"/></Hyperlink>
        </TextBlock>
      </DockPanel>
    </Border>

    <Border DockPanel.Dock="Bottom" Background="#E8E8E8" BorderBrush="#D0D0D0" BorderThickness="0,1,0,0" Padding="12,6">
      <DockPanel>
        <Button x:Name="BtnRefresh" DockPanel.Dock="Right" Content="@Btn.Refresh"/>
        <StackPanel x:Name="PwdPanel" DockPanel.Dock="Right" Orientation="Horizontal" Visibility="Collapsed" Margin="0,0,4,0">
          <TextBlock Text="@Pwd.Label" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <PasswordBox x:Name="BiosPwd" Width="140" VerticalAlignment="Center"/>
        </StackPanel>
        <ProgressBar x:Name="Busy" DockPanel.Dock="Left" Width="70" Height="8" IsIndeterminate="True"
                     Visibility="Collapsed" Margin="0,0,10,0" VerticalAlignment="Center"/>
        <TextBlock x:Name="StatusText" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
      </DockPanel>
    </Border>

    <TabControl x:Name="Tabs" Margin="12" Padding="16">

      <TabItem Header="@Tab.Battery">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <StackPanel>
            <StackPanel Orientation="Horizontal">
              <TextBlock x:Name="BatPercent" Text="–" FontSize="40" FontWeight="SemiBold"/>
              <StackPanel Margin="18,0,0,0" VerticalAlignment="Center">
                <TextBlock x:Name="BatState" FontSize="16"/>
                <TextBlock x:Name="BatTime" Style="{StaticResource Hint}"/>
              </StackPanel>
            </StackPanel>
            <ProgressBar x:Name="BatBar" Height="8" Margin="0,10,0,16" Maximum="100"/>
            <Border Style="{StaticResource Card}">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/>
                  <RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/>
                </Grid.RowDefinitions>
                <TextBlock Grid.Row="0" Style="{StaticResource Label}" Text="@Bat.Mode"/>
                <TextBlock Grid.Row="0" Grid.Column="1" x:Name="BatMode" Style="{StaticResource Value}" Text="–"/>
                <TextBlock Grid.Row="1" Style="{StaticResource Label}" Text="@Bat.Remaining"/>
                <TextBlock Grid.Row="1" Grid.Column="1" x:Name="BatRemaining" Style="{StaticResource Value}" Text="–"/>
                <TextBlock Grid.Row="2" Style="{StaticResource Label}" Text="@Bat.Full"/>
                <TextBlock Grid.Row="2" Grid.Column="1" x:Name="BatFull" Style="{StaticResource Value}" Text="–"/>
                <TextBlock Grid.Row="3" Style="{StaticResource Label}" Text="@Bat.Design"/>
                <TextBlock Grid.Row="3" Grid.Column="1" x:Name="BatDesign" Style="{StaticResource Value}" Text="–"/>
                <TextBlock Grid.Row="4" Style="{StaticResource Label}" Text="@Bat.Health"/>
                <TextBlock Grid.Row="4" Grid.Column="1" x:Name="BatHealth" Style="{StaticResource Value}" Text="–"/>
                <TextBlock Grid.Row="5" Style="{StaticResource Label}" Text="@Bat.Cycles"/>
                <TextBlock Grid.Row="5" Grid.Column="1" x:Name="BatCycles" Style="{StaticResource Value}" Text="–"/>
                <TextBlock Grid.Row="6" Style="{StaticResource Label}" Text="@Bat.Rate"/>
                <TextBlock Grid.Row="6" Grid.Column="1" x:Name="BatRate" Style="{StaticResource Value}" Text="–"/>
                <TextBlock Grid.Row="7" Style="{StaticResource Label}" Text="@Bat.Voltage"/>
                <TextBlock Grid.Row="7" Grid.Column="1" x:Name="BatVoltage" Style="{StaticResource Value}" Text="–"/>
              </Grid>
            </Border>
            <TextBlock Style="{StaticResource Hint}" Margin="0,10,0,0" Text="@Bat.Hint"/>
          </StackPanel>
        </ScrollViewer>
      </TabItem>

      <TabItem Header="@Tab.Mode">
        <DockPanel>
          <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="ModeRevert" Content="@Btn.Revert"/>
            <Button x:Name="ModeApply" Content="@Btn.Apply" Style="{StaticResource Primary}"/>
          </StackPanel>
          <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel>
              <TextBlock Style="{StaticResource Hint}" Margin="0,0,0,8" Text="@Mode.Intro"/>
              <StackPanel x:Name="ModeList"/>
              <Border x:Name="CustomPanel" Style="{StaticResource Card}" Margin="26,6,0,0">
                <StackPanel>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="56"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions><RowDefinition/><RowDefinition/></Grid.RowDefinitions>
                    <TextBlock Text="@Mode.Start" VerticalAlignment="Center" Margin="0,6,16,6"/>
                    <Slider x:Name="StartSlider" Grid.Column="1" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center"/>
                    <TextBlock x:Name="StartValue" Grid.Column="2" FontWeight="SemiBold" VerticalAlignment="Center" Margin="10,0,0,0"/>
                    <TextBlock Grid.Row="1" Text="@Mode.Stop" VerticalAlignment="Center" Margin="0,6,16,6"/>
                    <Slider x:Name="StopSlider" Grid.Row="1" Grid.Column="1" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center"/>
                    <TextBlock x:Name="StopValue" Grid.Row="1" Grid.Column="2" FontWeight="SemiBold" VerticalAlignment="Center" Margin="10,0,0,0"/>
                  </Grid>
                  <TextBlock Style="{StaticResource Hint}" Margin="0,8,0,0" Text="@Mode.CustomHint"/>
                </StackPanel>
              </Border>
              <TextBlock x:Name="ModeProblem" Style="{StaticResource Problem}" Margin="0,8,0,0"/>
            </StackPanel>
          </ScrollViewer>
        </DockPanel>
      </TabItem>

      <TabItem x:Name="TabAdv" Header="@Tab.Advanced">
        <DockPanel>
          <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="AdvRead" Content="@Btn.ReadBios"/>
            <Button x:Name="AdvCopy" Content="@Btn.CopyMonday"/>
            <Button x:Name="AdvRevert" Content="@Btn.Revert"/>
            <Button x:Name="AdvApply" Content="@Btn.Apply" Style="{StaticResource Primary}"/>
          </StackPanel>
          <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel>
              <TextBlock Style="{StaticResource Hint}" Text="@Adv.Hint"/>
              <CheckBox x:Name="AdvEnabled" Content="@Adv.Enable" Margin="0,12,0,12"/>
              <Grid x:Name="AdvGrid"/>
              <TextBlock x:Name="AdvRaw" Style="{StaticResource Hint}" Margin="0,12,0,0" FontSize="11"/>
              <TextBlock x:Name="AdvProblem" Style="{StaticResource Problem}" Margin="0,6,0,0"/>
            </StackPanel>
          </ScrollViewer>
        </DockPanel>
      </TabItem>

      <TabItem x:Name="TabPeak" Header="@Tab.PeakShift">
        <DockPanel>
          <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="PeakRead" Content="@Btn.ReadBios"/>
            <Button x:Name="PeakCopy" Content="@Btn.CopyMonday"/>
            <Button x:Name="PeakRevert" Content="@Btn.Revert"/>
            <Button x:Name="PeakApply" Content="@Btn.Apply" Style="{StaticResource Primary}"/>
          </StackPanel>
          <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel>
              <TextBlock Style="{StaticResource Hint}" Text="@Peak.Hint"/>
              <CheckBox x:Name="PeakEnabled" Content="@Peak.Enable" Margin="0,12,0,8"/>
              <Grid Margin="0,0,0,12">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="260"/>
                  <ColumnDefinition Width="60"/>
                </Grid.ColumnDefinitions>
                <TextBlock Text="@Peak.Threshold" VerticalAlignment="Center" Margin="0,0,16,0"/>
                <Slider x:Name="PeakThreshold" Grid.Column="1" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center"/>
                <TextBlock x:Name="PeakThresholdValue" Grid.Column="2" FontWeight="SemiBold" VerticalAlignment="Center" Margin="10,0,0,0"/>
              </Grid>
              <Grid x:Name="PeakGrid"/>
              <TextBlock x:Name="PeakRaw" Style="{StaticResource Hint}" Margin="0,12,0,0" FontSize="11"/>
              <TextBlock x:Name="PeakProblem" Style="{StaticResource Problem}" Margin="0,6,0,0"/>
            </StackPanel>
          </ScrollViewer>
        </DockPanel>
      </TabItem>

    </TabControl>
  </DockPanel>
</Window>
'@

# Dot-sourcé (tests) : on s'arrête ici, seules les fonctions du cœur sont chargées.
if ($MyInvocation.InvocationName -eq '.') { return }

# =====================================================================================
#  Démarrage : Windows PowerShell 5.1, STA, administrateur
# =====================================================================================
if (-not $Demo) {
    $principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $isSta = [Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA'
    if (-not $isAdmin -or -not $isSta -or $PSVersionTable.PSVersion.Major -ne 5) {
        $argList = '-NoProfile -ExecutionPolicy Bypass -STA -File "{0}"' -f $PSCommandPath
        if ($ShowConsole) { $argList += ' -ShowConsole' }
        if ($Language) { $argList += " -Language $Language" }
        $start = @{ FilePath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"; ArgumentList = $argList }
        if (-not $isAdmin) { $start.Verb = 'RunAs' }
        if (-not $ShowConsole) { $start.WindowStyle = 'Hidden' }
        try { Start-Process @start } catch { }   # élévation refusée : on quitte simplement
        exit
    }

    if (-not $ShowConsole) {
        Add-Type -Namespace DellBatteryManager -Name NativeConsole -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern System.IntPtr GetConsoleWindow();
[DllImport("kernel32.dll")] public static extern uint GetConsoleProcessList(uint[] list, uint count);
[DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
'@
        # On ne cache la console que si elle nous appartient (pas celle d'un terminal ouvert par l'utilisateur).
        if ([DellBatteryManager.NativeConsole]::GetConsoleProcessList((New-Object uint[] 4), 4) -le 1) {
            [void][DellBatteryManager.NativeConsole]::ShowWindow([DellBatteryManager.NativeConsole]::GetConsoleWindow(), 0)
        }
    }
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$window = [Windows.Markup.XamlReader]::Parse($AppXaml)
$ui = @{}
foreach ($m in [regex]::Matches($AppXaml, 'x:Name="(\w+)"')) { $ui[$m.Groups[1].Value] = $window.FindName($m.Groups[1].Value) }

# ---------------------------------------------------------------- Langue
$SettingsFile = Join-Path $env:LOCALAPPDATA 'DellBatteryManager\settings.json'

function Get-SavedLanguage {
    try {
        if (Test-Path -LiteralPath $SettingsFile) {
            $saved = (Get-Content -Raw -LiteralPath $SettingsFile | ConvertFrom-Json).Language
            if ($saved -in 'fr', 'en') { return $saved }
        }
    } catch { }
    'en'   # langue par défaut
}

function Save-Language {
    param([string]$Lang)
    try {
        New-Item -ItemType Directory -Force -Path (Split-Path $SettingsFile) | Out-Null
        @{ Language = $Lang } | ConvertTo-Json | Set-Content -Encoding UTF8 -LiteralPath $SettingsFile
    } catch { }   # préférence non enregistrée : sans gravité
}

$global:DbmShared = [hashtable]::Synchronized(@{ Lang = $(if ($Language) { $Language } else { Get-SavedLanguage }) })

# Éléments dont le texte suit la langue : { Element, Property, Key }.
$script:Localized = New-Object System.Collections.ArrayList

function Set-Loc {
    # Affiche la traduction de $Key dans $Element.$Property et la met à jour à chaque changement de langue.
    param($Element, [string]$Property, [string]$Key)
    [void]$script:Localized.Add(@{ Element = $Element; Property = $Property; Key = $Key })
    $Element.$Property = Tr $Key
}

function Register-LocalizedTree {
    # Parcourt l'arbre logique du XAML et enregistre chaque texte de la forme "@Clé".
    param($Node)
    $isSimpleText = $Node -is [Windows.Controls.TextBlock] -and $Node.Inlines.Count -le 1
    $props = if ($isSimpleText -or $Node -is [Windows.Documents.Run]) { 'Text' } else { 'Header', 'Content' }
    foreach ($prop in $props) {
        $value = $Node.$prop
        if ($value -is [string] -and $value.StartsWith('@')) { Set-Loc $Node $prop $value.Substring(1) }
    }
    if ($isSimpleText) { return }
    # Copie de la liste : remplacer un texte modifie les enfants logiques pendant le parcours.
    foreach ($child in @([Windows.LogicalTreeHelper]::GetChildren($Node))) {
        if ($child -is [Windows.DependencyObject]) { Register-LocalizedTree $child }
    }
}

$DayKeys = 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'
$TimeChoices = foreach ($h in 0..23) { foreach ($m in 0, 15, 30, 45) { '{0:D2}:{1:D2}' -f $h, $m } }

# Modes de charge connus, dans l'ordre d'affichage (textes : Mode.<clé>.Name / Mode.<clé>.Desc).
$ModeKeys = 'Standard', 'Express', 'PrimAcUse', 'Adaptive', 'Custom'

$SchedInfo = @{
    Advanced  = @{ Prefix = 'Adv'; Title = 'Tab.Advanced'; Headers = 'Adv.Col1', 'Adv.Col2'; Default = '08:00', '10:00' }
    PeakShift = @{ Prefix = 'Peak'; Title = 'Tab.PeakShift'; Headers = 'Peak.Col1', 'Peak.Col2', 'Peak.Col3'; Default = '16:00', '20:00', '22:00' }
}

$script:Bios = $null              # attributs WMI lus
$script:SchedData = $null         # plannings lus par cctk
$script:Base = @{}                # planning affiché au chargement (référence pour détecter les modifications)
$script:ModeRadios = $null
$script:PasswordRequired = $false
$script:Writing = $false
$script:SchedLoading = $false
$script:BatteryJobRunning = $false
$script:BatteryLive = $null
$script:Report = $null
$script:LastStatus = $null        # @{ Key; Args } : retraduit si la langue change
$script:RawInfo = @{}             # texte « Valeur lue par cctk » de chaque onglet de planning : @{ Raw; Note }
$script:CctkPath = $null

# ---------------------------------------------------------------- Tâches de fond
$iss = [Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
foreach ($name in $CoreFunctionNames) {
    $iss.Commands.Add((New-Object Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $name, (Get-Item "function:$name").Definition))
}
# Même objet dans toutes les tâches de fond : leurs messages d'erreur suivent la langue choisie.
$iss.Variables.Add((New-Object Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList 'DbmShared', $global:DbmShared, ''))
$pool = [runspacefactory]::CreateRunspacePool(1, 3, $iss, $Host)
$pool.Open()
$script:Jobs = New-Object System.Collections.ArrayList

function Start-Background {
    param(
        [scriptblock]$Work,
        [hashtable]$Arguments = @{},
        [string]$Status,
        $Context,
        [scriptblock]$OnSuccess,
        [scriptblock]$OnFinally,
        [switch]$Quiet
    )
    $ps = [powershell]::Create()
    $ps.RunspacePool = $pool
    [void]$ps.AddScript($Work.ToString())
    if ($Arguments.Count) { [void]$ps.AddParameters($Arguments) }
    [void]$script:Jobs.Add([pscustomobject]@{
        Ps = $ps; Handle = $ps.BeginInvoke(); Status = $Status; Context = $Context
        OnSuccess = $OnSuccess; OnFinally = $OnFinally; Quiet = [bool]$Quiet
    })
    Update-Status
}

$WorkReadBios = {
    $ErrorActionPreference = 'Stop'
    [pscustomobject]@{ Attributes = Get-DellBatteryAttributes; PasswordRequired = Test-DellAdminPassword }
}
$WorkBattery = { $ErrorActionPreference = 'Stop'; Get-BatteryLive }
$WorkReport = { $ErrorActionPreference = 'Stop'; Get-BatteryReport }
$WorkReadSchedules = { $ErrorActionPreference = 'Stop'; Read-CctkSchedules }
$WorkApplyMode = {
    param($Steps, [string]$Password)
    $ErrorActionPreference = 'Stop'
    foreach ($step in $Steps) { Set-DellBiosAttribute -Name $step.Name -Value ([string]$step.Value) -Password $Password }
    Get-DellBatteryAttributes
}
$WorkApplySchedule = {
    param([string]$Advanced, [string]$PeakShift, $Threshold, [string]$Password)
    $ErrorActionPreference = 'Stop'
    if ($null -ne $Threshold) { Set-DellBiosAttribute -Name 'PeakShiftBatteryThreshold' -Value ([string]$Threshold) -Password $Password }
    Write-CctkSchedules -Advanced $Advanced -PeakShift $PeakShift -Password $Password
    [pscustomobject]@{ Schedules = Read-CctkSchedules; Attributes = Get-DellBatteryAttributes }
}

# ---------------------------------------------------------------- Statut et erreurs
function Set-Status {
    # Set-Status 'Clé' [arguments...] : le texte est retraduit si la langue change.
    param([string]$Key)
    $script:LastStatus = @{ Key = $Key; Args = @($args) }
    Update-Status
}

function Update-Status {
    $active = @($script:Jobs | Where-Object { $_.Status })
    $ui.Busy.Visibility = if ($active.Count) { 'Visible' } else { 'Collapsed' }
    $ui.StatusText.Text = if ($active.Count) { Tr $active[-1].Status }
                          elseif ($script:LastStatus) { $a = $script:LastStatus.Args; Tr $script:LastStatus.Key @a }
                          else { '' }
}

function Get-InnermostMessage {
    param($ErrorObject)
    $ex = if ($ErrorObject -is [Management.Automation.ErrorRecord]) { $ErrorObject.Exception } else { $ErrorObject }
    while ($ex.InnerException) { $ex = $ex.InnerException }
    $ex.Message
}

$ErrorLogFile = Join-Path $env:LOCALAPPDATA 'DellBatteryManager\erreurs.log'

function Write-ErrorLog {
    # Journal de diagnostic : message, pile d'appels et ligne fautive de chaque erreur signalée.
    param($ErrorObject, [string]$Origin)
    try {
        $lines = @('==== {0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date), $Origin)
        $lines += Get-InnermostMessage $ErrorObject
        if ($ErrorObject -is [Management.Automation.ErrorRecord]) {
            if ($ErrorObject.ScriptStackTrace) { $lines += $ErrorObject.ScriptStackTrace }
            if ($ErrorObject.InvocationInfo) { $lines += $ErrorObject.InvocationInfo.PositionMessage }
        }
        New-Item -ItemType Directory -Force -Path (Split-Path $ErrorLogFile) | Out-Null
        Add-Content -Encoding UTF8 -LiteralPath $ErrorLogFile -Value $lines
    } catch { }
}

function Show-Error {
    param($ErrorObject, [string]$Origin = 'interface')
    $message = Get-InnermostMessage $ErrorObject
    Write-ErrorLog $ErrorObject $Origin
    Set-Status 'Status.Error' $message
    [void][Windows.MessageBox]::Show($window, $message, 'Dell Battery Manager', 'OK', 'Error')
}

function Invoke-Guarded {
    param([scriptblock]$Action)
    try { & $Action } catch { Show-Error $_ }
}

function Confirm-Write {
    param([string]$Summary)
    'Yes' -eq [Windows.MessageBox]::Show($window, (Tr 'Dlg.Confirm' $Summary), (Tr 'Dlg.ConfirmTitle'), 'YesNo', 'Question')
}

function Get-BiosPassword {
    if (-not $script:PasswordRequired) { return '' }
    if (-not $ui.BiosPwd.Password) { throw (Tr 'Err.PasswordNeeded') }
    $ui.BiosPwd.Password
}

# ---------------------------------------------------------------- Onglet Batterie
function Format-Wh {
    param($MilliWattHours)
    if (-not $MilliWattHours) { return '–' }
    Tr 'Fmt.Wh' ($MilliWattHours / 1000)
}

function Show-BatteryLive {
    param($Battery)
    if ($Battery) { $script:BatteryLive = $Battery } else { $Battery = $script:BatteryLive }
    if (-not $Battery) { return }
    $b = $Battery
    if ($null -ne $b.Percent) {
        $ui.BatPercent.Text = Tr 'Fmt.Pct' $b.Percent
        $ui.BatBar.Value = [double]$b.Percent
    }
    $hint = ''
    if ($b.PowerOnline) {
        if ($b.Charging) { $ui.BatState.Text = Tr 'Bat.Charging' }
        else {
            $ui.BatState.Text = Tr 'Bat.NotCharging'
            $hint = Tr 'Bat.Paused'
        }
    } else {
        $ui.BatState.Text = Tr 'Bat.OnBattery'
        if ($b.RunTimeMin -and $b.RunTimeMin -lt 71582788) {
            $hint = Tr 'Bat.RunTime' ([math]::Floor($b.RunTimeMin / 60)) ([int]($b.RunTimeMin % 60))
        }
    }
    $ui.BatTime.Text = $hint
    $ui.BatRemaining.Text = Format-Wh $b.Remaining
    $ui.BatFull.Text = Format-Wh $b.FullCharge
    $ui.BatRate.Text = if ($b.Charging -and $b.ChargeRate) { Tr 'Bat.RateCharging' ($b.ChargeRate / 1000) }
                       elseif ($b.Discharging -and $b.DischargeRate) { Tr 'Bat.RateDraw' ($b.DischargeRate / 1000) }
                       else { '0 W' }
    $ui.BatVoltage.Text = if ($b.Voltage) { Tr 'Fmt.V' ($b.Voltage / 1000) } else { '–' }

    $r = $script:Report
    $cycles = if ($b.CycleCount) { $b.CycleCount } elseif ($r -and $r.Cycles) { $r.Cycles } else { $null }
    $ui.BatCycles.Text = if ($null -ne $cycles) { "$cycles" } else { Tr 'Bat.NoCycles' }
    if ($r -and $r.Design) {
        $full = if ($b.FullCharge) { $b.FullCharge } else { $r.FullCharge }
        $ui.BatDesign.Text = Format-Wh $r.Design
        $ui.BatHealth.Text = Tr 'Bat.HealthValue' (100 * $full / $r.Design)
    }
}

function Start-BatteryRefresh {
    if ($script:BatteryJobRunning) { return }
    $script:BatteryJobRunning = $true
    Start-Background -Work $WorkBattery -Quiet -OnSuccess { param($b) Show-BatteryLive $b } -OnFinally { $script:BatteryJobRunning = $false }
}

function Show-ActiveMode {
    $b = $script:Bios
    $mode = $b['PrimaryBattChargeCfg'].Current
    $text = Get-ModeName $mode
    if ($mode -eq 'Custom') { $text += Tr 'Bat.ModeCustom' $b['CustomChargeStart'].Current $b['CustomChargeStop'].Current }
    if ($b['AdvBatteryChargeCfg'].Current -eq 'Enabled') { $text += Tr 'Bat.AdvOn' }
    if ($b['PeakShiftCfg'].Current -eq 'Enabled') { $text += Tr 'Bat.PeakOn' }
    $ui.BatMode.Text = $text
}

# ---------------------------------------------------------------- Onglet Mode de charge
function Get-ModeName {
    param([string]$Key)
    if ($ModeKeys -contains $Key) { Tr "Mode.$Key.Name" } else { $Key }
}

function Initialize-ModeList {
    param([string[]]$Possible)
    $ui.ModeList.Children.Clear()
    $script:ModeRadios = [ordered]@{}
    # Le BIOS peut renvoyer des entrées vides dans la liste des valeurs possibles : on les ignore.
    $Possible = @($Possible | Where-Object { $_ })
    $keys = @($ModeKeys | Where-Object { $Possible -contains $_ }) + @($Possible | Where-Object { $ModeKeys -notcontains $_ })
    foreach ($key in $keys) {
        $known = $ModeKeys -contains $key
        $title = New-Object Windows.Controls.TextBlock
        if ($known) { Set-Loc $title 'Text' "Mode.$key.Name" } else { $title.Text = $key }
        $title.FontWeight = 'SemiBold'
        $desc = New-Object Windows.Controls.TextBlock
        if ($known) { Set-Loc $desc 'Text' "Mode.$key.Desc" }
        $desc.Foreground = [Windows.Media.Brushes]::DimGray
        $desc.TextWrapping = 'Wrap'
        $content = New-Object Windows.Controls.StackPanel
        [void]$content.Children.Add($title)
        [void]$content.Children.Add($desc)

        $rb = New-Object Windows.Controls.RadioButton
        $rb.GroupName = 'ChargeMode'
        $rb.Tag = $key
        $rb.Margin = [Windows.Thickness]::new(0, 5, 0, 5)
        $rb.Content = $content
        $rb.Add_Checked({ Invoke-Guarded { Update-ModeTab } })
        [void]$ui.ModeList.Children.Add($rb)
        $script:ModeRadios[$key] = $rb
    }
}

function Get-SelectedMode {
    foreach ($rb in $script:ModeRadios.Values) { if ($rb.IsChecked) { return [string]$rb.Tag } }
    $null
}

function Set-Slider {
    param($Slider, $Attribute)
    if (-not $Attribute) { return }
    $Slider.Maximum = [double]$Attribute.Upper
    $Slider.Minimum = [double]$Attribute.Lower
    $Slider.Value = [double]$Attribute.Current
}

function Get-ModeChanges {
    $b = $script:Bios
    $changes = @(Get-ChargeWriteOrder -OldStart $b['CustomChargeStart'].Current -NewStart $ui.StartSlider.Value `
                                      -OldStop $b['CustomChargeStop'].Current -NewStop $ui.StopSlider.Value)
    $mode = Get-SelectedMode
    if ($mode -and $mode -ne $b['PrimaryBattChargeCfg'].Current) {
        # Le mode en dernier : passer en Personnalisé applique directement les nouveaux seuils.
        $changes += @{ Name = 'PrimaryBattChargeCfg'; Value = $mode }
    }
    , $changes
}

function Format-ModeChange {
    param($Change)
    $old = $script:Bios[$Change.Name].Current
    switch ($Change.Name) {
        'PrimaryBattChargeCfg' { Tr 'Change.Mode' (Get-ModeName $old) (Get-ModeName $Change.Value) }
        'CustomChargeStart' { Tr 'Change.Start' $old $Change.Value }
        'CustomChargeStop' { Tr 'Change.Stop' $old $Change.Value }
    }
}

function Update-ModeTab {
    $ui.StartValue.Text = Tr 'Fmt.Pct' ([int]$ui.StartSlider.Value)
    $ui.StopValue.Text = Tr 'Fmt.Pct' ([int]$ui.StopSlider.Value)
    if (-not $script:Bios) {
        $ui.ModeApply.IsEnabled = $false
        $ui.ModeRevert.IsEnabled = $false
        return
    }
    $mode = Get-SelectedMode
    $ui.ModeList.IsEnabled = -not $script:Writing
    $ui.CustomPanel.IsEnabled = ($mode -eq 'Custom') -and -not $script:Writing
    $changes = Get-ModeChanges
    $thresholdsChanged = @($changes | Where-Object { $_.Name -ne 'PrimaryBattChargeCfg' }).Count -gt 0
    $problem = ''
    if (($mode -eq 'Custom' -or $thresholdsChanged) -and ($ui.StopSlider.Value - $ui.StartSlider.Value) -lt 5) {
        $problem = Tr 'Mode.GapProblem'
    }
    $ui.ModeProblem.Text = $problem
    $ui.ModeApply.IsEnabled = $changes.Count -gt 0 -and -not $problem -and -not $script:Writing
    $ui.ModeRevert.IsEnabled = $changes.Count -gt 0 -and -not $script:Writing
}

function Invoke-ApplyMode {
    $changes = Get-ModeChanges
    if (-not $changes.Count) { return }
    $password = Get-BiosPassword
    $summary = ($changes | ForEach-Object { '•  ' + (Format-ModeChange $_) }) -join "`n"
    if (-not (Confirm-Write $summary)) { return }
    if ($Demo) { Set-Status 'Status.DemoNoWrite'; return }

    $script:Writing = $true
    Update-AllTabs
    Start-Background -Work $WorkApplyMode -Arguments @{ Steps = $changes; Password = $password } `
        -Status 'Status.Writing' -Context $changes -OnSuccess {
            param($attributes, $changes)
            Set-BiosState $attributes
            $pending = @($changes | Where-Object { $attributes[$_.Name].Current -ne [string]$_.Value })
            if ($pending.Count) {
                Set-Status 'Status.Pending' (($pending | ForEach-Object { $_.Name }) -join ', ')
            } else {
                Set-Status 'Status.Applied'
            }
        } -OnFinally { $script:Writing = $false; Update-AllTabs }
}

# ---------------------------------------------------------------- Onglets de planning
function New-ScheduleGrid {
    param($Grid, [string]$Which)
    $headers = @($SchedInfo[$Which].Headers)
    [void]$Grid.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition -Property @{ Width = [Windows.GridLength]::new(110) }))
    foreach ($h in $headers) {
        [void]$Grid.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition -Property @{ Width = [Windows.GridLength]::new(170) }))
    }
    for ($i = 0; $i -le $DayKeys.Count; $i++) { [void]$Grid.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition)) }

    for ($c = 0; $c -lt $headers.Count; $c++) {
        $tb = New-Object Windows.Controls.TextBlock
        Set-Loc $tb 'Text' $headers[$c]
        $tb.Foreground = [Windows.Media.Brushes]::DimGray
        $tb.Margin = [Windows.Thickness]::new(0, 0, 0, 4)
        [Windows.Controls.Grid]::SetColumn($tb, $c + 1)
        [void]$Grid.Children.Add($tb)
    }

    $rows = @{}
    $r = 1
    foreach ($d in $DayKeys) {
        $label = New-Object Windows.Controls.TextBlock
        Set-Loc $label 'Text' "Day.$d"
        $label.VerticalAlignment = 'Center'
        [Windows.Controls.Grid]::SetRow($label, $r)
        [void]$Grid.Children.Add($label)
        $boxes = @()
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $cb = New-Object Windows.Controls.ComboBox
            $cb.ItemsSource = $TimeChoices
            $cb.Tag = $Which
            $cb.Margin = [Windows.Thickness]::new(0, 2, 16, 2)
            $cb.Add_SelectionChanged({ param($s, $e) Invoke-Guarded { Update-ScheduleTab $s.Tag } })
            [Windows.Controls.Grid]::SetRow($cb, $r)
            [Windows.Controls.Grid]::SetColumn($cb, $c + 1)
            [void]$Grid.Children.Add($cb)
            $boxes += $cb
        }
        $rows[$d] = $boxes
        $r++
    }
    $rows
}

function Get-GridDays {
    param([string]$Which)
    $days = [ordered]@{}
    foreach ($d in $DayKeys) { $days[$d] = @($script:Rows[$Which][$d] | ForEach-Object { [string]$_.SelectedItem }) }
    $days
}

function Set-GridDays {
    param([string]$Which, $Days)
    foreach ($d in $DayKeys) {
        if (-not $Days.Contains($d)) { continue }
        $boxes = $script:Rows[$Which][$d]
        for ($i = 0; $i -lt $boxes.Count; $i++) { $boxes[$i].SelectedItem = [string]$Days[$d][$i] }
    }
}

function Get-ScheduleString {
    param([string]$Which, [switch]$Base)
    if ($Base) { return ConvertTo-CctkSchedule $script:Base[$Which].Enabled $script:Base[$Which].Days }
    $p = $SchedInfo[$Which].Prefix
    ConvertTo-CctkSchedule ([bool]$ui["${p}Enabled"].IsChecked) (Get-GridDays $Which)
}

function Test-ScheduleDirty {
    param([string]$Which)
    if (-not $script:Base.Contains($Which)) { return $false }
    $dirty = (Get-ScheduleString $Which) -ne (Get-ScheduleString $Which -Base)
    if ($Which -eq 'PeakShift' -and $script:Bios) {
        $dirty = $dirty -or ([int]$ui.PeakThreshold.Value -ne [int]$script:Bios['PeakShiftBatteryThreshold'].Current)
    }
    $dirty
}

function Update-ScheduleTab {
    param([string]$Which)
    $p = $SchedInfo[$Which].Prefix
    $loaded = $null -ne $script:SchedData
    $enabled = [bool]$ui["${p}Enabled"].IsChecked
    $editable = $loaded -and -not $script:Writing
    $ui["${p}Enabled"].IsEnabled = $editable
    $ui["${p}Grid"].IsEnabled = $editable -and $enabled
    $ui["${p}Copy"].IsEnabled = $editable -and $enabled
    $ui["${p}Read"].IsEnabled = -not $script:SchedLoading -and -not $script:Writing
    if ($Which -eq 'PeakShift') {
        $ui.PeakThreshold.IsEnabled = $editable -and ($null -ne $script:Bios)
        $ui.PeakThresholdValue.Text = Tr 'Fmt.Pct' ([int]$ui.PeakThreshold.Value)
    }
    $problems = @()
    $dirty = $false
    if ($loaded) {
        if ($enabled) { $problems = @(Get-ScheduleErrors (Get-GridDays $Which) $Which) }
        $dirty = Test-ScheduleDirty $Which
    }
    $ui["${p}Problem"].Text = $problems -join "`n"
    $ui["${p}Apply"].IsEnabled = $dirty -and -not $problems.Count -and -not $script:Writing
    $ui["${p}Revert"].IsEnabled = $dirty -and -not $script:Writing
}

function Update-AllTabs {
    Update-ModeTab
    Update-ScheduleTab 'Advanced'
    Update-ScheduleTab 'PeakShift'
}

$LastKnownFile = Join-Path $env:LOCALAPPDATA 'DellBatteryManager\plannings.json'

function Read-LastKnown {
    $known = @{}
    if (-not (Test-Path -LiteralPath $LastKnownFile)) { return $known }
    try {
        $json = Get-Content -Raw -LiteralPath $LastKnownFile | ConvertFrom-Json
        foreach ($which in 'Advanced', 'PeakShift') {
            if (-not $json.$which) { continue }
            $days = [ordered]@{}
            foreach ($prop in $json.$which.PSObject.Properties) { $days[$prop.Name] = @($prop.Value) }
            $known[$which] = $days
        }
    } catch { }
    $known
}

function Save-LastKnown {
    # Le BIOS ne renvoie pas forcément le planning d'un mode désactivé : on garde le dernier connu.
    param([string]$Which, $Days)
    $script:LastKnown[$Which] = $Days
    if ($Demo) { return }
    New-Item -ItemType Directory -Force -Path (Split-Path $LastKnownFile) | Out-Null
    $script:LastKnown | ConvertTo-Json -Depth 4 | Set-Content -Encoding UTF8 -LiteralPath $LastKnownFile
}

function Show-RawText {
    # Texte sous la grille : valeur brute lue par cctk et/ou remarque, dans la langue active.
    param([string]$Which)
    $info = $script:RawInfo[$Which]
    $lines = @()
    if ($info) {
        if ($null -ne $info.Raw) { $lines += Tr 'Sched.Raw' $info.Raw }
        if ($info.Note) { $lines += Tr $info.Note }
    }
    $ui["$($SchedInfo[$Which].Prefix)Raw"].Text = $lines -join "`n"
}

function Show-Schedule {
    param([string]$Which)
    $info = $SchedInfo[$Which]
    $p = $info.Prefix
    $read = $script:SchedData.$Which
    $note = ''
    if ($read.Days.Count) {
        $days = $read.Days
    } elseif ($script:LastKnown.Contains($Which)) {
        $days = $script:LastKnown[$Which]
        $note = 'Sched.NoteLastKnown'
    } else {
        $days = [ordered]@{}
        foreach ($d in $DayKeys) { $days[$d] = $info.Default }
        $note = 'Sched.NoteDefault'
    }
    $script:Base[$Which] = @{ Enabled = $read.Enabled; Days = $days }
    $ui["${p}Enabled"].IsChecked = $read.Enabled
    Set-GridDays $Which $days
    if ($Which -eq 'PeakShift' -and $script:Bios) { Set-Slider $ui.PeakThreshold $script:Bios['PeakShiftBatteryThreshold'] }
    $script:RawInfo[$Which] = @{ Raw = $read.Raw; Note = $note }
    Show-RawText $Which
    Update-ScheduleTab $Which
}

function Start-ReadSchedules {
    if ($script:SchedLoading) { return }
    if ($Demo) { Show-Schedule 'Advanced'; Show-Schedule 'PeakShift'; return }
    if (-not (Update-CctkBanner)) {
        # Sans cctk, inutile de lancer la lecture : la bannière explique comment installer l'outil.
        foreach ($which in 'Advanced', 'PeakShift') { $script:RawInfo[$which] = @{ Raw = $null; Note = 'Sched.NoCctk' }; Show-RawText $which }
        Set-Status 'Sched.NoCctk'
        return
    }
    $script:SchedLoading = $true
    Update-AllTabs
    Start-Background -Work $WorkReadSchedules -Status 'Status.ReadingSched' -OnSuccess {
        param($schedules)
        $script:SchedData = $schedules
        foreach ($which in 'Advanced', 'PeakShift') {
            if ($schedules.$which.Days.Count) { Save-LastKnown $which $schedules.$which.Days }
            Show-Schedule $which
        }
        Set-Status 'Status.SchedRead' (Get-Date)
    } -OnFinally { $script:SchedLoading = $false; Update-AllTabs }
}

function Format-ScheduleSummary {
    param([string]$Which, $Days)
    foreach ($d in $DayKeys) {
        $t = $Days[$d]
        if ($Which -eq 'Advanced') { Tr 'Sched.SumAdv' (Tr "Day.$d") $t[0] $t[1] }
        else { Tr 'Sched.SumPeak' (Tr "Day.$d") $t[0] $t[1] $t[2] }
    }
}

function Invoke-ApplySchedule {
    param([string]$Which)
    $info = $SchedInfo[$Which]
    $p = $info.Prefix
    $enabled = [bool]$ui["${p}Enabled"].IsChecked
    $days = Get-GridDays $Which
    $new = Get-ScheduleString $Which
    $arguments = @{ Advanced = ''; PeakShift = ''; Threshold = $null; Password = (Get-BiosPassword) }
    $lines = @()
    if ($new -ne (Get-ScheduleString $Which -Base)) {
        $arguments[$Which] = $new
        if ($enabled) { $lines += Tr 'Sched.ChangeOn' (Tr $info.Title); $lines += Format-ScheduleSummary $Which $days }
        else { $lines += Tr 'Sched.ChangeOff' (Tr $info.Title) }
    }
    if ($Which -eq 'PeakShift') {
        $old = [int]$script:Bios['PeakShiftBatteryThreshold'].Current
        $threshold = [int]$ui.PeakThreshold.Value
        if ($threshold -ne $old) { $arguments.Threshold = $threshold; $lines += Tr 'Sched.ChangeThresh' $old $threshold }
    }
    if (-not $lines) { return }
    if (-not (Confirm-Write ($lines -join "`n"))) { return }
    if ($Demo) { Set-Status 'Status.DemoNoWrite'; return }

    $script:Writing = $true
    Update-AllTabs
    Start-Background -Work $WorkApplySchedule -Arguments $arguments -Status 'Status.WritingSched' `
        -Context @{ Which = $Which; Enabled = $enabled; Days = $days } -OnSuccess {
            param($result, $ctx)
            if ($ctx.Enabled) { Save-LastKnown $ctx.Which $ctx.Days }
            $other = if ($ctx.Which -eq 'Advanced') { 'PeakShift' } else { 'Advanced' }
            $otherDirty = Test-ScheduleDirty $other
            $script:SchedData = $result.Schedules
            Set-BiosState $result.Attributes
            Show-Schedule $ctx.Which
            if (-not $otherDirty) { Show-Schedule $other }
            if ($result.Schedules.($ctx.Which).Enabled -ne $ctx.Enabled) {
                Set-Status 'Status.SchedDiffers'
            } else {
                Set-Status 'Status.SchedApplied'
            }
        } -OnFinally { $script:Writing = $false; Update-AllTabs }
}

function Copy-Monday {
    param([string]$Which)
    $rows = $script:Rows[$Which]
    foreach ($d in 'Tue', 'Wed', 'Thu', 'Fri') {
        for ($i = 0; $i -lt $rows['Mon'].Count; $i++) { $rows[$d][$i].SelectedItem = $rows['Mon'][$i].SelectedItem }
    }
}

# ---------------------------------------------------------------- Lecture BIOS
function Set-BiosState {
    param($Attributes)
    $script:Bios = $Attributes
    if (-not $script:ModeRadios) { Initialize-ModeList @($Attributes['PrimaryBattChargeCfg'].Possible) }
    $current = $Attributes['PrimaryBattChargeCfg'].Current
    foreach ($key in @($script:ModeRadios.Keys)) { $script:ModeRadios[$key].IsChecked = ($key -eq $current) }
    Set-Slider $ui.StartSlider $Attributes['CustomChargeStart']
    Set-Slider $ui.StopSlider $Attributes['CustomChargeStop']
    Set-Slider $ui.PeakThreshold $Attributes['PeakShiftBatteryThreshold']
    if (-not $script:SchedData) {
        # Interrupteurs connus par WMI avant la lecture (lente) des plannings par cctk.
        $ui.AdvEnabled.IsChecked = $Attributes['AdvBatteryChargeCfg'].Current -eq 'Enabled'
        $ui.PeakEnabled.IsChecked = $Attributes['PeakShiftCfg'].Current -eq 'Enabled'
    }
    Show-ActiveMode
    Update-AllTabs
}

function Start-ReadBios {
    if ($Demo) { return }
    Start-Background -Work $WorkReadBios -Status 'Status.ReadingBios' -OnSuccess {
        param($r)
        $script:PasswordRequired = [bool]$r.PasswordRequired
        $ui.PwdPanel.Visibility = if ($script:PasswordRequired) { 'Visible' } else { 'Collapsed' }
        Set-BiosState $r.Attributes
        Set-Status 'Status.BiosRead' (Get-Date)
    }
}

# ---------------------------------------------------------------- Bannière Dell Command | Configure
function Update-CctkBanner {
    # Affiche la bannière si cctk.exe est introuvable ; renvoie $true si l'outil est présent.
    $script:CctkPath = Find-Cctk
    $ui.CctkBanner.Visibility = if ($script:CctkPath) { 'Collapsed' } else { 'Visible' }
    [bool]$script:CctkPath
}

function Open-DellDownloadPage {
    $url = Tr 'Banner.Url'
    # L'app tourne en administrateur : passer par l'Explorateur ouvre le navigateur avec les droits normaux.
    try { Start-Process -FilePath "$env:SystemRoot\explorer.exe" -ArgumentList "`"$url`"" }
    catch { Start-Process $url }
}

function Invoke-RecheckCctk {
    if (Update-CctkBanner) {
        Set-Status 'Banner.Found'
        $tab = $ui.Tabs.SelectedItem
        if (($tab -eq $ui.TabAdv -or $tab -eq $ui.TabPeak) -and -not $script:SchedData) { Start-ReadSchedules }
    } else {
        Set-Status 'Banner.StillMissing'
    }
}

# ---------------------------------------------------------------- Changement de langue
function Set-Language {
    param([string]$Lang, [switch]$Save)
    $global:DbmShared.Lang = $Lang
    foreach ($item in $script:Localized) { $item.Element.($item.Property) = Tr $item.Key }
    $window.Title = 'Dell Battery Manager' + $(if ($Demo) { Tr 'App.DemoSuffix' })
    $window.Language = [Windows.Markup.XmlLanguage]::GetLanguage($(if ($Lang -eq 'en') { 'en-US' } else { 'fr-FR' }))
    $ui.MenuLangFr.IsChecked = $Lang -eq 'fr'
    $ui.MenuLangEn.IsChecked = $Lang -eq 'en'
    # Textes calculés : on les recalcule à partir de l'état courant (sans toucher aux modifications en cours).
    Show-BatteryLive $null
    if ($script:Bios) { Show-ActiveMode }
    foreach ($which in 'Advanced', 'PeakShift') { Show-RawText $which }
    Update-AllTabs
    Update-Status
    if ($Save) { Save-Language $Lang }
}

# ---------------------------------------------------------------- Événements
Register-LocalizedTree $window
$script:Rows = @{
    Advanced  = (New-ScheduleGrid $ui.AdvGrid 'Advanced')
    PeakShift = (New-ScheduleGrid $ui.PeakGrid 'PeakShift')
}
$script:LastKnown = Read-LastKnown

function Invoke-Refresh {
    [void](Update-CctkBanner)
    Start-ReadBios
    Start-BatteryRefresh
}

$ui.BtnRefresh.Add_Click({ Invoke-Guarded { Invoke-Refresh } })
$ui.MenuRefresh.Add_Click({ Invoke-Guarded { Invoke-Refresh } })
$window.Add_KeyDown({ param($s, $e) if ($e.Key -eq 'F5') { $e.Handled = $true; Invoke-Guarded { Invoke-Refresh } } })
$ui.MenuQuit.Add_Click({ $window.Close() })
$ui.MenuLangFr.Add_Click({ Invoke-Guarded { Set-Language 'fr' -Save } })
$ui.MenuLangEn.Add_Click({ Invoke-Guarded { Set-Language 'en' -Save } })
$ui.BannerLink.Add_Click({ Invoke-Guarded { Open-DellDownloadPage } })
$ui.BannerRecheck.Add_Click({ Invoke-Guarded { Invoke-RecheckCctk } })
$ui.StartSlider.Add_ValueChanged({ Invoke-Guarded { Update-ModeTab } })
$ui.StopSlider.Add_ValueChanged({ Invoke-Guarded { Update-ModeTab } })
$ui.ModeApply.Add_Click({ Invoke-Guarded { Invoke-ApplyMode } })
$ui.ModeRevert.Add_Click({ Invoke-Guarded { Set-BiosState $script:Bios } })
$ui.PeakThreshold.Add_ValueChanged({ Invoke-Guarded { Update-ScheduleTab 'PeakShift' } })

foreach ($which in 'Advanced', 'PeakShift') {
    $p = $SchedInfo[$which].Prefix
    foreach ($suffix in 'Enabled', 'Read', 'Copy', 'Revert', 'Apply') { $ui["$p$suffix"].Tag = $which }
    $ui["${p}Enabled"].Add_Checked({ param($s, $e) Invoke-Guarded { Update-ScheduleTab $s.Tag } })
    $ui["${p}Enabled"].Add_Unchecked({ param($s, $e) Invoke-Guarded { Update-ScheduleTab $s.Tag } })
    $ui["${p}Read"].Add_Click({ Invoke-Guarded { Start-ReadSchedules } })
    $ui["${p}Copy"].Add_Click({ param($s, $e) Invoke-Guarded { Copy-Monday $s.Tag } })
    $ui["${p}Revert"].Add_Click({ param($s, $e) Invoke-Guarded { Show-Schedule $s.Tag } })
    $ui["${p}Apply"].Add_Click({ param($s, $e) Invoke-Guarded { Invoke-ApplySchedule $s.Tag } })
}

$ui.Tabs.Add_SelectionChanged({
    param($s, $e)
    if ($e.OriginalSource -ne $ui.Tabs) { return }   # les ComboBox des grilles remontent aussi cet événement
    Invoke-Guarded {
        $tab = $ui.Tabs.SelectedItem
        if (($tab -eq $ui.TabAdv -or $tab -eq $ui.TabPeak) -and -not $script:SchedData) { Start-ReadSchedules }
    }
})

$jobTimer = New-Object Windows.Threading.DispatcherTimer
$jobTimer.Interval = [TimeSpan]::FromMilliseconds(150)
$jobTimer.Add_Tick({
    foreach ($job in @($script:Jobs)) {
        if (-not $job.Handle.IsCompleted) { continue }
        $script:Jobs.Remove($job)
        try {
            $output = $job.Ps.EndInvoke($job.Handle)
            if ($job.Ps.Streams.Error.Count) { throw $job.Ps.Streams.Error[0] }
            $data = if ($output.Count) { $output[0] } else { $null }
            if ($job.OnSuccess) { & $job.OnSuccess $data $job.Context }
        } catch {
            $origin = "tâche de fond $($job.Status) (OnSuccess ou tâche)"
            if ($job.Quiet) { Write-ErrorLog $_ $origin; Set-Status 'Status.Error' (Get-InnermostMessage $_) } else { Show-Error $_ $origin }
        } finally {
            $job.Ps.Dispose()
            if ($job.OnFinally) { & $job.OnFinally $job.Context }
        }
    }
    Update-Status
})

$batteryTimer = New-Object Windows.Threading.DispatcherTimer
$batteryTimer.Interval = [TimeSpan]::FromSeconds(15)
$batteryTimer.Add_Tick({ Invoke-Guarded { Start-BatteryRefresh } })

$window.Add_Closing({
    param($s, $e)
    if ($script:Writing) {
        $e.Cancel = $true
        [void][Windows.MessageBox]::Show($window, (Tr 'Dlg.WriteInProgress'), 'Dell Battery Manager', 'OK', 'Warning')
    }
})
$window.Add_Closed({ $jobTimer.Stop(); $batteryTimer.Stop(); $pool.Close() })

# ---------------------------------------------------------------- Mode démo
Set-Language $global:DbmShared.Lang

if ($Demo) {
    $demoAttribute = { param($name, $current, $possible, $lower, $upper)
        [pscustomobject]@{ Name = $name; Current = $current; Possible = @($possible); Lower = $lower; Upper = $upper } }
    Set-BiosState @{
        PrimaryBattChargeCfg      = & $demoAttribute 'PrimaryBattChargeCfg' 'Custom' @('Adaptive', 'Standard', 'Express', 'PrimAcUse', 'Custom')
        CustomChargeStart         = & $demoAttribute 'CustomChargeStart' '75' @() 50 95
        CustomChargeStop          = & $demoAttribute 'CustomChargeStop' '90' @() 55 100
        AdvBatteryChargeCfg       = & $demoAttribute 'AdvBatteryChargeCfg' 'Disabled' @('Disabled', 'Enabled')
        PeakShiftCfg              = & $demoAttribute 'PeakShiftCfg' 'Enabled' @('Disabled', 'Enabled')
        PeakShiftBatteryThreshold = & $demoAttribute 'PeakShiftBatteryThreshold' '15' @() 15 100
    }
    $script:SchedData = [pscustomobject]@{
        Advanced  = ConvertFrom-CctkSchedule 'Disabled' 2
        PeakShift = ConvertFrom-CctkSchedule 'Enabled,Mon-16:00/20:00/22:00,Tue-16:00/20:00/22:00,Wed-16:00/20:00/22:00,Thu-16:00/20:00/22:00,Fri-16:00/20:00/22:00,Sat-00:00/00:00/00:00,Sun-00:00/00:00/00:00' 3
    }
    Show-Schedule 'Advanced'
    Show-Schedule 'PeakShift'
    Set-Status 'Status.Demo'

    if ($Snapshot) {
        $window.WindowStartupLocation = 'Manual'
        $window.Left = -20000
        $window.ShowInTaskbar = $false
        $snapTimer = New-Object Windows.Threading.DispatcherTimer
        $snapTimer.Interval = [TimeSpan]::FromSeconds(5)   # laisse le temps aux infos batterie d'arriver
        $snapTimer.Add_Tick({
            $snapTimer.Stop()
            for ($i = 0; $i -lt $ui.Tabs.Items.Count; $i++) {
                $ui.Tabs.SelectedIndex = $i
                $window.UpdateLayout()
                $window.Dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Render, [Action] { })
                $content = $window.Content
                $bitmap = New-Object Windows.Media.Imaging.RenderTargetBitmap -ArgumentList ([int]$content.ActualWidth), ([int]$content.ActualHeight), 96, 96, ([Windows.Media.PixelFormats]::Pbgra32)
                $bitmap.Render($content)
                $encoder = New-Object Windows.Media.Imaging.PngBitmapEncoder
                $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
                $stream = [IO.File]::Create(('{0}-{1}.png' -f $Snapshot, ($i + 1)))
                try { $encoder.Save($stream) } finally { $stream.Dispose() }
            }
            $window.Close()
        })
        $window.Add_ContentRendered({ $snapTimer.Start() })
    }
}

# ---------------------------------------------------------------- Lancement
$jobTimer.Start()
$batteryTimer.Start()
[void](Update-CctkBanner)
Start-ReadBios
Start-BatteryRefresh
Start-Background -Work $WorkReport -Quiet -OnSuccess { param($r) $script:Report = $r; Show-BatteryLive $null }
Update-AllTabs
if (-not $Demo) { Set-Status 'Status.Starting' }
[void]$window.ShowDialog()
