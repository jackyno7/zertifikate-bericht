<#
.SYNOPSIS
    Certificate report: reads certificates on Windows servers and reports expired, soon-to-expire, and
    weak certificates, especially those actually used by web services or Remote Desktop. The result
    is an HTML report.

    Deutsch:
    Zertifikate-Bericht: liest die Zertifikate auf Windows-Servern aus und meldet abgelaufene, bald ablaufende und
    schwache Zertifikate, besonders solche, die von Webdiensten oder Remotedesktop tatsächlich verwendet werden. Ergebnis
    ist ein HTML-Bericht.

.DESCRIPTION
    The script operates exclusively READ-ONLY. It reads the certificate stores "My", "Remote Desktop"
    and "WebHosting" of the computer, the mapping of certificates to HTTPS bindings (HTTP.sys) and to Remote Desktop in
    the registry. It does not install, delete, or renew certificates, does not change bindings, and does not transfer
    data. Private keys are not read, only whether one is present.

    It always checks the computer on which it runs (except with -OhneLokalen). With -Computer or -AusAD, it queries additional
    computers via PowerShell remote management (WinRM). WinRM must be reachable there; the script does not change settings
    to achieve this. Unreachable computers are listed in the report as "not checked".

    Certificates bound to services without entries in HTTP.sys or Remote Desktop are not captured, for
    example SMTP, IMAP, and POP from Exchange, LDAPS, or services with their own certificate store. They appear in the
    list, but without the note "in use". Bindings to hostnames (SNI) are read; the central certificate store
    (CCS) of IIS is not evaluated, its use is reported as "not checked".

    Prerequisite for remote querying is an account that is a member of local Administrators or
    "Remote Management Users" on the target. During a remote query, certificate data leave the target computer on the way to
    the computer on which the script runs.

    Deutsch:
    Das Skript arbeitet ausschließlich LESEND. Es liest die Zertifikatspeicher "Eigene Zertifikate", "Remotedesktop"
    und "WebHosting" des Computers, die Zuordnung von Zertifikaten zu HTTPS-Bindungen (HTTP.sys) und zu Remotedesktop in
    der Registrierung. Es installiert, löscht oder erneuert keine Zertifikate, ändert keine Bindungen und überträgt keine
    Daten. Private Schlüssel werden nicht gelesen, nur ob einer vorhanden ist.

    Es prüft immer den Computer, auf dem es läuft (außer mit -OhneLokalen). Mit -Computer oder -AusAD fragt es weitere
    Computer über die PowerShell-Fernverwaltung (WinRM) ab. Dazu muss WinRM dort erreichbar sein, das Skript ändert keine Einstellungen,
    um das zu erreichen. Nicht erreichbare Computer stehen im Bericht als "nicht geprüft".

    Nicht erfasst werden Zertifikate, die an Dienste ohne Eintrag in HTTP.sys oder Remotedesktop gebunden sind, zum
    Beispiel SMTP, IMAP und POP von Exchange, LDAPS oder Dienste mit eigenem Zertifikatspeicher. Sie erscheinen in der
    Liste, aber ohne den Vermerk "in Gebrauch". Bindungen an Hostnamen (SNI) werden gelesen; der zentrale Zertifikatspeicher
    (CCS) von IIS wird nicht ausgewertet, sein Einsatz wird als "nicht geprüft" gemeldet.

    Voraussetzung für die Fernabfrage ist ein Konto, das auf dem Ziel Mitglied der lokalen Administratoren oder der
    "Remoteverwaltungsbenutzer" ist. Bei einer Fernabfrage verlassen die Zertifikatsdaten den Zielrechner auf dem Weg zu
    dem Computer, auf dem das Skript läuft.

.PARAMETER Ausgabe
    Path of the HTML file. Default: Zertifikate-Bericht_<Computer>_<Datum>.html in the current folder.

    Deutsch:
    Pfad der HTML-Datei. Standard: Zertifikate-Bericht_<Computer>_<Datum>.html im aktuellen Ordner.

.PARAMETER Computer
    Names of additional computers queried via WinRM.

    Deutsch:
    Namen weiterer Computer, die über WinRM abgefragt werden.

.PARAMETER AusAD
    Queries all active Windows servers from Active Directory (ActiveDirectory module required). Servers are
    computer accounts with "Server" in the operating system that have logged on in the last 60 days.

    Deutsch:
    Fragt alle aktiven Windows-Server aus dem Active Directory ab (Modul ActiveDirectory nötig). Als Server gelten
    Computerkonten mit "Server" im Betriebssystem, die sich in den letzten 60 Tagen angemeldet haben.

.PARAMETER OhneLokalen
    Do not check the computer on which the script runs (useful when started from a workstation).

    Deutsch:
    Den Computer, auf dem das Skript läuft, nicht prüfen (sinnvoll, wenn es von einem Arbeitsplatz aus gestartet wird).

.PARAMETER Server
    Specific domain controller for the AD query (only with -AusAD).

    Deutsch:
    Bestimmter Domänencontroller für die AD-Abfrage (nur mit -AusAD).

.PARAMETER Anmeldung
    Credentials (Get-Credential) for remote querying. Default: the current account.

    Deutsch:
    Anmeldedaten (Get-Credential) für die Fernabfrage. Standard: das aktuelle Konto.

.PARAMETER WarnTage
    Certificates that expire in this number of days are reported (default 30).

    Deutsch:
    Zertifikate, die in so vielen Tagen ablaufen, werden gemeldet (Standard 30).

.PARAMETER Csv
    Additionally, for each checkpoint with matches, writes a CSV file and the list of all certificates alongside the report.

    Deutsch:
    Schreibt zusätzlich je Prüfpunkt mit Treffern eine CSV-Datei und die Liste aller Zertifikate neben den Bericht.

.PARAMETER NichtOeffnen
    Does not automatically open the completed report in the browser.

    Deutsch:
    Öffnet den fertigen Bericht nicht automatisch im Browser.

.PARAMETER PassThru
    Additionally returns the findings as objects (for custom analysis).

    Deutsch:
    Gibt die Befunde zusätzlich als Objekte zurück (für eigene Auswertungen).

.PARAMETER Language
    Language of the report and the console output: 'de' (German) or 'en' (English). Default: German if the Windows display language
    is German, otherwise English.

    Deutsch:
    Sprache von Bericht und Konsolenausgabe: 'de' (Deutsch) oder 'en' (Englisch). Standard: Deutsch, wenn die Anzeigesprache von
    Windows Deutsch ist, sonst Englisch.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Zertifikate-Bericht.ps1

.EXAMPLE
    .\Zertifikate-Bericht.ps1 -AusAD -WarnTage 60 -Csv -Ausgabe C:\Berichte\zertifikate.html

.NOTES
    IT Admin Tools by MM - Zertifikate-Bericht Version 1.0 (11.2026), Lizenz: MIT (siehe LICENSE)
    Kein Ersatz für eine Zertifikatsverwaltung. Nutzung auf eigene Verantwortung.
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('de', 'en')][string]$Language = $(if ((Get-UICulture).TwoLetterISOLanguageName -eq 'de') { 'de' } else { 'en' }),
    [Alias('Output')][string]$Ausgabe,
    [string[]]$Computer,
    [switch]$AusAD,
    [switch]$OhneLokalen,
    [string]$Server,
    [Alias('Credential')][System.Management.Automation.PSCredential]$Anmeldung,
    [ValidateRange(1, 730)][int]$WarnTage = 30,
    [switch]$Csv,
    [Alias('NoOpen')][switch]$NichtOeffnen,
    [switch]$PassThru
)
# Sprache des Berichts: de oder en (Parameter -Language). T 'deutsch' 'english' liefert den Text der gewählten Sprache.
$En = ($Language -eq 'en')
function T([string]$de, [string]$en) { if ($script:En) { $en } else { $de } }

$ErrorActionPreference = 'Stop'
$Version = '1.1'
# Kommalisten: Beim Start über "powershell -File" kommt "a,b" als ein Wert an; hier in einzelne Namen trennen.
if ($PSBoundParameters.ContainsKey('Computer')) {
    $Computer = @($Computer | ForEach-Object { $_ -split '\s*,\s*' } | Where-Object { $_ })
    if (-not $Computer.Count) { throw (T 'Bei -Computer wurde kein Computer angegeben. Ohne den Parameter wird nur der lokale Computer geprüft.' 'No computer was specified for -Computer. Without the parameter, only the local computer is checked.') }
}

#@SPRACHE@
# ------------------------------------------------------------------ Hilfsfunktionen
function Datum($d, $leer = $null) { if ($null -eq $leer) { $leer = T 'nie' 'never' }; if ($d) { ([datetime]$d).ToString((T 'dd.MM.yyyy' 'yyyy-MM-dd')) } else { $leer } }
function JaNein($b) { if ($b) { T 'Ja' 'Yes' } else { T 'Nein' 'No' } }
function Ort($dn) { ($dn -replace '^(CN|OU)=.+?(?<!\\),', '') -replace ',DC=.*$', '' }
function Html($t) { [System.Net.WebUtility]::HtmlEncode([string]$t) }
function Kurz($t, $n = 220) { $t = ([string]$t -replace '\s+', ' ').Trim(); if ($t.Length -gt $n) { $t.Substring(0, $n) + ' ...' } else { $t } }

# CSV-Dateien öffnen sich in Excel: Werte, die mit = + - @ beginnen, würden dort als Formel ausgeführt (Namen von Dateien, Ordnern, Konten oder Computern
# stammen von Benutzern). Diese Funktion ersetzt für dieses Skript den Befehl Export-Csv und setzt vor solche Werte ein Hochkomma.
function Export-Csv {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline = $true)]$InputObject, [Parameter(Mandatory)][string]$Path, [switch]$NoTypeInformation, [string]$Delimiter = ';', [string]$Encoding = 'UTF8')
    begin { $alle = New-Object System.Collections.Generic.List[object] }
    process { if ($null -ne $InputObject) { $alle.Add($InputObject) } }
    end {
        $sicher = foreach ($o in $alle) {
            $n = [ordered]@{}
            foreach ($p in $o.PSObject.Properties) { $v = $p.Value; if (($v -is [string]) -and ($v -match '^[=+@\t\r]|^-.')) { $v = "'" + $v }; $n[$p.Name] = $v }
            [pscustomobject]$n
        }
        if ($alle.Count) { $sicher | Microsoft.PowerShell.Utility\Export-Csv -Path $Path -NoTypeInformation:$NoTypeInformation -Delimiter $Delimiter -Encoding $Encoding }   # leere Liste: keine Datei (sie hätte keine Kopfzeile)
    }
}

# Tabellen: ConvertTo-Html und Export-Csv nehmen die Spalten des ersten Objekts. Haben die Zeilen einer Liste verschiedene
# Eigenschaften (z. B. Fehlzeile neben normalen Treffern), gingen Spalten verloren: alle Namen der Reihe nach vereinigen.
function Einheitlich($liste) {
    $liste = @($liste | Where-Object { $null -ne $_ })
    if ($liste.Count -ge 2) {
        $namen = New-Object System.Collections.Generic.List[string]
        foreach ($o in $liste) { foreach ($p in $o.PSObject.Properties.Name) { if (-not $namen.Contains($p)) { $namen.Add($p) } } }
        if ($namen.Count -ne @($liste[0].PSObject.Properties.Name).Count) { $liste = @($liste | Select-Object -Property $namen) }
    }
    SpaltenAnzeige $liste
}
# Englische Berichte: Spaltenüberschriften aus der Übersetzungstabelle (intern bleiben die deutschen Namen, der Code liest sie weiter)
$SpaltenEn = @{
    'Abfrage' = 'Query'
    'Administrator' = 'Administrator'
    'Adresse' = 'Address'
    'Aktiv' = 'Active'
    'Aktiv verknüpft mit' = 'Active linked to'
    'Aktive Verknüpfungen' = 'Active links'
    'Alter in Tagen' = 'Age in days'
    'Angelegt' = 'Created'
    'Anmeldung' = 'Last logon'
    'Anmeldungen' = 'Logons'
    'Anteil (%)' = 'Share (%)'
    'Anzahl' = 'Count'
    'Art' = 'Account type'
    'Auf Domänencontroller' = 'On domain controller'
    'Aufräumen' = 'Cleanup'
    'Auslastung' = 'Utilization'
    'Auslösender Rechner' = 'Triggering computer'
    'Auslösender Rechner (Anzahl)' = 'Triggering computers (count)'
    'Aussteller' = 'Issuer'
    'Beispiele' = 'Examples'
    'Belegt' = 'Used'
    'Belegt (%)' = 'Used (%)'
    'Benutzer und Computer (auch verschachtelt)' = 'Users and computers (also nested)'
    'Benutzereinstellungen' = 'User settings'
    'Benutzername' = 'Username'
    'Berechtigter' = 'Principal'
    'Bereich' = 'Scope'
    'Bereiche' = 'Scopes'
    'Beschreibung' = 'Description'
    'Besitzer' = 'Owner'
    'Betriebssystem' = 'OS'
    'Betroffene Konten' = 'Affected accounts'
    'Bewertung' = 'Rating'
    'Bis' = 'To'
    'BitLocker' = 'BitLocker'
    'Computer' = 'Computer'
    'Computereinstellungen' = 'Computer settings'
    'Darf anwenden' = 'Can apply'
    'Darf delegieren an' = 'Can delegate to'
    'Darf sich hier als Benutzer ausgeben' = 'Can log on as user'
    'Datei' = 'File'
    'Dateien' = 'Files'
    'Dateisystem' = 'File system'
    'Dateityp' = 'File type'
    'Datenbank' = 'Database'
    'Detail' = 'Detail'
    'Details' = 'Details'
    'Direkte Mitglieder' = 'Direct members'
    'Domäne' = 'Domain'
    'Domänencontroller' = 'Domain controller'
    'Domänennamenmaster' = 'Domain naming master'
    'Dynamische Updates' = 'Dynamic updates'
    'E-Mail' = 'Email'
    'Eigenschaft' = 'Property'
    'Einstellung' = 'Setting'
    'Empfohlen' = 'Recommended'
    'Enthaltene Gruppen' = 'Included groups'
    'Ereignisse' = 'Events'
    'Erreichbar' = 'Reachable'
    'Erste' = 'First'
    'Erster' = 'First'
    'Erzwungen' = 'Forced'
    'Failover' = 'Failover'
    'Falsches Kennwort' = 'Wrong password'
    'Fehlanmeldungen am DC' = 'Failed logons at DC'
    'Fehlversuche' = 'Failed attempts'
    'Feststellung' = 'Finding'
    'Fingerabdruck' = 'Fingerprint'
    'Firewall' = 'Firewall'
    'Frei' = 'Free'
    'Freigabe' = 'Share'
    'Freigaberecht' = 'Share permission'
    'Funktion' = 'Function'
    'Gelesen von' = 'Read by'
    'Geprüft' = 'Checked'
    'Gesamtgröße' = 'Total size'
    'Geschützte Benutzer' = 'Protected users'
    'Gesperrt seit' = 'Locked since'
    'Geändert' = 'Changed'
    'Gilt' = 'Applies to'
    'Gilt für' = 'Applies to'
    'Globaler Katalog' = 'Global catalog'
    'Grenze' = 'Limit'
    'Groesse' = 'Size'
    'Grund' = 'Reason'
    'Gruppe' = 'Group'
    'Gruppen' = 'Groups'
    'Gruppen (auch verschachtelt)' = 'Groups (including nested)'
    'Gruppenrichtlinie' = 'Group policy'
    'Größe' = 'Size'
    'Größe (GB)' = 'Size (GB)'
    'Gültig ab' = 'Valid from'
    'Gültig bis' = 'Valid to'
    'Herkunft' = 'Origin'
    'Hersteller' = 'Vendor'
    'Hinweis' = 'Note'
    'Häufigster Grund' = 'Most common reason'
    'IP-Adresse' = 'IP address'
    'Im AD gespeichert' = 'Stored in AD'
    'In der Domäne vorhanden' = 'Present in domain'
    'Infrastrukturmaster' = 'Infrastructure master'
    'Inhalt' = 'Content'
    'Ist' = 'Is'
    'Kennung' = 'ID'
    'Kennwort gesetzt' = 'Password set'
    'Kennwort läuft nie ab' = 'Password never expires'
    'Kerberos' = 'Kerberos'
    'Konten' = 'Accounts'
    'Konten (auch verschachtelt)' = 'Accounts (also nested)'
    'Konto' = 'Account'
    'Konto für DNS-Einträge' = 'Account for DNS entries'
    'Kopie bleibt im Postfach' = 'Copy remains in mailbox'
    'LAPS' = 'LAPS'
    'LAPS-Ablauf' = 'LAPS expiry'
    'Laufwerk' = 'Drive'
    'Lease-Dauer' = 'Lease duration'
    'Leitet weiter an' = 'Forwards to'
    'Leseweg' = 'Read path'
    'Letzte Aktualisierung' = 'Last update'
    'Letzte Anmeldung' = 'Last logon'
    'Letzte Vollsicherung' = 'Last full backup'
    'Letzte inkrementelle Sicherung' = 'Last incremental backup'
    'Letzter' = 'Last'
    'Letzter Aufräumlauf' = 'Last cleanup run'
    'Letzter Start' = 'Last start'
    'Lokale Admins' = 'Local admins'
    'Mindestlänge' = 'Minimum length'
    'Mitglied' = 'Member'
    'Mitglieder' = 'Members'
    'NLA' = 'NLA'
    'NTLM' = 'NTLM'
    'Name' = 'Name'
    'Nr.' = 'No.'
    'Objekt' = 'Object'
    'Ordner' = 'Folder'
    'Organisationseinheit' = 'OU'
    'Ort' = 'Location'
    'PDC-Emulator' = 'PDC emulator'
    'Pfad' = 'Path'
    'Postfach' = 'Mailbox'
    'Postfach-Typ' = 'Mailbox type'
    'Postfachtyp' = 'Mailbox type'
    'Postfächer' = 'Mailboxes'
    'Privilegierte Gruppen' = 'Privileged groups'
    'Programm' = 'Program'
    'Prüfpunkt' = 'Checkpoint'
    'Quelle' = 'Source'
    'Quelle (Anzahl)' = 'Sources (count)'
    'RID-Master' = 'RID master'
    'Recht' = 'Permission'
    'Rechte' = 'Permissions'
    'Regel' = 'Rule'
    'Reihenfolge' = 'Order'
    'SPN' = 'SPN'
    'Schemamaster' = 'Schema master'
    'Schlüssel' = 'Key'
    'Schreibgeschützt' = 'Read-only'
    'Sendesperre ab' = 'Send blocked from'
    'Server' = 'Server'
    'Sid' = 'SID'
    'Signatur' = 'Signature'
    'Speicher' = 'Storage'
    'Sperrungen' = 'Lockouts'
    'Stand' = 'Status'
    'Standort' = 'Site'
    'Status' = 'Status'
    'Support-Ende' = 'End of support'
    'Support-Ende (spätestes)' = 'End of support (latest)'
    'Tage bis Ablauf' = 'Days until expiry'
    'Tage her' = 'Days ago'
    'Tage ohne Anmeldung' = 'Days without logon'
    'Tage seit Ablauf' = 'Days since expiry'
    'Treffer' = 'Hits'
    'Typ' = 'Type'
    'Vererbung' = 'Inheritance'
    'Verschiedene Konten' = 'Distinct accounts'
    'Verschlüsselt' = 'Encrypted'
    'Version' = 'Version'
    'Version C/B' = 'Version C/B'
    'Verweigert für' = 'Denied for'
    'Verwendung' = 'Use'
    'Virenschutz' = 'Antivirus'
    'Von' = 'From'
    'WMI-Filter' = 'WMI filter'
    'Weg' = 'Path'
    'Weiterleitungen' = 'Forwarding'
    'Wert' = 'Value'
    'Wirksame Richtlinie' = 'Effective policy'
    'Wirksames Recht' = 'Effective permission'
    'Wirkung' = 'Effect'
    'Zeit' = 'Time'
    'Zeitpunkt' = 'Time'
    'Zeitraum' = 'Period'
    'Zertifikat' = 'Certificate'
    'Zone' = 'Zone'
    'Zonen' = 'Zones'
    'Zonenübertragung' = 'Zone transfer'
    'Zugriff für große Gruppen' = 'Access for large groups'
    'Zugriffsbasierte Aufzählung' = 'Access-based enumeration'
    'Zuletzt' = 'Last'
    'Zustand' = 'State'
    'davon Größe (GB)' = 'Of which size (GB)'
    'davon aktiv' = 'Of which active'
    'davon inaktiv' = 'Of which inactive'
    'davon unbekannte Konten' = 'Of which unknown accounts'
    'ist Mitglied von' = 'Member of'
    'Älter als 5 Jahre' = 'Older than 5 years'
    'Älter als 5 Jahren' = 'Older than 5 years'
    'Ältester Eintrag' = 'Oldest entry'
    'Über' = 'Via'
    'Überwachungsprotokoll' = 'Audit log'
}
# Werte, die der Code intern auf Deutsch vergleicht (z. B. Art = 'Benutzer'), erscheinen im englischen Bericht übersetzt (nur Anzeige und CSV)
$WerteEn = New-Object System.Collections.Specialized.OrderedDictionary([System.StringComparer]::Ordinal)   # Groß-/Kleinschreibung zählt (Verwaist / verwaist)
$WerteEn.Add('Alle erweiterten Rechte (u. a. Kennwort zurücksetzen)', 'All extended rights (including password reset)')
$WerteEn.Add('Kein Serverbetriebssystem (Support nicht bewertet)', 'No server operating system (support not assessed)')
$WerteEn.Add('Rechte je Ordner (nur dort, wo sie gesetzt werden)', 'Permissions per folder (only where set)')
$WerteEn.Add('Mehrfach freigegebene Ordner (nur einmal gezählt)', 'Folders shared multiple times (counted only once)')
$WerteEn.Add('Übersprungene Verknüpfungen (Junctions/Symlinks)', 'Skipped links (Junctions/Symlinks)')
$WerteEn.Add('Der Ordner der Freigabe existiert nicht mehr.', 'The share folder no longer exists.')
$WerteEn.Add('DHCP-Überwachungsprotokoll ist ausgeschaltet', 'DHCP audit log is disabled')
$WerteEn.Add('Vollzugriff und "Senden als" im Überblick', 'Full access and "Send as" overview')
$WerteEn.Add('Ordner der Freigabe existiert nicht mehr', 'Share folder no longer exists')
$WerteEn.Add('Enthält außer Sicherheitseinstellungen:', 'Contains except security settings:')
$WerteEn.Add('Nicht geprüft - Postfach nicht gefunden', 'Not checked - mailbox not found')
$WerteEn.Add('WMI-Filter nicht lesbar - nicht geprüft', 'WMI filter not readable - not checked')
$WerteEn.Add('Abfrage fehlgeschlagen - nicht geprüft', 'Query failed - not checked')
$WerteEn.Add('Nicht im Active Directory gespeichert', 'Not stored in Active Directory')
$WerteEn.Add('keine im Active Directory autorisiert', 'None authorized in Active Directory')
$WerteEn.Add('Zugriff auf den Ordner nicht möglich', 'Access to folder not possible')
$WerteEn.Add('Zertifikate (öffentliche Schlüssel)', 'Certificates (public keys)')
$WerteEn.Add('Domänencontroller der Organisation', 'Organization domain controllers')
$WerteEn.Add('Remotedesktop erlaubt Verbindungen', 'Remote Desktop allows connections')
$WerteEn.Add('autorisierte DHCP-Server ermitteln', 'Determine authorized DHCP servers')
$WerteEn.Add('Windows Server 2008 R2 und älter', 'Windows Server 2008 R2 and older')
$WerteEn.Add('Konto deaktiviert oder gesperrt', 'Account disabled or locked')
$WerteEn.Add('Vollzugriff (Zusatz: ObjectType', 'Full access (additional: ObjectType')
$WerteEn.Add('Kennwort oder Konto abgelaufen', 'Password or account expired')
$WerteEn.Add('Lokale Administratorengruppen', 'Local administrator groups')
$WerteEn.Add('Berechtigungen/Besitz ändern', 'Change permissions/ownership')
$WerteEn.Add('Freigabepfade nicht sichtbar', 'Share paths not visible')
$WerteEn.Add('Aktive privilegierte Konten', 'Active privileged accounts')
$WerteEn.Add('Gruppe hat keine Mitglieder', 'Group has no members')
$WerteEn.Add('Keine Verbindung zum Server', 'No connection to server')
$WerteEn.Add('unerlaubte Zeichen im Namen', 'Illegal characters in name')
$WerteEn.Add('Geprüfte Domänencontroller', 'Checked domain controllers')
$WerteEn.Add('Authentifizierte Benutzer', 'Authenticated Users')
$WerteEn.Add('Schwellen dieses Berichts', 'Thresholds of this report')
$WerteEn.Add('Benutzer (inetOrgPerson)', 'User (inetOrgPerson)')
$WerteEn.Add('Datenbestand je Freigabe', 'Data inventory per share')
$WerteEn.Add('Kontosperrungsrichtlinie', 'Account lockout policy')
$WerteEn.Add('Sicherheitseinstellungen', 'Security settings')
$WerteEn.Add('Alle erweiterten Rechte', 'All extended rights')
$WerteEn.Add('Hinweise zur Auswertung', 'Notes for evaluation')
$WerteEn.Add('Nicht abgefragte Server', 'Servers not queried')
$WerteEn.Add('Unterordner und Dateien', 'Subfolders and files')
$WerteEn.Add('Bestätigtes Schreiben:', 'Validated write:')
$WerteEn.Add('Installierte Programme', 'Installed programs')
$WerteEn.Add('Konto für DNS-Einträge', 'Account for DNS records')
$WerteEn.Add('Kontosperrungsschwelle', 'Account lockout threshold')
$WerteEn.Add('Nicht lesbare Bereiche', 'Unreadable areas')
$WerteEn.Add('Berechtigungen ändern', 'Change permissions')
$WerteEn.Add('Hinweise zum Auslesen', 'Notes for reading')
$WerteEn.Add('Windows 8.1 und älter', 'Windows 8.1 and older')
$WerteEn.Add('Überwachungsprotokoll', 'Audit log')
$WerteEn.Add('Nicht berücksichtigt', 'Not considered')
$WerteEn.Add('Anonymous-Anmeldung', 'Anonymous logon')
$WerteEn.Add('Infrastrukturmaster', 'Infrastructure master')
$WerteEn.Add('Secure Boot ist aus', 'Secure Boot is off')
$WerteEn.Add('Domänennamenmaster', 'Domain naming master')
$WerteEn.Add('Kennwortrichtlinie', 'Password policy')
$WerteEn.Add('Komponentenwartung', 'Component servicing')
$WerteEn.Add('alle Eigenschaften', 'all properties')
$WerteEn.Add('Besitz übernehmen', 'Take ownership')
$WerteEn.Add('Domänencontroller', 'Domain Controllers')
$WerteEn.Add('Falsches Kennwort', 'Incorrect password')
$WerteEn.Add('Konten-Operatoren', 'Account Operators')
$WerteEn.Add('Nur dieser Ordner', 'Only this folder')
$WerteEn.Add('Domänen-Benutzer', 'Domain Users')
$WerteEn.Add('Format unbekannt', 'Format unknown')
$WerteEn.Add('Sonstige Dateien', 'Other files')
$WerteEn.Add('Ältester Eintrag', 'Oldest entry')
$WerteEn.Add('Domänencomputer', 'Domain Computers')
$WerteEn.Add('Konto unbekannt', 'Account unknown')
$WerteEn.Add('Standard (alle)', 'Default (all)')
$WerteEn.Add('Computerkonten', 'Computer accounts')
$WerteEn.Add('Konto gesperrt', 'Account locked')
$WerteEn.Add('Nicht erreicht', 'Not reached')
$WerteEn.Add('DHCP-Bereiche', 'DHCP scopes')
$WerteEn.Add('Dieser Ordner', 'This folder')
$WerteEn.Add('Gruppe/System', 'Group/system')
$WerteEn.Add('Remotedesktop', 'Remote Desktop')
$WerteEn.Add('Arbeitsplatz', 'Workstation')
$WerteEn.Add('Schemamaster', 'Schema master')
$WerteEn.Add('DHCP-Server', 'DHCP server')
$WerteEn.Add('Datenbanken', 'Databases')
$WerteEn.Add('Sonderkonto', 'Special account')
$WerteEn.Add('Unterordner', 'Subfolders')
$WerteEn.Add('Vollzugriff', 'Full control')
$WerteEn.Add('Zugriff von', 'Access from')
$WerteEn.Add('DNS-Server', 'DNS server')
$WerteEn.Add('Inaktiv ab', 'Inactive from')
$WerteEn.Add('Schreiben:', 'Write:')
$WerteEn.Add('Verweigern', 'Deny')
$WerteEn.Add('DNS-Zonen', 'DNS zones')
$WerteEn.Add('Schreiben', 'Write')
$WerteEn.Add('Sonstiges', 'Other')
$WerteEn.Add('Standorte', 'Sites')
$WerteEn.Add('Unbekannt', 'Unknown')
$WerteEn.Add('unbekannt', 'unknown')
$WerteEn.Add('(Wurzel)', '(root)')
$WerteEn.Add('Anlegen:', 'Create:')
$WerteEn.Add('Benutzer', 'User')
$WerteEn.Add('Erlauben', 'Allow')
$WerteEn.Add('Netzwerk', 'Network')
$WerteEn.Add('Verwaist', 'Orphaned')
$WerteEn.Add('verwaist', 'orphaned')
$WerteEn.Add('Dateien', 'Files')
$WerteEn.Add('Gruppen', 'Groups')
$WerteEn.Add('Anonym', 'Anonymous')
$WerteEn.Add('Domäne', 'Domain')
$WerteEn.Add('Gruppe', 'Group')
$WerteEn.Add('Umfang', 'Scope')
$WerteEn.Add('direkt', 'direct')
$WerteEn.Add('Ändern', 'Change')
$WerteEn.Add('Lesen', 'Read')
$WerteEn.Add('über', 'via')
$WerteMuster = @()
foreach ($k in ($WerteEn.Keys | Sort-Object Length -Descending)) { $WerteMuster += , @([regex]::new('(?<![\p{L}\d])' + [regex]::Escape($k) + '(?![\p{L}\d])'), $WerteEn[$k]) }
function WertAnzeige($v) {
    if (-not $script:En -or $v -isnot [string] -or -not $v) { return $v }
    if ($v.Contains('\')) { return $v }   # Kontonamen (DOMÄNE\Name) und Pfade bleiben, wie sie sind
    foreach ($m in $WerteMuster) { $v = $m[0].Replace($v, $m[1]) }
    $v
}
function SpaltenAnzeige($liste) {
    if (-not $script:En -or -not @($liste).Count) { return $liste }
    foreach ($o in @($liste)) {
        $n = [ordered]@{}
        foreach ($p in $o.PSObject.Properties) { $k = $p.Name; if ($SpaltenEn.ContainsKey($k)) { $k = $SpaltenEn[$k] }; $i = 2; $k0 = $k; while ($n.Contains($k)) { $k = "$k0 $i"; $i++ }; $n[$k] = WertAnzeige $p.Value }
        [pscustomobject]$n
    }
}
# Sortierschluessel fuer Nummern wie K2, K10: erst Buchstaben, dann die Zahl als Zahl
function IdZahl($id) { $m = [regex]::Match([string]$id, '\d+'); if ($m.Success) { [int]$m.Value } else { 0 } }
function IdText($id) { [regex]::Replace([string]$id, '\d+', '') }

$Befunde = New-Object System.Collections.Generic.List[object]

# Platzhalterzeilen ("nicht geprüft") erkennt das Skript nur an den vollständigen, festen Texten, die es selbst erzeugt - keine Teilstring-Suche
# in Konto-, Objekt- oder Beschreibungstexten, die ein Angreifer beeinflussen könnte
# Deutsch und Englisch: Skripte erzeugen den Text je nach -Language, erkannt werden beide Fassungen
$PlatzhalterTexte = @('Abfrage fehlgeschlagen - nicht geprüft', 'Berechtigungen nicht lesbar - nicht geprüft', 'SYSVOL nicht lesbar - nicht geprüft', 'WMI-Filter nicht lesbar - nicht geprüft', 'Kein Domänencontroller ermittelbar - SYSVOL nicht geprüft', 'Inhalt des Ordners nicht lesbar - Richtlinie nicht geprüft',
    'Query failed - not checked', 'Permissions not readable - not checked', 'SYSVOL not readable - not checked', 'WMI filter not readable - not checked', 'No domain controller found - SYSVOL not checked', 'Folder content not readable - policy not checked')
function IstPlatzhalter($o) {
    # Ein Skript kann $script:PlatzhalterSpalten setzen: dann zählen nur diese Spalten (Namensspalten sind von Benutzern beeinflussbar)
    foreach ($pr in $o.PSObject.Properties) {
        if ($script:PlatzhalterSpalten -and ($pr.Name -notin $script:PlatzhalterSpalten)) { continue }
        if (($pr.Value -is [string]) -and ($PlatzhalterTexte -ccontains $pr.Value)) { return $true }
    }
    ("$($o.Feststellung)" -clike 'Nicht geprüft - *') -or ("$($o.Feststellung)" -clike 'Not checked - *')
}

# Besteht ein Prüfpunkt nur aus Platzhalterzeilen (Abfrage fehlgeschlagen), ist das kein Mangel dieses Prüfpunkts: als Info führen statt
# mit seiner Stufe, damit es in den Kacheln nicht als hoher Befund zählt. Nur für Prüfpunkte, deren Ergebnis von einer Abfrage abhängt.
function InfoWennNurNichtGeprueft([string]$IdMuster) {
    foreach ($x in ($Befunde | Where-Object { $_.Id -match $IdMuster -and $_.Anzahl -gt 0 -and $_.Stufe -ne 'Info' })) {
        if (-not @($x.Objekte | Where-Object { -not (IstPlatzhalter $_) }).Count) { $x.Stufe = 'Info' }
    }
}

# Prüfpunkte, die für einzelne Server nicht ausgewertet werden konnten (Abfrage fehlgeschlagen), dürfen nicht als "ohne Befund" erscheinen:
# je Server eine Zeile "nicht geprüft" anhängen; hatte der Prüfpunkt sonst keinen Treffer, steht er als Info statt als OK.
function NichtGeprueft([string]$IdMuster, [string[]]$Namen, [string]$Verweis, [string]$Spalte = 'Server', [string]$Grund = $(T 'Abfrage nicht möglich' 'query not possible')) {
    foreach ($x in ($Befunde | Where-Object { $_.Id -match $IdMuster })) {
        $war = $x.Anzahl
        $zeilen = @($Namen | ForEach-Object { [pscustomobject][ordered]@{ $Spalte = $_; Feststellung = (T "Nicht geprüft - $Grund (siehe $Verweis)" "Not checked - $Grund (see $Verweis)"); Details = '' } })
        $x.Objekte = @($x.Objekte) + $zeilen
        $x.Anzahl = $x.Objekte.Count
        if (-not $war) { $x.Stufe = 'Info' }
    }
}
function Befund {
    param([string]$Id, [string]$Bereich, [string]$Titel, [ValidateSet('Hoch', 'Mittel', 'Niedrig', 'Info')][string]$Stufe,
        $Objekte, [string]$Erklaerung, [string]$Empfehlung)
    $liste = @($Objekte | Where-Object { $null -ne $_ })
    $status = $Stufe
    if ($liste.Count -eq 0) { $status = 'OK' }
    $Befunde.Add([pscustomobject]@{
            Id = $Id; Bereich = $Bereich; Titel = $Titel; Stufe = $status; Anzahl = $liste.Count
            Objekte = $liste; Erklaerung = $Erklaerung; Empfehlung = $Empfehlung
        })
}

# Schreibt den HTML-Bericht (und auf Wunsch CSV-Dateien). $Abschnitte: Überschrift -> Objekte oder fertiges HTML,
# erscheinen vor den Befunden. $Anhang: Überschrift -> Text, erscheint am Ende unverändert.
function BerichtSchreiben {
    param([string]$Name, [string]$Datei, [string]$Bezug, [string]$BezugName = $(T 'Domäne' 'Domain'), [System.Collections.IDictionary]$Umgebung,
        [System.Collections.IDictionary]$Abschnitte = @{}, [System.Collections.IDictionary]$Anhang = @{}, [string]$Fuss)
    Write-Host (T '  Bericht schreiben ...' '  Writing report ...')
    $jetzt = Get-Date
    $ziel = $script:Ausgabe
    if (-not $ziel) { $ziel = Join-Path (Get-Location).Path ("{0}_{1}_{2:yyyy-MM-dd_HHmm}.html" -f $Datei, ($Bezug -replace '[^\w.\-]+', '-'), $jetzt) }
    $ziel = [System.IO.Path]::GetFullPath($ziel)
    $ordner = Split-Path $ziel -Parent
    if (-not (Test-Path $ordner)) { New-Item -ItemType Directory -Path $ordner -Force | Out-Null }

    $stufeText = @{ Hoch = (T 'Hoch' 'High'); Mittel = (T 'Mittel' 'Medium'); Niedrig = (T 'Niedrig' 'Low'); Info = 'Info'; OK = 'OK' }
    $rang = @{ Hoch = 0; Mittel = 1; Niedrig = 2; Info = 3; OK = 4 }
    $zahl = @{}
    foreach ($s in 'Hoch', 'Mittel', 'Niedrig', 'Info', 'OK') { $zahl[$s] = @($Befunde | Where-Object Stufe -eq $s).Count }
    $css = @'
body{font-family:Segoe UI,Arial,sans-serif;color:#2b2f36;margin:0;background:#f4f6fa}
.seite{max-width:1100px;margin:0 auto;padding:28px 24px 60px}
h1{color:#1f5fae;margin:0 0 4px;font-size:28px} h2{color:#1f5fae;font-size:19px;margin:34px 0 10px;border-bottom:2px solid #d0d5dd;padding-bottom:6px}
.meta{color:#6b7280;font-size:13px;margin-bottom:22px}
.kacheln{display:flex;gap:14px;flex-wrap:wrap;margin:18px 0}
.kachel{flex:1;min-width:150px;background:#fff;border-radius:10px;padding:14px 18px;border-left:8px solid #999;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.kachel b{display:block;font-size:30px} .kachel span{font-size:13px;color:#6b7280}
table{border-collapse:collapse;width:100%;background:#fff;font-size:13px;margin:8px 0 4px}
th{background:#1f5fae;color:#fff;text-align:left;padding:7px 9px;font-weight:600} td{padding:6px 9px;border-bottom:1px solid #e5e7eb;vertical-align:top;overflow-wrap:break-word}
tr:nth-child(even) td{background:#f9fafb}
.marke{display:inline-block;min-width:58px;text-align:center;padding:2px 9px;border-radius:11px;color:#fff;font-size:12px;font-weight:600}
.m-Hoch{background:#d92d20}.m-Mittel{background:#f79009}.m-Niedrig{background:#ca8a04}.m-OK{background:#12b76a}.m-Info{background:#1f5fae}
.befund{background:#fff;border-radius:10px;padding:16px 20px;margin:14px 0;border-left:8px solid #999;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.befund h3{margin:0 0 6px;font-size:16px} .befund p{margin:5px 0;font-size:14px;line-height:1.45} .befund .emp{color:#14532d}
.Hoch{border-color:#d92d20}.Mittel{border-color:#f79009}.Niedrig{border-color:#eab308}.OK{border-color:#12b76a}.Info{border-color:#1f5fae}
.klein{color:#6b7280;font-size:12px} a{color:#1f5fae;text-decoration:none}
pre{background:#fff;border:1px solid #d0d5dd;border-radius:8px;padding:12px 14px;font-size:12px;white-space:pre-wrap;word-break:break-word}
.fuss{margin-top:40px;color:#6b7280;font-size:12px;border-top:1px solid #d0d5dd;padding-top:12px}
@media print{body{background:#fff}.befund,.kachel{box-shadow:none;border:1px solid #d0d5dd;border-left-width:8px;break-inside:avoid}}
'@
    $h = New-Object System.Text.StringBuilder
    [void]$h.Append("<!DOCTYPE html><html lang='$Language'><head><meta charset='utf-8'><title>$(Html $Name) $(Html $Bezug)</title><style>$css</style></head><body><div class='seite'>")
    [void]$h.Append("<h1>$(Html $Name)</h1><div class='meta'>$(Html $BezugName) <b>$(Html $Bezug)</b> · $(T 'erstellt am' 'created') $($jetzt.ToString((T 'dd.MM.yyyy HH:mm' 'yyyy-MM-dd HH:mm'))) $(T 'von' 'by') $(Html "$env:USERDOMAIN\$env:USERNAME") $(T 'auf' 'on') $(Html $env:COMPUTERNAME)</div>")
    [void]$h.Append("<div class='kacheln'>")
    foreach ($s in 'Hoch', 'Mittel', 'Niedrig', 'Info', 'OK') {
        $text = @{ Hoch = (T 'Befunde mit hoher Priorität' 'findings, high priority'); Mittel = (T 'mit mittlerer Priorität' 'medium priority'); Niedrig = (T 'mit niedriger Priorität' 'low priority'); Info = (T 'Hinweise und nicht geprüft' 'notes and not checked'); OK = (T 'Prüfpunkte ohne Befund' 'checks without findings') }[$s]
        [void]$h.Append("<div class='kachel $s'><b>$($zahl[$s])</b><span>$text</span></div>")
    }
    [void]$h.Append("</div><h2>$(T 'Umgebung' 'Environment')</h2><table><tr><th>$(T 'Eigenschaft' 'Property')</th><th>$(T 'Wert' 'Value')</th></tr>")
    foreach ($k in $Umgebung.Keys) { [void]$h.Append("<tr><td>$(Html (WertAnzeige $k))</td><td>$(Html (WertAnzeige $Umgebung[$k]))</td></tr>") }
    [void]$h.Append("</table>")

    [void]$h.Append("<h2>$(T 'Übersicht der Prüfpunkte' 'Overview of checks')</h2><table><tr><th>$(T 'Nr.' 'No.')</th><th>$(T 'Bereich' 'Area')</th><th>$(T 'Prüfpunkt' 'Check')</th><th>$(T 'Treffer' 'Hits')</th><th>$(T 'Bewertung' 'Rating')</th></tr>")
    foreach ($x in $Befunde) {
        [void]$h.Append("<tr><td>$($x.Id)</td><td>$(Html $x.Bereich)</td><td><a href='#$($x.Id)'>$(Html $x.Titel)</a></td><td>$($x.Anzahl)</td><td><span class='marke m-$($x.Stufe)'>$($stufeText[$x.Stufe])</span></td></tr>")
    }
    [void]$h.Append("</table>")

    foreach ($k in $Abschnitte.Keys) {
        [void]$h.Append("<h2>$(Html (WertAnzeige $k))</h2>")
        $inhalt = $Abschnitte[$k]
        if ($inhalt -is [string]) { [void]$h.Append($inhalt) }
        elseif (@($inhalt).Count) { [void]$h.Append(((Einheitlich $inhalt | ConvertTo-Html -Fragment) -join "`n")) }
        else { [void]$h.Append("<p class='klein'>$(T 'Keine Einträge.' 'No entries.')</p>") }
    }

    [void]$h.Append("<h2>$(T 'Befunde' 'Findings')</h2>")
    $offen = @($Befunde | Where-Object { $_.Stufe -ne 'OK' } | Sort-Object { $rang[$_.Stufe] }, { IdText $_.Id }, { IdZahl $_.Id })
    foreach ($x in $offen) {
        [void]$h.Append("<div class='befund $($x.Stufe)' id='$($x.Id)'><h3><span class='marke m-$($x.Stufe)'>$($stufeText[$x.Stufe])</span> $($x.Id) · $(Html $x.Titel) <span class='klein'>($($x.Anzahl) $(T 'Treffer' 'hits'))</span></h3>")
        [void]$h.Append("<p>$(Html $x.Erklaerung)</p><p class='emp'><b>$(T 'Empfehlung' 'Recommendation'):</b> $(Html $x.Empfehlung)</p>")
        [void]$h.Append(((Einheitlich $x.Objekte | ConvertTo-Html -Fragment) -join "`n"))
        [void]$h.Append("</div>")
    }
    if (-not $offen.Count) { [void]$h.Append("<p>$(T 'Keine Auffälligkeiten gefunden.' 'No anomalies found.')</p>") }
    $ok = @($Befunde | Where-Object Stufe -eq 'OK')
    if ($ok.Count) {
        [void]$h.Append("<h2>$(T 'Ohne Befund' 'No findings')</h2><table><tr><th>$(T 'Nr.' 'No.')</th><th>$(T 'Prüfpunkt' 'Check')</th></tr>")
        foreach ($x in $ok) { [void]$h.Append("<tr id='$($x.Id)'><td>$($x.Id)</td><td>$(Html $x.Titel)</td></tr>") }
        [void]$h.Append("</table>")
    }
    foreach ($k in $Anhang.Keys) { [void]$h.Append("<h2>$(Html $k)</h2><pre>$(Html $Anhang[$k])</pre>") }
    [void]$h.Append("<div class='fuss'>$(Html $Name) Version $script:Version · IT Admin Tools by MM · $(T 'Das Skript hat ausschließlich gelesen und nichts verändert.' 'The script only read and changed nothing.') $(Html $Fuss)</div>")
    [void]$h.Append("</div></body></html>")
    [System.IO.File]::WriteAllText($ziel, $h.ToString(), (New-Object System.Text.UTF8Encoding($true)))

    $basis = [System.IO.Path]::ChangeExtension($ziel, $null).TrimEnd('.')
    if ($script:Csv) {
        foreach ($x in ($Befunde | Where-Object { $_.Anzahl -gt 0 })) {
            Einheitlich $x.Objekte | Export-Csv -Path "${basis}_$($x.Id).csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8
        }
    }
    Write-Host ''
    Write-Host ((T "Fertig: {0} hoch, {1} mittel, {2} niedrig, {3} ohne Befund." "Done: {0} high, {1} medium, {2} low, {3} without findings.") -f $zahl.Hoch, $zahl.Mittel, $zahl.Niedrig, $zahl.OK) -ForegroundColor Green
    Write-Host ((T 'Bericht' 'Report') + ": $ziel")
    if (-not $script:NichtOeffnen -and [Environment]::UserInteractive) { try { Start-Process $ziel } catch { } }
    $basis
}

# ------------------------------------------------------------------ Auslesen auf einem Computer
# Dieser Block läuft lokal oder per Invoke-Command auf dem Zielcomputer. Er liest nur.
$lese = {
    $ErrorActionPreference = 'Stop'   # auch in der Fernsitzung: Lesefehler werden gemeldet, nicht verschluckt
    $ergebnis = [ordered]@{ Zertifikate = @(); Fehler = @(); Ccs = $false }
    # Zertifikate, die in Gebrauch sind: HTTPS-Bindungen von HTTP.sys und Remotedesktop
    $gebrauch = @{}
    try {
        # HTTP.sys: Bindungen an IP:Port (SslBindingInfo) und an Hostnamen mit SNI (SslSniBindingInfo)
        foreach ($art in @(@('SslBindingInfo', 'HTTPS'), @('SslSniBindingInfo', 'HTTPS (SNI)'))) {
            $wurzel = "HKLM:\SYSTEM\CurrentControlSet\Services\HTTP\Parameters\$($art[0])"
            if (Test-Path $wurzel) {
                foreach ($k in Get-ChildItem $wurzel) {
                    $h = (Get-ItemProperty -Path $k.PSPath -Name SslCertHash -ErrorAction SilentlyContinue).SslCertHash
                    $hp = (Get-ItemProperty -Path $k.PSPath -Name HostnamePort -ErrorAction SilentlyContinue).HostnamePort   # nur bei SNI vorhanden, dort ist der Schlüsselname eine Kennung
                    $ziel = if ($hp) { [string]$hp } else { $k.PSChildName }
                    if ($h) { $hex = (($h | ForEach-Object { $_.ToString('X2') }) -join ''); $gebrauch[$hex] = (@($gebrauch[$hex]) + "$($art[1]) $ziel") | Where-Object { $_ } }
                }
            }
        }
        # Zentraler Zertifikatspeicher (CCS): die Zuordnung liegt im Dateisystem, nicht in der Registrierung - nur melden, dass er benutzt wird
        $ccs = 'HKLM:\SYSTEM\CurrentControlSet\Services\HTTP\Parameters\SslCcsBindingInfo'
        if ((Test-Path $ccs) -and @(Get-ChildItem $ccs).Count) { $ergebnis.Ccs = $true }
    }
    catch { $ergebnis.Fehler += (T "HTTPS-Bindungen nicht lesbar: $($_.Exception.Message)" "HTTPS bindings not readable: $($_.Exception.Message)") }
    try {
        $rdp = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name SSLCertificateSHA1Hash -ErrorAction SilentlyContinue).SSLCertificateSHA1Hash
        if ($rdp) { $hex = (($rdp | ForEach-Object { $_.ToString('X2') }) -join ''); $gebrauch[$hex] = (@($gebrauch[$hex]) + 'Remotedesktop') | Where-Object { $_ } }
    }
    catch { $ergebnis.Fehler += (T "Remotedesktop-Zertifikat nicht lesbar: $($_.Exception.Message)" "Remote Desktop certificate not readable: $($_.Exception.Message)") }

    foreach ($speicher in 'My', 'Remote Desktop', 'WebHosting') {
        if (-not (Test-Path "Cert:\LocalMachine\$speicher")) { continue }
        try {
            foreach ($c in (Get-ChildItem "Cert:\LocalMachine\$speicher")) {
                $bits = $null; $alg = ''
                try {
                    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($c)
                    if ($rsa) { $bits = $rsa.KeySize; $alg = 'RSA' }
                    else { $alg = [string]$c.PublicKey.Oid.FriendlyName }
                }
                catch { $alg = [string]$c.PublicKey.Oid.FriendlyName }
                $san = ''
                $e = $c.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.17' } | Select-Object -First 1
                if ($e) { $san = ($e.Format($false) -replace '\s+', ' ') }
                $ergebnis.Zertifikate += [pscustomobject]@{
                    Speicher = $speicher; Subject = [string]$c.Subject; Issuer = [string]$c.Issuer; Thumbprint = [string]$c.Thumbprint
                    NotBefore = $c.NotBefore; NotAfter = $c.NotAfter; PrivaterSchluessel = [bool]$c.HasPrivateKey
                    Signatur = [string]$c.SignatureAlgorithm.FriendlyName; Schluessel = $alg; Bits = $bits; San = $san
                    Verwendung = (@($gebrauch[[string]$c.Thumbprint]) -join ', ')
                }
            }
        }
        catch { $ergebnis.Fehler += (T "Speicher $speicher nicht lesbar: $($_.Exception.Message)" "Store $speicher not readable: $($_.Exception.Message)") }
    }
    [pscustomobject]$ergebnis
}

# ------------------------------------------------------------------ Ziele bestimmen
$jetzt = Get-Date
$ziele = New-Object System.Collections.Generic.List[string]
$lokal = $env:COMPUTERNAME
if ($Computer) { foreach ($c in $Computer) { if ($c) { $ziele.Add($c) } } }
if ($AusAD) {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Host (T 'Das PowerShell-Modul "ActiveDirectory" fehlt (für -AusAD nötig).' 'The PowerShell module "ActiveDirectory" is missing (required for -AusAD).') -ForegroundColor Red
        Write-Host 'Windows Server:  Install-WindowsFeature RSAT-AD-PowerShell'
        exit 2
    }
    Import-Module ActiveDirectory
    $ad = @{}
    if ($Server) { $ad.Server = $Server }
    $grenze = $jetzt.AddDays(-60)
    Get-ADComputer -Filter { Enabled -eq $true -and OperatingSystem -like '*Server*' } -Properties LastLogonDate @ad |
        Where-Object { $_.LastLogonDate -and $_.LastLogonDate -gt $grenze } | ForEach-Object { $ziele.Add([string]$_.DNSHostName) }
}
$fern = @($ziele | Where-Object { $_ -and $_.Split('.')[0] -ne $lokal } | Sort-Object -Unique)
$pruefLokal = -not $OhneLokalen

# ------------------------------------------------------------------ Einlesen
$alle = New-Object System.Collections.Generic.List[object]
$nichtGeprueft = New-Object System.Collections.Generic.List[object]
$geprueft = New-Object System.Collections.Generic.List[string]
$ccsRechner = New-Object System.Collections.Generic.List[string]   # Computer mit zentralem Zertifikatspeicher: nur die Zuordnung (Z1, Z3, Z5) ist unvollständig
$nichtErreicht = 0

function Aufnehmen([string]$Name, $Ergebnis) {
    $geprueft.Add($Name)
    if ($Ergebnis.Ccs) { $ccsRechner.Add($Name) }
    foreach ($f in @($Ergebnis.Fehler)) { if ($f) { $nichtGeprueft.Add([pscustomobject][ordered]@{ Computer = $Name; Grund = $f }) } }
    foreach ($z in @($Ergebnis.Zertifikate)) {
        if (-not $z) { continue }
        $alle.Add([pscustomobject]@{
                Computer = $Name; Speicher = $z.Speicher; Subject = $z.Subject; Issuer = $z.Issuer; Thumbprint = $z.Thumbprint
                Von = [datetime]$z.NotBefore; Bis = [datetime]$z.NotAfter; Schluessel = [bool]$z.PrivaterSchluessel
                Signatur = $z.Signatur; Algorithmus = $z.Schluessel; Bits = $z.Bits; San = $z.San; Verwendung = $z.Verwendung
            })
    }
}

if ($pruefLokal) {
    Write-Host (T "Lese Zertifikate von $lokal ..." "Reading certificates from $lokal ...") -ForegroundColor Cyan
    try { Aufnehmen $lokal (& $lese) }
    catch { $nichtGeprueft.Add([pscustomobject][ordered]@{ Computer = $lokal; Grund = $_.Exception.Message }) }
}
if ($fern.Count) {
    Write-Host (T "Frage $($fern.Count) Computer über WinRM ab ..." "Querying $($fern.Count) computers over WinRM ...") -ForegroundColor Cyan
    $ic = @{ ComputerName = $fern; ScriptBlock = $lese; ErrorAction = 'SilentlyContinue'; ErrorVariable = 'icFehler' }
    if ($Anmeldung) { $ic.Credential = $Anmeldung }
    $icFehler = @()
    $antworten = @(Invoke-Command @ic)
    $erhalten = @{}
    foreach ($a in $antworten) { $n = [string]$a.PSComputerName; $erhalten[$n.Split('.')[0].ToLowerInvariant()] = $true; Aufnehmen $n $a }
    foreach ($f in $fern) {
        if (-not $erhalten.ContainsKey($f.Split('.')[0].ToLowerInvariant())) {
            $grund = @($icFehler | Where-Object { ([string]$_.TargetObject) -like "$($f.Split('.')[0])*" -or ([string]$_.OriginInfo.PSComputerName) -like "$($f.Split('.')[0])*" } | ForEach-Object { ($_.Exception.Message -replace '\s+', ' ') } | Select-Object -First 1)
            if (-not $grund) { $grund = (T 'keine Antwort' 'no response') }
            $nichtErreicht++
            $nichtGeprueft.Add([pscustomobject][ordered]@{ Computer = $f; Grund = Kurz $grund 200 })
        }
    }
}

# ------------------------------------------------------------------ Prüfungen
function Zeile($z, $zusatz = @{}) {
    $r = [ordered]@{ Computer = $z.Computer; Zertifikat = Kurz $z.Subject 90; Speicher = $z.Speicher; 'Gültig bis' = Datum $z.Bis; Verwendung = $(if ($z.Verwendung) { $z.Verwendung } else { '-' }) }
    foreach ($k in $zusatz.Keys) { $r[$k] = $zusatz[$k] }
    [pscustomobject]$r
}
$gueltig = @($alle | Where-Object { $_.Bis -ge $jetzt })
$abgelaufen = @($alle | Where-Object { $_.Bis -lt $jetzt })

Befund 'Z1' (T 'Ablauf' 'Expiration') (T 'Abgelaufenes Zertifikat wird verwendet (HTTPS oder Remotedesktop)' 'Expired certificate is in use (HTTPS or Remote Desktop)') 'Hoch' `
    @($abgelaufen | Where-Object Verwendung | ForEach-Object { Zeile $_ @{ 'Tage seit Ablauf' = [int]($jetzt - $_.Bis).TotalDays } }) `
    (T 'Ein Dienst liefert weiter ein abgelaufenes Zertifikat aus. Browser und Programme zeigen Warnungen oder verweigern die Verbindung, Anmeldungen über Remotedesktop zeigen eine Sicherheitswarnung.' 'A service continues to serve an expired certificate. Browsers and applications show warnings or reject the connection, Remote Desktop logons show a security warning.') `
    (T 'Zertifikat erneuern und in der Bindung des Dienstes austauschen. Für die Zukunft eine Erinnerung oder automatische Erneuerung einrichten.' 'Renew the certificate and replace it in the service binding. Set up a reminder or automatic renewal for the future.')

Befund 'Z2' (T 'Ablauf' 'Expiration') (T "Zertifikate laufen in den nächsten $WarnTage Tagen ab" "Certificates expire in the next $WarnTage days") 'Mittel' `
    @($gueltig | Where-Object { $_.Bis -lt $jetzt.AddDays($WarnTage) -and $_.Schluessel } | Sort-Object Bis | ForEach-Object { Zeile $_ @{ 'Tage bis Ablauf' = [int]($_.Bis - $jetzt).TotalDays } }) `
    (T 'Läuft ein Zertifikat ab, fallen die damit gesicherten Dienste aus, oft am Wochenende und ohne Vorwarnung. Betroffen sind nur Zertifikate mit privatem Schlüssel, also solche, die ein Server selbst verwendet.' 'If a certificate expires, the services secured with it fail, often on weekends and without warning. Only certificates with private keys are affected, i.e., those used by a server itself.') `
    (T 'Erneuerung rechtzeitig einplanen. Bei öffentlichen Zertifikaten die Laufzeit der Bestellung beachten.' 'Plan renewal in time. For public certificates, note the validity period of the order.')

Befund 'Z3' (T 'Aufräumen' 'Cleanup') (T 'Abgelaufene Zertifikate mit privatem Schlüssel, die nicht verwendet werden' 'Expired certificates with private keys that are not in use') 'Niedrig' `
    @($abgelaufen | Where-Object { $_.Schluessel -and -not $_.Verwendung } | Sort-Object Bis | ForEach-Object { Zeile $_ @{ 'Tage seit Ablauf' = [int]($jetzt - $_.Bis).TotalDays } }) `
    (T 'Der Server speichert alte Zertifikate samt Schlüssel weiter. Sie sind nutzlos, belegen aber Platz im Speicher und verwirren bei der Fehlersuche. Der private Schlüssel bleibt ein Risiko, wenn er je kopiert wurde.' 'The server continues to store old certificates along with their keys. They are useless but occupy space in the store and complicate troubleshooting. The private key remains a risk if it was ever copied.') `
    (T 'Nicht mehr benötigte Zertifikate nach einem Blick, ob sie von Diensten verwendet werden, löschen. Dienste ohne Eintrag in HTTP.sys (z. B. SMTP) vorher prüfen.' 'Delete certificates that are no longer needed after checking whether they are used by services. Check services without an entry in HTTP.sys (e.g., SMTP) first.')

Befund 'Z4' (T 'Verschlüsselung' 'Encryption') (T 'Gültige Zertifikate mit SHA-1-/MD5-Signatur oder RSA-Schlüssel unter 2048 Bit' 'Valid certificates with SHA-1/MD5 signatures or RSA keys under 2048 bits') 'Mittel' `
    @($gueltig | Where-Object { $_.Signatur -match 'sha1|md5|md2|md4' -or ($_.Algorithmus -eq 'RSA' -and $_.Bits -and $_.Bits -lt 2048) } | ForEach-Object {
            Zeile $_ @{ Signatur = $_.Signatur; Schlüssel = $(if ($_.Bits) { (T "$($_.Algorithmus) $($_.Bits) Bit" "$($_.Algorithmus) $($_.Bits) bits") } else { $_.Algorithmus }) } }) `
    (T 'SHA-1, MD5 und kurze RSA-Schlüssel gelten als unsicher. Moderne Browser und Programme lehnen solche Zertifikate ab, ältere Verfahren lassen sich fälschen oder knacken.' 'SHA-1, MD5, and short RSA keys are considered insecure. Modern browsers and programs reject such certificates; older methods can be forged or cracked.') `
    (T 'Zertifikat mit SHA-256 und mindestens 2048 Bit (besser 3072 Bit oder ECDSA) neu ausstellen. Bei eigener Zertifizierungsstelle die Vorlage anpassen.' 'Reissue the certificate with SHA-256 and at least 2048 bits (preferably 3072 bits or ECDSA). If you have your own certification authority, adjust the template.')

Befund 'Z5' (T 'Vertrauen' 'Trust') (T 'Selbstsignierte Zertifikate in Gebrauch' 'Self-signed certificates in use') 'Info' `
    @($gueltig | Where-Object { $_.Verwendung -and $_.Subject -eq $_.Issuer } | ForEach-Object { Zeile $_ }) `
    (T 'Selbstsignierte Zertifikate werden von keinem Gerät automatisch als vertrauenswürdig behandelt. Benutzer gewöhnen sich an Warnungen, und Angreifer können eigene Zertifikate unterschieben. Remotedesktop erzeugt ein solches Zertifikat von sich aus; der Bericht zählt es nur, wenn es als Remotedesktop-Zertifikat festgelegt ist.' 'Self-signed certificates are not automatically trusted by any device. Users become accustomed to warnings, and attackers can insert their own certificates. Remote Desktop generates such a certificate by default; the report counts it only if it is set as a Remote Desktop certificate.') `
    (T 'Für Dienste, die Benutzer ansprechen, ein Zertifikat der eigenen oder einer öffentlichen Zertifizierungsstelle verwenden und für Remotedesktop per Gruppenrichtlinie festlegen.' 'For services that users access, use a certificate from your own or a public certification authority and set it for Remote Desktop via Group Policy.')

$ccsZeilen = @($ccsRechner | ForEach-Object { [pscustomobject][ordered]@{ Computer = $_; Grund = (T 'Zentraler Zertifikatspeicher (CCS) ist in Gebrauch: die dort verwendeten Zertifikate werden nicht zugeordnet und können als "ohne Verwendung" erscheinen' 'Central Certificate Store (CCS) is in use: certificates used there are not assigned and may appear as "unused"') } })
Befund 'Z6' (T 'Abfrage' 'Query') (T 'Computer oder Speicher, die nicht abgefragt werden konnten' 'Computers or stores that could not be queried') 'Info' (@($nichtGeprueft.ToArray()) + $ccsZeilen) `
    (T 'Für diese Computer fehlt die Auswertung. Sie sind in den übrigen Prüfpunkten nicht enthalten, das heißt nicht, dass dort alles in Ordnung ist.' 'Evaluation is missing for these computers. They are not included in the other checks, which does not mean that everything is fine there.') `
    (T 'WinRM auf dem Computer prüfen (Test-WSMan), Firewallregel und Berechtigung des abfragenden Kontos. Das Skript ändert diese Einstellungen nicht.' 'Check WinRM on the computer (Test-WSMan), firewall rule, and permission of the querying account. The script does not change these settings.')

# Computer, die nicht oder nur teilweise gelesen werden konnten: in den Prüfpunkten Z1-Z5 als "nicht geprüft" führen, nie als "ohne Befund"
$ausgefallen = @($nichtGeprueft | ForEach-Object { [string]$_.Computer } | Sort-Object -Unique)
if ($ausgefallen.Count) { NichtGeprueft '^Z[1-5]$' $ausgefallen 'Z6' 'Computer' (T 'Abfrage unvollständig' 'Query incomplete') }
if ($ccsRechner.Count) { NichtGeprueft '^Z[135]$' @($ccsRechner | Sort-Object -Unique) 'Z6' 'Computer' (T 'Zuordnung unvollständig (zentraler Zertifikatspeicher)' 'Incomplete assignment (central certificate store)') }

# ------------------------------------------------------------------ Bericht
$liste = $alle | Sort-Object Computer, Bis | ForEach-Object {
    [pscustomobject][ordered]@{
        Computer = $_.Computer; Zertifikat = Kurz $_.Subject 100; Aussteller = Kurz $_.Issuer 70; Speicher = $_.Speicher; 'Gültig ab' = Datum $_.Von; 'Gültig bis' = Datum $_.Bis
        Status = $(if ($_.Bis -lt $jetzt) { 'abgelaufen' } elseif ($_.Bis -lt $jetzt.AddDays($WarnTage)) { (T 'läuft bald ab' 'expires soon') } else { (T 'gültig' 'valid') })
        Schlüssel = $(if ($_.Schluessel) { 'ja' } else { 'nein' }); Verwendung = $_.Verwendung; Fingerabdruck = $_.Thumbprint
    }
}
$bezug = if ($geprueft.Count -eq 1) { $geprueft[0] } else { "$($geprueft.Count) Computer" }
$umgebung = [ordered]@{
    (T 'Abgefragte Computer' 'Queried computers')       = (T "$($geprueft.Count) von $($geprueft.Count + $nichtErreicht)" "$($geprueft.Count) out of $($geprueft.Count + $nichtErreicht)")
    (T 'Zertifikate (gesamt)' 'Certificates (total)')      = $alle.Count
    (T 'davon abgelaufen' 'expired among them')          = $abgelaufen.Count
    (T 'Zertifikatspeicher' 'Certificate store')        = (T 'Eigene Zertifikate, Remotedesktop, WebHosting (Computerkonto)' 'Own certificates, Remote Desktop, WebHosting (computer account)')
    (T 'Schwelle dieses Berichts' 'Threshold for this report')  = (T "Warnung $WarnTage Tage vor Ablauf" "Warning $WarnTage days before expiration")
}
$basis = BerichtSchreiben -Name (T 'Zertifikate-Bericht' 'Certificate report') -Datei 'Zertifikate-Bericht' -Bezug $bezug -BezugName (T 'Geprüft' 'Checked') -Umgebung $umgebung `
    -Abschnitte ([ordered]@{ (T 'Alle Zertifikate' 'All certificates') = $liste }) `
    -Fuss (T 'Grundlage sind die Zertifikatspeicher der Computer und die Zuordnung zu HTTPS-Bindungen und Remotedesktop. Dienste, die ihr Zertifikat anderswo ablegen (z. B. SMTP, LDAPS), erscheinen ohne den Vermerk "in Gebrauch". Der Bericht ersetzt keine Zertifikatsverwaltung und erhebt keinen Anspruch auf Vollständigkeit.' 'The basis is the computers'' certificate stores and their assignment to HTTPS bindings and Remote Desktop. Services that store their certificates elsewhere (e.g., SMTP, LDAPS) appear without the note "in use". The report does not replace certificate management and makes no claim to completeness.')
if ($Csv) { SpaltenAnzeige $liste | Export-Csv -Path "${basis}_Zertifikate.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8 }
if ($PassThru) { $Befunde }
