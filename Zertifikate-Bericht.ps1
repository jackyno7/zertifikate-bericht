#Requires -Version 5.1
<#
.SYNOPSIS
    Zertifikate-Bericht: liest die Zertifikate auf Windows-Servern aus und meldet abgelaufene, bald ablaufende und
    schwache Zertifikate, besonders solche, die von Webdiensten oder Remotedesktop tatsächlich verwendet werden. Ergebnis
    ist ein HTML-Bericht.

.DESCRIPTION
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
    Pfad der HTML-Datei. Standard: Zertifikate-Bericht_<Computer>_<Datum>.html im aktuellen Ordner.

.PARAMETER Computer
    Namen weiterer Computer, die über WinRM abgefragt werden.

.PARAMETER AusAD
    Fragt alle aktiven Windows-Server aus dem Active Directory ab (Modul ActiveDirectory nötig). Als Server gelten
    Computerkonten mit "Server" im Betriebssystem, die sich in den letzten 60 Tagen angemeldet haben.

.PARAMETER OhneLokalen
    Den Computer, auf dem das Skript läuft, nicht prüfen (sinnvoll, wenn es von einem Arbeitsplatz aus gestartet wird).

.PARAMETER Server
    Bestimmter Domänencontroller für die AD-Abfrage (nur mit -AusAD).

.PARAMETER Anmeldung
    Anmeldedaten (Get-Credential) für die Fernabfrage. Standard: das aktuelle Konto.

.PARAMETER WarnTage
    Zertifikate, die in so vielen Tagen ablaufen, werden gemeldet (Standard 30).

.PARAMETER Csv
    Schreibt zusätzlich je Prüfpunkt mit Treffern eine CSV-Datei und die Liste aller Zertifikate neben den Bericht.

.PARAMETER NichtOeffnen
    Öffnet den fertigen Bericht nicht automatisch im Browser.

.PARAMETER PassThru
    Gibt die Befunde zusätzlich als Objekte zurück (für eigene Auswertungen).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Zertifikate-Bericht.ps1

.EXAMPLE
    .\Zertifikate-Bericht.ps1 -AusAD -WarnTage 60 -Csv -Ausgabe C:\Berichte\zertifikate.html

.NOTES
    IT Admin Tools by MM - Zertifikate-Bericht Version 1.0 (11.2026), Lizenz: MIT (siehe LICENSE)
    Kein Ersatz für eine Zertifikatsverwaltung. Nutzung auf eigene Verantwortung.
#>
[CmdletBinding()]
param(
    [string]$Ausgabe,
    [string[]]$Computer,
    [switch]$AusAD,
    [switch]$OhneLokalen,
    [string]$Server,
    [System.Management.Automation.PSCredential]$Anmeldung,
    [ValidateRange(1, 730)][int]$WarnTage = 30,
    [switch]$Csv,
    [switch]$NichtOeffnen,
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'
$Version = '1.0'
# Kommalisten: Beim Start über "powershell -File" kommt "a,b" als ein Wert an; hier in einzelne Namen trennen.
if ($PSBoundParameters.ContainsKey('Computer')) {
    $Computer = @($Computer | ForEach-Object { $_ -split '\s*,\s*' } | Where-Object { $_ })
    if (-not $Computer.Count) { throw 'Bei -Computer wurde kein Computer angegeben. Ohne den Parameter wird nur der lokale Computer geprüft.' }
}

# ------------------------------------------------------------------ Hilfsfunktionen
function Datum($d, $leer = 'nie') { if ($d) { ([datetime]$d).ToString('dd.MM.yyyy') } else { $leer } }
function JaNein($b) { if ($b) { 'Ja' } else { 'Nein' } }
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
        $sicher | Microsoft.PowerShell.Utility\Export-Csv -Path $Path -NoTypeInformation:$NoTypeInformation -Delimiter $Delimiter -Encoding $Encoding
    }
}

# Tabellen: ConvertTo-Html und Export-Csv nehmen die Spalten des ersten Objekts. Haben die Zeilen einer Liste verschiedene
# Eigenschaften (z. B. Fehlzeile neben normalen Treffern), gingen Spalten verloren: alle Namen der Reihe nach vereinigen.
function Einheitlich($liste) {
    $liste = @($liste | Where-Object { $null -ne $_ })
    if ($liste.Count -lt 2) { return $liste }
    $namen = New-Object System.Collections.Generic.List[string]
    foreach ($o in $liste) { foreach ($p in $o.PSObject.Properties.Name) { if (-not $namen.Contains($p)) { $namen.Add($p) } } }
    if ($namen.Count -eq @($liste[0].PSObject.Properties.Name).Count) { return $liste }
    @($liste | Select-Object -Property $namen)
}
# Sortierschluessel fuer Nummern wie K2, K10: erst Buchstaben, dann die Zahl als Zahl
function IdZahl($id) { $m = [regex]::Match([string]$id, '\d+'); if ($m.Success) { [int]$m.Value } else { 0 } }
function IdText($id) { [regex]::Replace([string]$id, '\d+', '') }

$Befunde = New-Object System.Collections.Generic.List[object]

# Platzhalterzeilen ("nicht geprüft") erkennt das Skript nur an den vollständigen, festen Texten, die es selbst erzeugt - keine Teilstring-Suche
# in Konto-, Objekt- oder Beschreibungstexten, die ein Angreifer beeinflussen könnte
$PlatzhalterTexte = @('Abfrage fehlgeschlagen - nicht geprüft', 'Berechtigungen nicht lesbar - nicht geprüft', 'SYSVOL nicht lesbar - nicht geprüft', 'WMI-Filter nicht lesbar - nicht geprüft', 'Kein Domänencontroller ermittelbar - SYSVOL nicht geprüft', 'Inhalt des Ordners nicht lesbar - Richtlinie nicht geprüft')
function IstPlatzhalter($o) {
    # Ein Skript kann $script:PlatzhalterSpalten setzen: dann zählen nur diese Spalten (Namensspalten sind von Benutzern beeinflussbar)
    foreach ($pr in $o.PSObject.Properties) {
        if ($script:PlatzhalterSpalten -and ($pr.Name -notin $script:PlatzhalterSpalten)) { continue }
        if (($pr.Value -is [string]) -and ($PlatzhalterTexte -ccontains $pr.Value)) { return $true }
    }
    ("$($o.Feststellung)" -clike 'Nicht geprüft - *')
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
function NichtGeprueft([string]$IdMuster, [string[]]$Namen, [string]$Verweis, [string]$Spalte = 'Server', [string]$Grund = 'Abfrage nicht möglich') {
    foreach ($x in ($Befunde | Where-Object { $_.Id -match $IdMuster })) {
        $war = $x.Anzahl
        $zeilen = @($Namen | ForEach-Object { [pscustomobject][ordered]@{ $Spalte = $_; Feststellung = "Nicht geprüft - $Grund (siehe $Verweis)"; Details = '' } })
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
    param([string]$Name, [string]$Datei, [string]$Bezug, [string]$BezugName = 'Domäne', [System.Collections.IDictionary]$Umgebung,
        [System.Collections.IDictionary]$Abschnitte = @{}, [System.Collections.IDictionary]$Anhang = @{}, [string]$Fuss)
    Write-Host '  Bericht schreiben ...'
    $jetzt = Get-Date
    $ziel = $script:Ausgabe
    if (-not $ziel) { $ziel = Join-Path (Get-Location).Path ("{0}_{1}_{2:yyyy-MM-dd_HHmm}.html" -f $Datei, ($Bezug -replace '[^\w.\-]+', '-'), $jetzt) }
    $ziel = [System.IO.Path]::GetFullPath($ziel)
    $ordner = Split-Path $ziel -Parent
    if (-not (Test-Path $ordner)) { New-Item -ItemType Directory -Path $ordner -Force | Out-Null }

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
    [void]$h.Append("<!DOCTYPE html><html lang='de'><head><meta charset='utf-8'><title>$(Html $Name) $(Html $Bezug)</title><style>$css</style></head><body><div class='seite'>")
    [void]$h.Append("<h1>$(Html $Name)</h1><div class='meta'>$(Html $BezugName) <b>$(Html $Bezug)</b> · erstellt am $($jetzt.ToString('dd.MM.yyyy HH:mm')) von $(Html "$env:USERDOMAIN\$env:USERNAME") auf $(Html $env:COMPUTERNAME)</div>")
    [void]$h.Append("<div class='kacheln'>")
    foreach ($s in 'Hoch', 'Mittel', 'Niedrig', 'Info', 'OK') {
        $text = @{ Hoch = 'Befunde mit hoher Priorität'; Mittel = 'mit mittlerer Priorität'; Niedrig = 'mit niedriger Priorität'; Info = 'Hinweise und nicht geprüft'; OK = 'Prüfpunkte ohne Befund' }[$s]
        [void]$h.Append("<div class='kachel $s'><b>$($zahl[$s])</b><span>$text</span></div>")
    }
    [void]$h.Append("</div><h2>Umgebung</h2><table><tr><th>Eigenschaft</th><th>Wert</th></tr>")
    foreach ($k in $Umgebung.Keys) { [void]$h.Append("<tr><td>$(Html $k)</td><td>$(Html $Umgebung[$k])</td></tr>") }
    [void]$h.Append("</table>")

    [void]$h.Append("<h2>Übersicht der Prüfpunkte</h2><table><tr><th>Nr.</th><th>Bereich</th><th>Prüfpunkt</th><th>Treffer</th><th>Bewertung</th></tr>")
    foreach ($x in $Befunde) {
        [void]$h.Append("<tr><td>$($x.Id)</td><td>$(Html $x.Bereich)</td><td><a href='#$($x.Id)'>$(Html $x.Titel)</a></td><td>$($x.Anzahl)</td><td><span class='marke m-$($x.Stufe)'>$($x.Stufe)</span></td></tr>")
    }
    [void]$h.Append("</table>")

    foreach ($k in $Abschnitte.Keys) {
        [void]$h.Append("<h2>$(Html $k)</h2>")
        $inhalt = $Abschnitte[$k]
        if ($inhalt -is [string]) { [void]$h.Append($inhalt) }
        elseif (@($inhalt).Count) { [void]$h.Append(((Einheitlich $inhalt | ConvertTo-Html -Fragment) -join "`n")) }
        else { [void]$h.Append("<p class='klein'>Keine Einträge.</p>") }
    }

    [void]$h.Append("<h2>Befunde</h2>")
    $offen = @($Befunde | Where-Object { $_.Stufe -ne 'OK' } | Sort-Object { $rang[$_.Stufe] }, { IdText $_.Id }, { IdZahl $_.Id })
    foreach ($x in $offen) {
        [void]$h.Append("<div class='befund $($x.Stufe)' id='$($x.Id)'><h3><span class='marke m-$($x.Stufe)'>$($x.Stufe)</span> $($x.Id) · $(Html $x.Titel) <span class='klein'>($($x.Anzahl) Treffer)</span></h3>")
        [void]$h.Append("<p>$(Html $x.Erklaerung)</p><p class='emp'><b>Empfehlung:</b> $(Html $x.Empfehlung)</p>")
        [void]$h.Append(((Einheitlich $x.Objekte | ConvertTo-Html -Fragment) -join "`n"))
        [void]$h.Append("</div>")
    }
    if (-not $offen.Count) { [void]$h.Append("<p>Keine Auffälligkeiten gefunden.</p>") }
    $ok = @($Befunde | Where-Object Stufe -eq 'OK')
    if ($ok.Count) {
        [void]$h.Append("<h2>Ohne Befund</h2><table><tr><th>Nr.</th><th>Prüfpunkt</th></tr>")
        foreach ($x in $ok) { [void]$h.Append("<tr id='$($x.Id)'><td>$($x.Id)</td><td>$(Html $x.Titel)</td></tr>") }
        [void]$h.Append("</table>")
    }
    foreach ($k in $Anhang.Keys) { [void]$h.Append("<h2>$(Html $k)</h2><pre>$(Html $Anhang[$k])</pre>") }
    [void]$h.Append("<div class='fuss'>$(Html $Name) Version $script:Version · IT Admin Tools by MM · Das Skript hat ausschließlich gelesen und nichts verändert. $(Html $Fuss)</div>")
    [void]$h.Append("</div></body></html>")
    [System.IO.File]::WriteAllText($ziel, $h.ToString(), (New-Object System.Text.UTF8Encoding($true)))

    $basis = [System.IO.Path]::ChangeExtension($ziel, $null).TrimEnd('.')
    if ($script:Csv) {
        foreach ($x in ($Befunde | Where-Object { $_.Anzahl -gt 0 })) {
            Einheitlich $x.Objekte | Export-Csv -Path "${basis}_$($x.Id).csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8
        }
    }
    Write-Host ''
    Write-Host ("Fertig: {0} hoch, {1} mittel, {2} niedrig, {3} ohne Befund." -f $zahl.Hoch, $zahl.Mittel, $zahl.Niedrig, $zahl.OK) -ForegroundColor Green
    Write-Host "Bericht: $ziel"
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
    catch { $ergebnis.Fehler += "HTTPS-Bindungen nicht lesbar: $($_.Exception.Message)" }
    try {
        $rdp = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name SSLCertificateSHA1Hash -ErrorAction SilentlyContinue).SSLCertificateSHA1Hash
        if ($rdp) { $hex = (($rdp | ForEach-Object { $_.ToString('X2') }) -join ''); $gebrauch[$hex] = (@($gebrauch[$hex]) + 'Remotedesktop') | Where-Object { $_ } }
    }
    catch { $ergebnis.Fehler += "Remotedesktop-Zertifikat nicht lesbar: $($_.Exception.Message)" }

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
        catch { $ergebnis.Fehler += "Speicher $speicher nicht lesbar: $($_.Exception.Message)" }
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
        Write-Host 'Das PowerShell-Modul "ActiveDirectory" fehlt (für -AusAD nötig).' -ForegroundColor Red
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
    Write-Host "Lese Zertifikate von $lokal ..." -ForegroundColor Cyan
    try { Aufnehmen $lokal (& $lese) }
    catch { $nichtGeprueft.Add([pscustomobject][ordered]@{ Computer = $lokal; Grund = $_.Exception.Message }) }
}
if ($fern.Count) {
    Write-Host "Frage $($fern.Count) Computer über WinRM ab ..." -ForegroundColor Cyan
    $ic = @{ ComputerName = $fern; ScriptBlock = $lese; ErrorAction = 'SilentlyContinue'; ErrorVariable = 'icFehler' }
    if ($Anmeldung) { $ic.Credential = $Anmeldung }
    $icFehler = @()
    $antworten = @(Invoke-Command @ic)
    $erhalten = @{}
    foreach ($a in $antworten) { $n = [string]$a.PSComputerName; $erhalten[$n.Split('.')[0].ToLowerInvariant()] = $true; Aufnehmen $n $a }
    foreach ($f in $fern) {
        if (-not $erhalten.ContainsKey($f.Split('.')[0].ToLowerInvariant())) {
            $grund = @($icFehler | Where-Object { ([string]$_.TargetObject) -like "$($f.Split('.')[0])*" -or ([string]$_.OriginInfo.PSComputerName) -like "$($f.Split('.')[0])*" } | ForEach-Object { ($_.Exception.Message -replace '\s+', ' ') } | Select-Object -First 1)
            if (-not $grund) { $grund = 'keine Antwort' }
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

Befund 'Z1' 'Ablauf' 'Abgelaufenes Zertifikat wird verwendet (HTTPS oder Remotedesktop)' 'Hoch' `
    @($abgelaufen | Where-Object Verwendung | ForEach-Object { Zeile $_ @{ 'Tage seit Ablauf' = [int]($jetzt - $_.Bis).TotalDays } }) `
    'Ein Dienst liefert weiter ein abgelaufenes Zertifikat aus. Browser und Programme zeigen Warnungen oder verweigern die Verbindung, Anmeldungen über Remotedesktop zeigen eine Sicherheitswarnung.' `
    'Zertifikat erneuern und in der Bindung des Dienstes austauschen. Für die Zukunft eine Erinnerung oder automatische Erneuerung einrichten.'

Befund 'Z2' 'Ablauf' "Zertifikate laufen in den nächsten $WarnTage Tagen ab" 'Mittel' `
    @($gueltig | Where-Object { $_.Bis -lt $jetzt.AddDays($WarnTage) -and $_.Schluessel } | Sort-Object Bis | ForEach-Object { Zeile $_ @{ 'Tage bis Ablauf' = [int]($_.Bis - $jetzt).TotalDays } }) `
    'Läuft ein Zertifikat ab, fallen die damit gesicherten Dienste aus, oft am Wochenende und ohne Vorwarnung. Betroffen sind nur Zertifikate mit privatem Schlüssel, also solche, die ein Server selbst verwendet.' `
    'Erneuerung rechtzeitig einplanen. Bei öffentlichen Zertifikaten die Laufzeit der Bestellung beachten.'

Befund 'Z3' 'Aufräumen' 'Abgelaufene Zertifikate mit privatem Schlüssel, die nicht verwendet werden' 'Niedrig' `
    @($abgelaufen | Where-Object { $_.Schluessel -and -not $_.Verwendung } | Sort-Object Bis | ForEach-Object { Zeile $_ @{ 'Tage seit Ablauf' = [int]($jetzt - $_.Bis).TotalDays } }) `
    'Der Server speichert alte Zertifikate samt Schlüssel weiter. Sie sind nutzlos, belegen aber Platz im Speicher und verwirren bei der Fehlersuche. Der private Schlüssel bleibt ein Risiko, wenn er je kopiert wurde.' `
    'Nicht mehr benötigte Zertifikate nach einem Blick, ob sie von Diensten verwendet werden, löschen. Dienste ohne Eintrag in HTTP.sys (z. B. SMTP) vorher prüfen.'

Befund 'Z4' 'Verschlüsselung' 'Gültige Zertifikate mit SHA-1-/MD5-Signatur oder RSA-Schlüssel unter 2048 Bit' 'Mittel' `
    @($gueltig | Where-Object { $_.Signatur -match 'sha1|md5|md2|md4' -or ($_.Algorithmus -eq 'RSA' -and $_.Bits -and $_.Bits -lt 2048) } | ForEach-Object {
            Zeile $_ @{ Signatur = $_.Signatur; Schlüssel = $(if ($_.Bits) { "$($_.Algorithmus) $($_.Bits) Bit" } else { $_.Algorithmus }) } }) `
    'SHA-1, MD5 und kurze RSA-Schlüssel gelten als unsicher. Moderne Browser und Programme lehnen solche Zertifikate ab, ältere Verfahren lassen sich fälschen oder knacken.' `
    'Zertifikat mit SHA-256 und mindestens 2048 Bit (besser 3072 Bit oder ECDSA) neu ausstellen. Bei eigener Zertifizierungsstelle die Vorlage anpassen.'

Befund 'Z5' 'Vertrauen' 'Selbstsignierte Zertifikate in Gebrauch' 'Info' `
    @($gueltig | Where-Object { $_.Verwendung -and $_.Subject -eq $_.Issuer } | ForEach-Object { Zeile $_ }) `
    'Selbstsignierte Zertifikate werden von keinem Gerät automatisch als vertrauenswürdig behandelt. Benutzer gewöhnen sich an Warnungen, und Angreifer können eigene Zertifikate unterschieben. Remotedesktop erzeugt ein solches Zertifikat von sich aus; der Bericht zählt es nur, wenn es als Remotedesktop-Zertifikat festgelegt ist.' `
    'Für Dienste, die Benutzer ansprechen, ein Zertifikat der eigenen oder einer öffentlichen Zertifizierungsstelle verwenden und für Remotedesktop per Gruppenrichtlinie festlegen.'

$ccsZeilen = @($ccsRechner | ForEach-Object { [pscustomobject][ordered]@{ Computer = $_; Grund = 'Zentraler Zertifikatspeicher (CCS) ist in Gebrauch: die dort verwendeten Zertifikate werden nicht zugeordnet und können als "ohne Verwendung" erscheinen' } })
Befund 'Z6' 'Abfrage' 'Computer oder Speicher, die nicht abgefragt werden konnten' 'Info' (@($nichtGeprueft.ToArray()) + $ccsZeilen) `
    'Für diese Computer fehlt die Auswertung. Sie sind in den übrigen Prüfpunkten nicht enthalten, das heißt nicht, dass dort alles in Ordnung ist.' `
    'WinRM auf dem Computer prüfen (Test-WSMan), Firewallregel und Berechtigung des abfragenden Kontos. Das Skript ändert diese Einstellungen nicht.'

# Computer, die nicht oder nur teilweise gelesen werden konnten: in den Prüfpunkten Z1-Z5 als "nicht geprüft" führen, nie als "ohne Befund"
$ausgefallen = @($nichtGeprueft | ForEach-Object { [string]$_.Computer } | Sort-Object -Unique)
if ($ausgefallen.Count) { NichtGeprueft '^Z[1-5]$' $ausgefallen 'Z6' 'Computer' 'Abfrage unvollständig' }
if ($ccsRechner.Count) { NichtGeprueft '^Z[135]$' @($ccsRechner | Sort-Object -Unique) 'Z6' 'Computer' 'Zuordnung unvollständig (zentraler Zertifikatspeicher)' }

# ------------------------------------------------------------------ Bericht
$liste = $alle | Sort-Object Computer, Bis | ForEach-Object {
    [pscustomobject][ordered]@{
        Computer = $_.Computer; Zertifikat = Kurz $_.Subject 100; Aussteller = Kurz $_.Issuer 70; Speicher = $_.Speicher; 'Gültig ab' = Datum $_.Von; 'Gültig bis' = Datum $_.Bis
        Status = $(if ($_.Bis -lt $jetzt) { 'abgelaufen' } elseif ($_.Bis -lt $jetzt.AddDays($WarnTage)) { 'läuft bald ab' } else { 'gültig' })
        Schlüssel = $(if ($_.Schluessel) { 'ja' } else { 'nein' }); Verwendung = $_.Verwendung; Fingerabdruck = $_.Thumbprint
    }
}
$bezug = if ($geprueft.Count -eq 1) { $geprueft[0] } else { "$($geprueft.Count) Computer" }
$umgebung = [ordered]@{
    'Abgefragte Computer'       = "$($geprueft.Count) von $($geprueft.Count + $nichtErreicht)"
    'Zertifikate (gesamt)'      = $alle.Count
    'davon abgelaufen'          = $abgelaufen.Count
    'Zertifikatspeicher'        = 'Eigene Zertifikate, Remotedesktop, WebHosting (Computerkonto)'
    'Schwelle dieses Berichts'  = "Warnung $WarnTage Tage vor Ablauf"
}
$basis = BerichtSchreiben -Name 'Zertifikate-Bericht' -Datei 'Zertifikate-Bericht' -Bezug $bezug -BezugName 'Geprüft' -Umgebung $umgebung `
    -Abschnitte ([ordered]@{ 'Alle Zertifikate' = $liste }) `
    -Fuss 'Grundlage sind die Zertifikatspeicher der Computer und die Zuordnung zu HTTPS-Bindungen und Remotedesktop. Dienste, die ihr Zertifikat anderswo ablegen (z. B. SMTP, LDAPS), erscheinen ohne den Vermerk "in Gebrauch". Der Bericht ersetzt keine Zertifikatsverwaltung und erhebt keinen Anspruch auf Vollständigkeit.'
if ($Csv) { $liste | Export-Csv -Path "${basis}_Zertifikate.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8 }
if ($PassThru) { $Befunde }
