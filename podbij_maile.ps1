# ============================================================
#  BOT DO PODBIJANIA MAILI (follow-up) - zespolu
#  1. Sprawdza skrzynke zespolu: kto odpowiedzial -> wpisuje TAK w kol. 3
#  2. Kto nie odpowiedzial -> podbija oryginalny mail (RE: w tym samym watku)
#     krotkim przypomnieniem z pliku followup.docx
#  Uruchomienie: PODBIJ.bat (wysylka) lub PODBIJ_PODGLAD.bat (tylko szkice)
# ============================================================

param([switch]$Podglad)

$ErrorActionPreference = 'Stop'

# --- USTAWIENIA ---
$Folder        = Split-Path -Parent $MyInvocation.MyCommand.Path
$PlikFollowup  = Join-Path $Folder 'followup.docx'
$PlikExcel     = Join-Path $Folder 'maile.xlsx'
$Placeholder   = '[place holder z excela]'
$Nadawca = $env:OUTREACH_MAILBOX
if (-not $Nadawca) { throw 'Ustaw zmienna OUTREACH_MAILBOX na adres skrzynki wspoldzielonej (patrz README).' }
$KolOdpisal    = 3     # kolumna 'czy odpisal'
$KolWyslano    = 4     # kolumna 'wyslano'
$KolPodbito    = 5     # kolumna 'podbito' (tworzy sie sama)
$MinDniOdWyslania = 3  # nie podbijaj wczesniej niz X dni od wyslania
$PauzaSekund   = 3


if (-not (Test-Path $PlikFollowup)) { Write-Host "BLAD: brak pliku followup.docx" -ForegroundColor Red; pause; exit 1 }
if (-not (Test-Path $PlikExcel))    { Write-Host "BLAD: brak pliku maile.xlsx"   -ForegroundColor Red; pause; exit 1 }

# --- wczytaj tekst przypomnienia z followup.docx (bez Worda - prosto z pliku) ---
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($PlikFollowup)
$entry = $zip.GetEntry('word/document.xml')
$sr = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
$docXml = [xml]$sr.ReadToEnd()
$sr.Close(); $zip.Dispose()

$nsm = New-Object System.Xml.XmlNamespaceManager($docXml.NameTable)
$nsm.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$AkapityFollowup = @()
foreach ($p in $docXml.SelectNodes('//w:body/w:p', $nsm)) {
    $t = ($p.SelectNodes('.//w:t', $nsm) | ForEach-Object { $_.InnerText }) -join ''
    $AkapityFollowup += $t
}
if (-not ($AkapityFollowup -join '')) { Write-Host "BLAD: followup.docx jest pusty" -ForegroundColor Red; pause; exit 1 }
Write-Host ("Wczytano przypomnienie: {0} akapitow" -f $AkapityFollowup.Count)

function Parse-DataWyslania($val, $text) {
    if ($val -is [datetime]) { return $val }
    if ($val -is [double] -or $val -is [int]) { try { return [datetime]::FromOADate([double]$val) } catch { } }
    if ($text) {
        try { return [datetime]::Parse([string]$text, [System.Globalization.CultureInfo]::GetCultureInfo('pl-PL')) } catch { }
        try { return [datetime]::Parse([string]$text) } catch { }
    }
    return $null
}

# --- EXCEL ---
Write-Host "Wczytuje liste z maile.xlsx..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($PlikExcel)
$ws = $wb.Worksheets.Item(1)
if (-not $ws.Cells.Item(1, $KolPodbito).Value2) { $ws.Cells.Item(1, $KolPodbito).Value2 = 'podbito' }

$Lista = @()
$row = 2
while ($ws.Cells.Item($row, 1).Value2) {
    $Lista += [PSCustomObject]@{
        Wiersz  = $row
        Adres   = ([string]$ws.Cells.Item($row, 1).Value2).Trim()
        Zwrot   = $(if ($ws.Cells.Item($row, 2).Value2) { ([string]$ws.Cells.Item($row, 2).Value2).Trim() } else { '' })
        Odpisal = [string]$ws.Cells.Item($row, $KolOdpisal).Value2
        Wyslano = $ws.Cells.Item($row, $KolWyslano).Value2
        WyslanoData = $ws.Cells.Item($row, $KolWyslano).Text
        DataWyslania = Parse-DataWyslania $ws.Cells.Item($row, $KolWyslano).Value2 $ws.Cells.Item($row, $KolWyslano).Text
        Podbito = [string]$ws.Cells.Item($row, $KolPodbito).Value2
    }
    $row++
}

# granica skanowania skrzynek: najstarsza data wysylki minus 1 dzien
$DatyWyslania = @($Lista | Where-Object { $_.DataWyslania } | ForEach-Object { $_.DataWyslania })
if ($DatyWyslania.Count -gt 0) { $Cutoff = ($DatyWyslania | Sort-Object)[0].AddDays(-1) } else { $Cutoff = (Get-Date).AddDays(-30) }

# --- OUTLOOK ---
if (-not (Get-Process OUTLOOK -ErrorAction SilentlyContinue)) {
    Write-Host "Outlook nie jest uruchomiony - odpalam go i czekam..." -ForegroundColor Yellow
    Start-Process outlook.exe
    $czekaj = 0
    while (-not (Get-Process OUTLOOK -ErrorAction SilentlyContinue) -and $czekaj -lt 60) { Start-Sleep -Seconds 2; $czekaj += 2 }
    Start-Sleep -Seconds 15
}
Write-Host "Lacze sie z Outlookiem..."
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.Session

$recip = $ns.CreateRecipient($Nadawca)
$null = $recip.Resolve()
if (-not $recip.Resolved) { Write-Host "BLAD: Outlook nie rozpoznaje '$Nadawca'." -ForegroundColor Red; $wb.Close($false); $excel.Quit(); pause; exit 1 }

# foldery skrzynki wspoldzielonej zespolu
$inbox = $null; $sentf = $null
try { $inbox = $ns.GetSharedDefaultFolder($recip, 6) } catch { }   # 6 = skrzynka odbiorcza
try { $sentf = $ns.GetSharedDefaultFolder($recip, 5) } catch { }   # 5 = elementy wyslane
if ($inbox -and -not $sentf) {
    try { $sentf = $inbox.Store.GetDefaultFolder(5) } catch { }
}
if (-not $inbox) { Write-Host "BLAD: nie znalazlem skrzynki odbiorczej zespolu." -ForegroundColor Red; $wb.Close($false); $excel.Quit(); pause; exit 1 }
if (-not $sentf) { Write-Host "BLAD: nie znalazlem elementow wyslanych zespolu." -ForegroundColor Red; $wb.Close($false); $excel.Quit(); pause; exit 1 }
Write-Host ("Foldery OK: '{0}' / '{1}'" -f $inbox.FolderPath, $sentf.FolderPath)

function SmtpNadawcy($it) {
    try {
        if ($it.SenderEmailType -eq 'SMTP') { return ([string]$it.SenderEmailAddress).ToLower().Trim() }
        $ex = $it.Sender.GetExchangeUser()
        if ($ex) { return ([string]$ex.PrimarySmtpAddress).ToLower().Trim() }
    } catch { }
    return ''
}
function SmtpAdresata($r) {
    try {
        $p = $r.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x39FE001E')
        if ($p) { return ([string]$p).ToLower().Trim() }
    } catch { }
    try { return ([string]$r.Address).ToLower().Trim() } catch { return '' }
}

# mapa: kto odpowiedzial (nadawca -> data ostatniego maila w skrzynce zespolu)
Write-Host ("Skanuje skrzynke odbiorcza zespolu (maile od {0})..." -f $Cutoff.ToString('yyyy-MM-dd'))
$Odpowiedzi = @{}
$itemsIn = $inbox.Items
$itemsIn.Sort('[ReceivedTime]', $true)   # od najnowszych
foreach ($it in $itemsIn) {
    try {
        if ($it.Class -ne 43) { continue }   # 43 = zwykly mail
        if ($it.ReceivedTime -lt $Cutoff) { break }
        $s = SmtpNadawcy $it
        if ($s -and (-not $Odpowiedzi.ContainsKey($s) -or $it.ReceivedTime -gt $Odpowiedzi[$s])) { $Odpowiedzi[$s] = $it.ReceivedTime }
    } catch { }
}
Write-Host ("Znaleziono maile od {0} roznych nadawcow." -f $Odpowiedzi.Count)

# mapa: wyslane oryginaly (adresat -> najnowszy wyslany mail)
Write-Host "Skanuje elementy wyslane zespolu (oryginaly)..."
$Oryginaly = @{}
$itemsOut = $sentf.Items
$itemsOut.Sort('[SentOn]', $true)        # od najnowszych
foreach ($it in $itemsOut) {
    try {
        if ($it.Class -ne 43) { continue }
        if ($it.SentOn -lt $Cutoff) { break }
        foreach ($r in $it.Recipients) {
            $a = SmtpAdresata $r
            if ($a -and (-not $Oryginaly.ContainsKey($a) -or $it.SentOn -gt $Oryginaly[$a].SentOn)) { $Oryginaly[$a] = $it }
        }
    } catch { }
}

# --- ANALIZA ---
$Odpisali   = @()
$DoPodbicia = @()
$ZaSwieze   = @()
$Teraz = Get-Date
foreach ($o in $Lista) {
    if (-not $o.Wyslano) { continue }                       # jeszcze nie wyslano oryginalu
    $a = $o.Adres.ToLower()
    if ($Odpowiedzi.ContainsKey($a)) {
        if (-not $o.Odpisal) { $Odpisali += $o }            # do oznaczenia TAK
        continue
    }
    if ($o.Odpisal) { continue }                            # oznaczony recznie
    if ($o.Podbito) { continue }                            # juz podbity
    if (-not $o.DataWyslania) { $ZaSwieze += $o; continue }   # nieczytelna data -> bezpieczniej NIE podbijac
    if (($Teraz - $o.DataWyslania).TotalDays -lt $MinDniOdWyslania) { $ZaSwieze += $o; continue }
    $DoPodbicia += $o
}

Write-Host ""
Write-Host "=== PODSUMOWANIE ===" -ForegroundColor Cyan
Write-Host ("Nowe odpowiedzi (oznacze TAK w excelu): {0}" -f $Odpisali.Count) -ForegroundColor Green
foreach ($o in $Odpisali) { Write-Host ("  TAK: {0}" -f $o.Adres) -ForegroundColor Green }
Write-Host ("Za swieze na podbicie (mniej niz {0} dni): {1}" -f $MinDniOdWyslania, $ZaSwieze.Count)
Write-Host ("Do podbicia: {0}" -f $DoPodbicia.Count) -ForegroundColor Yellow
foreach ($o in $DoPodbicia) { Write-Host ("  -> {0} (wyslano {1})" -f $o.Adres, $o.WyslanoData) }
Write-Host ""

# oznacz odpowiedzi w excelu (takze w podgladzie - to tylko odnotowanie faktow)
foreach ($o in $Odpisali) {
    $ws.Cells.Item($o.Wiersz, $KolOdpisal).Value2 = ('TAK ' + (Get-Date $Odpowiedzi[$o.Adres.ToLower()] -Format 'yyyy-MM-dd'))
}
if ($Odpisali.Count -gt 0) { $wb.Save() }

if ($DoPodbicia.Count -eq 0) {
    Write-Host "Nie ma nic do podbicia." -ForegroundColor Green
    $wb.Close($true); $excel.Quit(); pause; exit 0
}

if ($Podglad) {
    Write-Host "TRYB PODGLADU - podbicia otworza sie jako szkice, NIC nie zostanie wyslane." -ForegroundColor Yellow
} else {
    $odp = Read-Host "Wpisz T i Enter, zeby WYSLAC podbicia (cokolwiek innego = anuluj)"
    if ($odp -ne 'T' -and $odp -ne 't') { Write-Host "Anulowano."; $wb.Close($true); $excel.Quit(); exit 0 }
}

# --- PODBIJANIE ---
$Wyslane = 0
$Bledy   = 0
foreach ($o in $DoPodbicia) {
    try {
        $a = $o.Adres.ToLower()
        if (-not $Oryginaly.ContainsKey($a)) {
            Write-Host ("POMIJAM {0} - nie znalazlem oryginalu w elementach wyslanych zespolu" -f $o.Adres) -ForegroundColor Yellow
            continue
        }
        Write-Host ("Podbijam {0}..." -f $o.Adres)
        $org = $Oryginaly[$a]
        $fwd = $org.Forward()                      # forward = cytat oryginalu + zalaczniki (one-pager idzie ponownie)
        $fwd.Subject = 'RE: ' + $org.Subject       # zeby weszlo w ten sam watek jako przypomnienie
        $fwd.To      = $o.Adres

        # zbuduj HTML przypomnienia (placeholder -> zwrot z excela) i wstaw NAD cytowanym oryginalem
        $htmlAkapity = foreach ($ak in $AkapityFollowup) {
            $tekst = $ak.Replace($Placeholder, $o.Zwrot)
            $tekst = [System.Net.WebUtility]::HtmlEncode($tekst)
            if ($tekst.Trim()) { '<p style="margin:0 0 12px 0">' + $tekst + '</p>' } else { '<br>' }
        }
        $prefiks = '<div style="font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt">' + ($htmlAkapity -join '') + '</div>'
        $fwd.HTMLBody = $prefiks + $fwd.HTMLBody

        $fwd.SentOnBehalfOfName = $Nadawca

        if ($Podglad) {
            $fwd.Display()
            Write-Host ("PODGLAD: {0}" -f $o.Adres) -ForegroundColor Yellow
        } else {
            $fwd.Send()
            $ws.Cells.Item($o.Wiersz, $KolPodbito).Value2 = (Get-Date -Format 'yyyy-MM-dd HH:mm')
            $wb.Save()
            $Wyslane++
            Write-Host ("PODBITO: {0}" -f $o.Adres) -ForegroundColor Green
            Start-Sleep -Seconds $PauzaSekund
        }
    }
    catch {
        $Bledy++
        Write-Host ("BLAD przy {0} (linia {1}): {2}" -f $o.Adres, $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message) -ForegroundColor Red
    }
}

# --- SPRZATANIE ---
$wb.Close($true)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host ""
if ($Podglad) {
    Write-Host "Gotowe - sprawdz otwarte szkice w Outlooku (nic nie wyslano)." -ForegroundColor Cyan
} else {
    Write-Host ("Gotowe. Podbito: {0}, bledy: {1}. Statusy w maile.xlsx." -f $Wyslane, $Bledy) -ForegroundColor Cyan
}
pause
