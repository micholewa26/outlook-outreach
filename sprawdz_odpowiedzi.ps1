# ============================================================
#  SPRAWDZANIE ODPOWIEDZI - zespolu
#  Od wiersza $OdWiersza w dol: jesli firma odpisala na maila,
#  wpisuje "TAK <data>" w kolumnie 3 (czy odpisal). NIC nie wysyla.
#  Uruchomienie: SPRAWDZ_ODPOWIEDZI.bat
# ============================================================

$ErrorActionPreference = 'Stop'

$Folder    = Split-Path -Parent $MyInvocation.MyCommand.Path
$PlikExcel = Join-Path $Folder 'maile.xlsx'
$Nadawca = $env:OUTREACH_MAILBOX
if (-not $Nadawca) { throw 'Ustaw zmienna OUTREACH_MAILBOX na adres skrzynki wspoldzielonej (patrz README).' }
$KolOdpisal = 3
$KolWyslano = 4
$OdWiersza  = 2      # sprawdzaj od tego wiersza excela w dol
$DniWstecz  = 45      # jak gleboko skanowac skrzynke odbiorcza


if (-not (Test-Path $PlikExcel)) { Write-Host "BLAD: brak maile.xlsx" -ForegroundColor Red; pause; exit 1 }

Write-Host "Wczytuje maile.xlsx..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($PlikExcel)
$ws = $wb.Worksheets.Item(1)

$Lista = @()
$WszystkieAdresy = @{}
$row = 2
while ($ws.Cells.Item($row, 1).Value2) {
    $adres = ([string]$ws.Cells.Item($row, 1).Value2).Trim().ToLower()
    $WszystkieAdresy[$adres] = $row
    if ($row -ge $OdWiersza) {
        $Lista += [PSCustomObject]@{
            Wiersz  = $row
            Adres   = $adres
            Odpisal = [string]$ws.Cells.Item($row, $KolOdpisal).Value2
        }
    }
    $row++
}
Write-Host ("Wierszy do sprawdzenia (od {0}): {1}" -f $OdWiersza, $Lista.Count)

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

# korzen skrzynki zespolu (zeby skanowac tez podfoldery, np. "Wspolprace w rozmowie")
$root = $null
if (-not $root) {
    try { $root = $ns.GetSharedDefaultFolder($recip, 6).Parent } catch { }
}
if (-not $root) { Write-Host "BLAD: nie znalazlem skrzynki zespolu w Outlooku." -ForegroundColor Red; $wb.Close($false); $excel.Quit(); pause; exit 1 }
Write-Host ("Skrzynka: {0}" -f $root.Name)

$PomijaneFoldery = @('Elementy wys*','Wersje robocze','Skrzynka nadawcza','Wiadomo*mieci','Junk*','Spam*','Sent*','Drafts','Outbox','Kalendarz','Kontakty','Zadania','Notatki','Dziennik','Calendar','Contacts','Tasks','Notes','Journal','Kana*y informacyjne*','RSS*','Historia konwersacji','Conversation History')
function Czy-Pomijac($nazwa) {
    foreach ($p in $PomijaneFoldery) { if ($nazwa -ilike $p) { return $true } }
    return $false
}
$FolderyDoSkanu = New-Object System.Collections.ArrayList
function Zbierz-Foldery($folder, $glebokosc) {
    if ($glebokosc -gt 4) { return }
    foreach ($sub in $folder.Folders) {
        if (Czy-Pomijac $sub.Name) { continue }
        [void]$FolderyDoSkanu.Add($sub)
        Zbierz-Foldery $sub ($glebokosc+1)
    }
}
Zbierz-Foldery $root 0

function SmtpNadawcy($it) {
    try {
        if ($it.SenderEmailType -eq 'SMTP') { return ([string]$it.SenderEmailAddress).ToLower().Trim() }
        $ex = $it.Sender.GetExchangeUser()
        if ($ex) { return ([string]$ex.PrimarySmtpAddress).ToLower().Trim() }
    } catch { }
    return ''
}

$Cutoff = (Get-Date).AddDays(-$DniWstecz)
Write-Host ("Skanuje {0} folderow skrzynki zespolu (maile od {1})..." -f $FolderyDoSkanu.Count, $Cutoff.ToString('yyyy-MM-dd'))
$Odpowiedzi = @{}
foreach ($fold in $FolderyDoSkanu) {
    $n = 0
    try {
        $items = $fold.Items
        $items.Sort('[ReceivedTime]', $true)
        foreach ($it in $items) {
            try {
                if ($it.Class -ne 43) { continue }
                if ($it.ReceivedTime -lt $Cutoff) { break }
                $s = SmtpNadawcy $it
                if ($s -and (-not $Odpowiedzi.ContainsKey($s) -or $it.ReceivedTime -gt $Odpowiedzi[$s])) { $Odpowiedzi[$s] = $it.ReceivedTime }
                $n++
            } catch { }
        }
    } catch { }
    if ($n -gt 0) { Write-Host ("  {0}: {1} maili" -f $fold.Name, $n) }
}
Write-Host ("Maile od {0} roznych nadawcow w skrzynce." -f $Odpowiedzi.Count)

# diagnostyka: pokaz nadawcow, ktorzy NIE pasuja do zadnego adresu z excela
$Niedopasowani = @()
foreach ($s in $Odpowiedzi.Keys) {
    if (-not $WszystkieAdresy.ContainsKey($s)) { $Niedopasowani += $s }
}
if ($Niedopasowani.Count -gt 0) {
    Write-Host ""
    Write-Host ("Nadawcy bez dopasowania w excelu ({0}) - moze odpisali z innego adresu niz ten na liscie:" -f $Niedopasowani.Count) -ForegroundColor Yellow
    foreach ($s in ($Niedopasowani | Sort-Object)) {
        $dom = ($s -split '@')[-1]
        $hint = ''
        foreach ($a in $WszystkieAdresy.Keys) {
            if (($a -split '@')[-1] -eq $dom) { $hint = (' <- ta sama domena co {0} (wiersz {1})' -f $a, $WszystkieAdresy[$a]); break }
        }
        Write-Host ("  ? {0}{1}" -f $s, $hint)
    }
    Write-Host "Tych ze wskazana domena mozesz oznaczyc recznie TAK w excelu po sprawdzeniu tresci maila." -ForegroundColor Yellow
}

$Oznaczone = 0
foreach ($o in $Lista) {
    if ($o.Odpisal) { continue }
    if ($Odpowiedzi.ContainsKey($o.Adres)) {
        $ws.Cells.Item($o.Wiersz, $KolOdpisal).Value2 = ('TAK ' + (Get-Date $Odpowiedzi[$o.Adres] -Format 'yyyy-MM-dd'))
        Write-Host ("TAK: {0} (odpisal {1})" -f $o.Adres, (Get-Date $Odpowiedzi[$o.Adres] -Format 'yyyy-MM-dd')) -ForegroundColor Green
        $Oznaczone++
    }
}

$wb.Close($true)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host ""
Write-Host ("Gotowe. Oznaczono TAK: {0} z {1} sprawdzonych." -f $Oznaczone, $Lista.Count) -ForegroundColor Cyan
pause
