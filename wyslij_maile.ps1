# ============================================================
#  BOT DO WYSYLANIA MAILI - zespolu
#  Czyta: mailform.docx (temat + tresc), maile.xlsx (adres, zwrot)
#  Wysyla przez klasycznego Outlooka ze skonfigurowanej skrzynki.
#  Uruchomienie:  WYSLIJ.bat  (wysylka)  lub  PODGLAD.bat  (tylko podglad, nic nie wysyla)
# ============================================================

param([switch]$Podglad, [switch]$Test)

$ErrorActionPreference = 'Stop'

# --- USTAWIENIA ---
$Folder       = Split-Path -Parent $MyInvocation.MyCommand.Path
$PlikWord     = Join-Path $Folder 'mailform.docx'
$PlikExcel    = Join-Path $Folder 'maile.xlsx'
$FolderZal    = Join-Path $Folder 'zalaczniki'
$Placeholder  = '[place holder z excela]'   # ten tekst w wordzie jest podmieniany na zwrot z excela
$Nadawca = $env:OUTREACH_MAILBOX
if (-not $Nadawca) { throw 'Ustaw zmienna OUTREACH_MAILBOX na adres skrzynki wspoldzielonej (patrz README).' }
$KolumnaStatus = 4                          # kolumna 'wyslano' w excelu
$PauzaSekund  = 3                           # przerwa miedzy mailami


if (-not (Test-Path $PlikWord))  { Write-Host "BLAD: brak pliku mailform.docx" -ForegroundColor Red; pause; exit 1 }
if (-not (Test-Path $PlikExcel)) { Write-Host "BLAD: brak pliku maile.xlsx"   -ForegroundColor Red; pause; exit 1 }

$Zalaczniki = @()
if (Test-Path $FolderZal) { $Zalaczniki = @(Get-ChildItem $FolderZal -File | ForEach-Object { $_.FullName }) }

# --- EXCEL: wczytaj liste adresatow ---
Write-Host "Wczytuje liste z maile.xlsx..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($PlikExcel)
$ws = $wb.Worksheets.Item(1)

if (-not $ws.Cells.Item(1, $KolumnaStatus).Value2) {
    $ws.Cells.Item(1, $KolumnaStatus).Value2 = 'wyslano'
}

$Lista = @()
$row = 2
while ($ws.Cells.Item($row, 1).Value2) {
    $adres   = ([string]$ws.Cells.Item($row, 1).Value2).Trim()
    $zwrot   = [string]$ws.Cells.Item($row, 2).Value2
    if ($zwrot) { $zwrot = $zwrot.Trim() } else { $zwrot = '' }
    $wyslano = [string]$ws.Cells.Item($row, $KolumnaStatus).Value2
    $Lista += [PSCustomObject]@{ Wiersz = $row; Adres = $adres; Zwrot = $zwrot; Wyslano = $wyslano }
    $row++
}

$DoWyslania = @($Lista | Where-Object { -not $_.Wyslano })
$Pominiete  = @($Lista | Where-Object { $_.Wyslano })

# tryb testowy: jeden mail na podany adres, bez zapisywania statusow w excelu
if ($Test) {
    $adresTest = (Read-Host "TRYB TESTOWY - podaj adres, na ktory wyslac pojedyncza wiadomosc").Trim()
    if (-not $adresTest) { Write-Host "Nie podano adresu."; $wb.Close($false); $excel.Quit(); pause; exit 1 }
    $DoWyslania = @([PSCustomObject]@{ Wiersz = 0; Adres = $adresTest; Zwrot = 'Szanowni Panstwo,'; Wyslano = $null })
    $Pominiete  = @()
}

# --- PODSUMOWANIE PRZED WYSYLKA ---
Write-Host ""
Write-Host "=== PODSUMOWANIE ===" -ForegroundColor Cyan
Write-Host ("Adresatow na liscie : {0}" -f $Lista.Count)
Write-Host ("Juz wyslane (pomijam): {0}" -f $Pominiete.Count)
Write-Host ("Do wyslania          : {0}" -f $DoWyslania.Count) -ForegroundColor Yellow
Write-Host ("Zalaczniki           : {0}" -f $(if ($Zalaczniki.Count) { ($Zalaczniki | Split-Path -Leaf) -join ', ' } else { 'BRAK' }))
Write-Host ("Nadawca              : {0}" -f $Nadawca)
foreach ($o in $DoWyslania) { Write-Host ("  -> {0}  ({1})" -f $o.Adres, $o.Zwrot) }
Write-Host ""

if ($DoWyslania.Count -eq 0) {
    Write-Host "Nie ma nic do wyslania." -ForegroundColor Green
    $wb.Close($true); $excel.Quit(); pause; exit 0
}

if ($Podglad) {
    Write-Host "TRYB PODGLADU - maile tylko otworza sie jako szkice, NIC nie zostanie wyslane." -ForegroundColor Yellow
} elseif ($Test) {
    Write-Host ("TRYB TESTOWY - wysylam 1 mail na {0}" -f $DoWyslania[0].Adres) -ForegroundColor Yellow
} else {
    $odp = Read-Host "Wpisz T i Enter, zeby WYSLAC maile (cokolwiek innego = anuluj)"
    if ($odp -ne 'T' -and $odp -ne 't') {
        Write-Host "Anulowano."; $wb.Close($false); $excel.Quit(); exit 0
    }
}

# --- OUTLOOK ---
# klasyczny Outlook MUSI byc uruchomiony z oknem - inaczej edytor wiadomosci nie dziala
if (-not (Get-Process OUTLOOK -ErrorAction SilentlyContinue)) {
    Write-Host "Outlook nie jest uruchomiony - odpalam go i czekam..." -ForegroundColor Yellow
    Start-Process outlook.exe
    $czekaj = 0
    while (-not (Get-Process OUTLOOK -ErrorAction SilentlyContinue) -and $czekaj -lt 60) { Start-Sleep -Seconds 2; $czekaj += 2 }
    Start-Sleep -Seconds 15   # czas na zaladowanie profilu i skrzynek
}
Write-Host "Lacze sie z Outlookiem..."
$outlook = New-Object -ComObject Outlook.Application
Write-Host "Polaczono z Outlookiem."

# skrzynka zespolu to skrzynka udostepniona (nie osobne konto) -> wysylka "w imieniu" (SentOnBehalf).
# Sprawdz, czy Outlook umie rozwiazac adres nadawcy w katalogu organizacji:
$recip = $outlook.Session.CreateRecipient($Nadawca)
$null = $recip.Resolve()
if (-not $recip.Resolved) {
    Write-Host ""
    Write-Host "BLAD: Outlook nie rozpoznaje adresu nadawcy '$Nadawca' - NIC NIE WYSYLAM." -ForegroundColor Red
    Write-Host "Sprawdz, czy klasyczny Outlook jest otwarty i zalogowany na koncie z dostepem do skrzynki." -ForegroundColor Yellow
    $wb.Close($false); $excel.Quit(); pause; exit 1
}
Write-Host ("Wysylka w imieniu: {0} (rozpoznano: {1})" -f $Nadawca, $recip.Name) -ForegroundColor Green

$Wyslane = 0
$Bledy   = 0

foreach ($o in $DoWyslania) {
    try {
        Write-Host ("Tworze wiadomosc dla {0}..." -f $o.Adres)
        $mail = $outlook.CreateItem(0)   # 0 = MailItem
        $mail.BodyFormat = 2             # 2 = HTML

        # edytor Worda w mailu -> wstaw tresc docx (jawne parametry: NIE jako zalacznik)
        $insp = $mail.GetInspector
        $wdoc = $insp.WordEditor
        if (-not $wdoc) {
            # edytor nie wstal (Outlook swiezo uruchomiony) - otworz okno wiadomosci, to go inicjalizuje
            $mail.Display()
            Start-Sleep -Milliseconds 1000
            $wdoc = $insp.WordEditor
        }
        if (-not $wdoc) { throw "Outlook nie udostepnil edytora wiadomosci - otworz klasycznego Outlooka (pelne okno) i odpal skrypt ponownie." }
        $wdoc.Content.Delete() | Out-Null    # usun ew. automatyczny podpis
        $wdoc.Range(0,0).InsertFile($PlikWord, "", $false, $false, $false) | Out-Null

        # 1. akapit = "Temat: ..." -> wyciagnij i usun z tresci
        $p1 = $wdoc.Paragraphs.Item(1).Range
        $temat = ($p1.Text -replace '^Temat:\s*','').Trim()
        $p1.Delete() | Out-Null

        # podmien placeholder na zwrot z excela
        $find = $wdoc.Content.Find
        $null = $find.Execute($Placeholder, $false, $false, $false, $false, $false, $true, 1, $false, $o.Zwrot, 2)

        $mail.Subject = $temat
        $mail.To      = $o.Adres

        # nadawca: wysylka w imieniu skrzynki zespolu
        $mail.SentOnBehalfOfName = $Nadawca

        foreach ($z in $Zalaczniki) { $mail.Attachments.Add($z) | Out-Null }

        if ($Podglad) {
            $mail.Display()
            Write-Host ("PODGLAD: {0}" -f $o.Adres) -ForegroundColor Yellow
        } else {
            # zapisz mail i zamknij niewidoczny edytor Worda - bez tego Send() rzuca
            # "Wartosc jest spoza oczekiwanego zakresu"
            $mail.Save()
            $wdoc = $null; $find = $null; $p1 = $null
            $insp.Close(0) | Out-Null   # 0 = zamknij z zapisem
            $insp = $null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            Start-Sleep -Milliseconds 300
            $mail.Send()
            if ($o.Wiersz -gt 0) {
                $ws.Cells.Item($o.Wiersz, $KolumnaStatus).Value2 = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                $wb.Save()
            }
            $Wyslane++
            Write-Host ("WYSLANO: {0}" -f $o.Adres) -ForegroundColor Green
            Start-Sleep -Seconds $PauzaSekund
        }
    }
    catch {
        $Bledy++
        Write-Host ("BLAD przy {0} (linia {1}): {2}" -f $o.Adres, $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message) -ForegroundColor Red
    }
}

# --- SPRZATANIE ---
if ($Podglad) { $wb.Close($false) } else { $wb.Close($true) }
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host ""
if ($Podglad) {
    Write-Host "Gotowe - sprawdz otwarte szkice w Outlooku (nic nie wyslano)." -ForegroundColor Cyan
} else {
    Write-Host ("Gotowe. Wyslano: {0}, bledy: {1}. Statusy zapisane w maile.xlsx." -f $Wyslane, $Bledy) -ForegroundColor Cyan
}
pause
