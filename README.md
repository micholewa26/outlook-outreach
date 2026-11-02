# Asystent korespondencji partnerskiej

Narzędzie powstało na potrzeby kontaktu koła naukowego z potencjalnymi partnerami. Przy kilkudziesięciu równoległych rozmowach trzeba pamiętać, kto dostał propozycję, kto odpisał i komu warto się przypomnieć. Skrypty łączą listę kontaktów w Excelu, szablony Worda i wspólną skrzynkę Outlooka. Pierwsza działająca wersja powstała w 3–4 godziny.

## Jak wygląda praca z narzędziem

1. Koordynator wpisuje adres i zwrot grzecznościowy do arkusza. Temat i treść propozycji edytuje w Wordzie.
2. Skrypt pomija wiersze z datą wysyłki. Dla pozostałych przygotowuje osobne wiadomości z właściwym powitaniem i załącznikami. Po wysłaniu zapisuje datę w arkuszu.
3. Koordynator sprawdza odpowiedzi. Skrypt przeszukuje foldery pocztowe z ostatnich 45 dni i oznacza kontakty po adresie nadawcy. Przy odpowiedziach z innego adresu tej samej firmy podpowiada dopasowanie do ręcznej oceny.
4. Po co najmniej trzech dniach można przygotować jedno przypomnienie dla kontaktów bez odpowiedzi. Skrypt dołącza treść i załączniki z najnowszej znalezionej wiadomości wysłanej do danego adresu. Gdy z tym samym partnerem toczy się kilka rozmów, sprawdź w podglądzie, którą wybrał.

## Uruchomienie podglądu

Potrzebujesz Windows, Windows PowerShell 5.1, desktopowych Worda i Excela oraz klasycznego Outlooka z profilem Exchange i dostępem do skrzynki współdzielonej. Wysyłka wymaga uprawnienia do nadawania z tej skrzynki. Integracja działa lokalnie przez COM.

W nowej kopii repozytorium otwórz PowerShell i uruchom poniższe polecenia. W drugim poleceniu wpisz rzeczywisty adres swojej skrzynki. Przykładowe dokumenty skopiuj tylko przy pierwszym uruchomieniu, aby zachować późniejsze zmiany w arkuszu.

```powershell
Copy-Item .\przyklady\* -Destination .
$env:OUTREACH_MAILBOX = 'skrzynka-zespolu@example.invalid'
.\PODGLAD.bat
```

Zobaczysz podsumowanie dwóch fikcyjnych kontaktów. Jeden ma już datę wysyłki, więc Outlook powinien otworzyć tylko jeden szkic, do `partner@example.invalid`, z tematem „Warsztaty robotyczne - propozycja współpracy” i powitaniem „Szanowni Państwo,”. Podgląd kończy się przed wysyłką.

## Polecenia

Uruchamiaj pliki BAT z tego samego okna PowerShell, w którym ustawiono adres skrzynki. Zamknij roboczy arkusz w Excelu przed uruchomieniem skryptu.

| Polecenie | Efekt |
| --- | --- |
| PODGLAD.bat | Otwiera spersonalizowane wiadomości do sprawdzenia. |
| TEST.bat | Wysyła jedną rzeczywistą wiadomość na adres wpisany przez operatora. |
| WYSLIJ.bat | Po potwierdzeniu wysyła wiadomości i zapisuje daty. |
| SPRAWDZ_ODPOWIEDZI.bat | Aktualizuje kolumnę odpowiedzi, także na podstawie wiadomości w podfolderach. |
| PODBIJ_PODGLAD.bat | Otwiera szkice przypomnień i aktualizuje wykryte odpowiedzi w arkuszu. |
| PODBIJ.bat | Po potwierdzeniu wysyła przypomnienia i zapisuje ich daty. |

Pliki robocze znajdują się obok skryptów. W pierwszym arkuszu `maile.xlsx` kolumny A–E oznaczają kolejno adres, powitanie, odpowiedź, datę wysyłki i datę przypomnienia. Zostawiaj statusy nowych kontaktów puste i wpisuj adresy bez pustych wierszy między nimi. Pierwszy akapit `mailform.docx` zawiera temat. W obu szablonach tekst `[place holder z excela]` jest zastępowany powitaniem. Pliki z opcjonalnego folderu `zalaczniki` są dodawane do każdej pierwszej wiadomości.
