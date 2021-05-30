# ipscanner

## Opis ogólny
Program służy do zbierania parametrów komputerów w sieci lokalnej z zadanego zakresu adresów IP. Program nie wykorzystuje agentów instalowanych na komputerach.

Zakres adresów IP wyznaczany jest na podstawie aktualnego adresu IP komputera.

Program sprawdza czy pod danym adresem IP jest włączony komputer próbując połaczyć się na porcie TCP 445 (NetBIOS). Timeout nawiązania połączenia 200 ms. Jeśli nie uda się nawiązać połączenia, to dany adres IP nie pojawi się na liście wyjściowej.

Dane wyjściowe są prezentowanie w tabeli Powershella. Tabela umożliwia łatwe filtrowanie i sortowanie danych.

## Funkcje
1. **Name** - nazwa komputera. Pobierana komendą Powershell *Get-WmiObject win32_computersystem*.
1. **Last Logged User** - nazwa użytkownika, który ostatnio logował się na komputerze. Określana na podstawie pliku na pulpicie "Skrót do moje konto".
1. **User Name** - nazwa aktualnie zalogowanego użytkownika. Pobierana komendą Powershell *Get-WmiObject win32_computersystem*.
1. **OS** - nazwa systemu operacyjnego komputera. Pobierana komendą Powershell *Get-WmiObject Win32_OperatingSystem*.
1. **OS Version** - nazwa i wersja systemu operacyjnego komputera pobierana komendą Powershell *Get-WmiObject Win32_OperatingSystem*. Pobierany jest numer build i na podstawie tabeli zamienianu=y na numer wydania.
1. **System Family** - ... Pobierana komendą Powershell *Get-WmiObject win32_computersystem*.
1. **Logical Processors** - liczba procesorów logicznych komputera. Pobierana komendą Powershell *Get-WmiObject win32_computersystem*.
1. **Memory** - ilość pamięci fizycznej komputera (dostepnej dla systemu operacyjnego). Pobierana komendą Powershell *Get-WmiObject win32_computersystem*.
1. **Time Source** - źródło czasu (serwer NTP). Jeśli nie jest to nazwa serwera, oznacza to, że usługo w32time nie działa poprawnie. Pobierane skryptem Powershell *w32tm /query /status | Select-String -Pattern "^Source:.\*"
1. **Chrome Version** - wersja zainstalowanej przegladarki Chrome. Pobierane przez wyczytanie wersji z pliku chrome.exe. 
1. **PrintConfig.dll** - obecność pliku PrinConfig.dll w katalogu *C:\Windows\System32\spool\drivers\x64\3*. Brak tego pliku powoduje błąd przy próbie druku z aplikacji Universal Windows Platform.
1. **Office**  - wersja oprogramowania MS Office. Pobierane skryptem Powershell *Get-WmiObject win32_product | where{$_.Name -like "Microsoft Office Standard\*" -or $_.Name -like "Microsoft Office Professional\*"} | select Name,Version*.

### Dodawanie nowych funkcji
Nowe funkcjonalności można dodawać tworząc nowe funkcje klasy *Func* i dodając wpis do listy *properties*.

## Wykryte problemy
1. W przypadku kilku kart sieciowych (także wirtualnych) zakres IP może być wyznaczony na bazie nieodpowiedniej karty.

## Pakowanie do exe
1. W środowisku wirtualnym pipenv zainstalować pyinstaller (jesli nie jest juz zainstalowany).
`pipenv install pyinstaller`
1. Uruchomić shell w środowisku wirtualnym.
`pipenv shell`
1. Skompilować plik.
`pyinstaller -w --onefile ipscan.py`
1. Plik znajduje sie w podkatalogu *dist*
