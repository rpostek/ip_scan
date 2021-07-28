# ipscanner

## Opis ogólny
Program służy do zbierania parametrów komputerów w sieci lokalnej z zadanego zakresu adresów IP. Program nie wykorzystuje agentów instalowanych na komputerach.

Zakres adresów IP wyznaczany jest na podstawie aktualnego adresu IP komputera.

Program sprawdza czy pod danym adresem IP jest włączony komputer próbując połaczyć się na porcie TCP 445 (NetBIOS). Timeout nawiązania połączenia 200 ms. Jeśli nie uda się nawiązać połączenia, to dany adres IP nie pojawi się na liście wyjściowej.

Dane wyjściowe są prezentowanie w tabeli Powershella. Tabela umożliwia łatwe filtrowanie i sortowanie danych.

## Funkcje
1. **Name** - nazwa komputera. Pobierana komendą Powershell *Get-WmiObject win32_computersystem*.
1. **MAC Address** - adres aktywnej karty sieciowej *Get-NetNeighbor*.
1. **Last Logged User** - nazwa użytkownika, który ostatnio logował się na komputerze. Określana na podstawie pliku na pulpicie "Skrót do moje konto".
1. **User Name** - nazwa aktualnie zalogowanego użytkownika. Pobierana komendą Powershell *Get-WmiObject win32_computersystem*.
1. **OS** - nazwa systemu operacyjnego komputera. Pobierana komendą Powershell *Get-WmiObject Win32_OperatingSystem*.
1. **OS Version** - nazwa i wersja systemu operacyjnego komputera pobierana komendą Powershell *Get-WmiObject Win32_OperatingSystem*. Pobierany jest numer build i na podstawie tabeli zamienianu=y na numer wydania.
1. **Manufacturer** - Nazwa producenta komputera. Pobierana komendą Powershell *Get-WmiObject win32_computersystem*.
1. **System Family** - ... Pobierana komendą Powershell *Get-WmiObject win32_computersystem*.
1. **Model** - model komputera. Pobierana komendą Powershell *Get-WmiObject win32_computersystem*.
1. **BIOS** - wersja BIOSu.
1. **BIOS date** - data wydania BIOSu.
2. **Processor** - model procesora.
3. **Logical Processors** - liczba procesorów logicznych komputera. Pobierana komendą Powershell *Get-WmiObject win32_computersystem*.
4. **Memory** - ilość pamięci fizycznej komputera (dostepnej dla systemu operacyjnego). Pobierana komendą Powershell *Get-WmiObject win32_computersystem*.
5. **Time Source** - źródło czasu (serwer NTP). Jeśli nie jest to nazwa serwera, oznacza to, że usługo w32time nie działa poprawnie. Pobierane skryptem Powershell *w32tm /query /status | Select-String -Pattern "^Source:.\*"
6. **Chrome Version** - wersja zainstalowanej przegladarki Chrome. Pobierane przez wyczytanie wersji z pliku chrome.exe. 
7. **PrintConfig.dll** - obecność pliku PrinConfig.dll w katalogu *C:\Windows\System32\spool\drivers\x64\3*. Brak tego pliku powoduje błąd przy próbie druku z aplikacji Universal Windows Platform.
8. **Office** - wersja oprogramowania MS Office. Pobierane skryptem Powershell *Get-WmiObject win32_product | where{$_.Name -like "Microsoft Office Standard\*" -or $_.Name -like "Microsoft Office Professional\*"} | select Name,Version*.
9. **Monitor YearOfManufacture** - rok produkcji monitora. Pobierane skryptem Powershell *gwmi -Namespace root\\wmi -Class wmiMonitorID ...*.
10. **Monitor Name** - nazwa moniotora. Pobierane skryptem Powershell *gwmi -Namespace root\\wmi -Class wmiMonitorID ...*.
11. **Monitor SN** - numer seryjny monitora. Pobierane skryptem Powershell *gwmi -Namespace root\\wmi -Class wmiMonitorID ...*.

Jeśli zbierane są co najmniej dane: nazwa komputera, ostatni zalogowany użytkownik, MAC, to dane sa zapisywane w lokalnej bazie danych - do wykorzystania przez moduł WakOnLan.
### Dodawanie nowych funkcji
Nowe funkcjonalności można dodawać tworząc nowe funkcje klasy *Func* i dodając wpis do listy *properties*.

## Wykryte problemy
1. W przypadku kilku kart sieciowych (także wirtualnych) zakres IP może być wyznaczony na bazie nieodpowiedniej karty.
2. Część parametrów jest niedostępna dla komputera, z którego aplikacja jest uruchamiana. Rozwiązaniem jest uruchamianie z podniesionymi uporawnieniami.

## Pakowanie do exe
1. W środowisku wirtualnym pipenv zainstalować pyinstaller (jesli nie jest juz zainstalowany).
`pipenv install pyinstaller`
1. Uruchomić shell w środowisku wirtualnym.
`pipenv shell`
1. Skompilować plik.
`pyinstaller -w --onefile ipscan.py`
1. Plik znajduje sie w podkatalogu *dist*
