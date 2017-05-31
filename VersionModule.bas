Attribute VB_Name = "VersionModule"

' PFEP 2.0
' 2015-08-24
' nowa wersja PFEPA z dodatkiem beton, ktory dzialac bedzie w przeciwna strone
' zasada dzialania bedzie bardzo podobna do makra dynamic POP
' to jest bedzie mozliwosc wybierania pol jakie chce sie zobaczyc na raporcie koncowym
' =======================================================================================================


' PFEB 2.1
' 2015-08-25
' plan for every beton
' preinput logika wlasciwie ready
' moge pomalu niewinnie zaczac pisac main run
' =======================================================================================================


' PFEB 2.2
' 2015-08-26
' plan for every beton
' logika juz dziala
' teraz iu i bedzie cacy
' plus dodatowo scenariusze okreslone przez Paule ktore odpowiednio ostawiaja arkusz config
' =======================================================================================================

 
' PFEB 2.3 1st of Sept 2015
' plan for every beton
' teraz trzeba dopisac interfejsy
' dwa dodatkowe pola na ekranie ms7p5100
' plus poprawic zle zaciagany std pack na ekranie ms9pv400
' arkusze output dodawane na koncu
' =======================================================================================================

' PFEB 2.3 4th of Sept 2015
' =======================================================================================================
' scenario scaffold on place!
' add some forms
' =======================================================================================================


' PFEB 2.4 2015,09,07
' =======================================================================================================
' forms under development
'

'
' to be implemented:
' define at runtime new scenarios
' buttons already created
' waiting for natty & simple code
' for dynamic assigning in free slots (overwrite also)
' =======================================================================================================

' PFEB 2.5 2015-09-08
' =======================================================================================================
' dynamic assigning scenarios developed (9 slots with 2 const)
' maybe new form for adding to have possibility to add to place
' =======================================================================================================


' PFEB 2.6 2015-10-02
' =======================================================================================================
' poprawione scenario - nie przesuwa sie juz
' osea scenario modernizacja
' android layout
' filter clearing
' preinput to input poprawka
' =======================================================================================================


' PFEB 2.7 2015-10-02
' =======================================================================================================
' dodanie MYSPTOG0 - dziala jak zk7ptogl
' =======================================================================================================


' PFEB 2.71 2015-10-02
' =======================================================================================================
' drobna kosmetyka
' =======================================================================================================

' PFEB 2.73 2015-10-02
' =======================================================================================================
' poprawienie logiki pre inputu duns wchodzil w nieistniejace pola
' =======================================================================================================


' PFEB 2.8 2015-10-28
' =======================================================================================================
' zera w duns i pn
' =======================================================================================================


' PFEB 2.81 2015-11-06
' =======================================================================================================
' dodatkowy rozszerzony zapis dla lsity mgo screenu do otwarcia - screen openers
' =======================================================================================================


' PFEB 2.82 2016-04-04
' =======================================================================================================
' output adjsutment layout by Paulina
' dodanie ikonek do adjustmentu graficznego
' =======================================================================================================


' PFEB 2.83 2016-08-02
' =======================================================================================================
' output adjsutment layout OSEA by Paulina
' dodanie pola COUNTRY CODE do scenariusza Osea
' wdrozenie adjustmentu szerokosci kolumn, czcionek, dodanie legendy
' =======================================================================================================


' PFEB 2.84 2016-08-03
' =======================================================================================================
' output adjsutment layout FMA by Paulina
' wdrozenie adjustmentu szerokosci kolumn, czcionek, dodanie legendy
' =======================================================================================================


' PFEB 2.85 2016-08-03
' =======================================================================================================
' output adjsutment layout Component by Paulina
' wdrozenie adjustmentu szerokosci kolumn, czcionek, dodanie legendy
' =======================================================================================================


' PFEB 2.86 2016-08-04
' =======================================================================================================
' output adjsutment OSEA mark mistakes by Paulina
' dodanie warunkow oznaczenia bledow dla OSEA
' =======================================================================================================


' PFEB 2.87 2016-08-04
' =======================================================================================================
' output adjsutment FMA mark mistakes by Paulina
' dodanie warunkow oznaczenia bledow dla FMA
' =======================================================================================================



' PFEB 2.88 2016-08-04
' =======================================================================================================
' output adjsutment COMPONENT mark mistakes by Paulina
' dodanie warunkow oznaczenia bledow dla COMPONENT
' dodanie blokad adjustmentow dla niewlasciwych scenariuszy
' =======================================================================================================



' PFEB 2.89 2016-08-05
' =======================================================================================================
' ZK7PCONT - dodanie numeru kontraktu , DESC, zmiana formatu dat dla kontraktow - by Mateusz
' kolorystyczne oznaczenie waznosci kontraktu
' =======================================================================================================



' PFEB 2.9 2016-08-08
' =======================================================================================================
' all collumns scenario adjustment added (layout + mistakes)
' dodane instrukcji - by Paulina
' dodanie przycisku adjustmentu (All adjustments) by Mateusz
' =======================================================================================================



' PFEB 2.91 2016-08-12
' =======================================================================================================
' vlookup for osea scenario added - by Mateusz
' Osea scenario -  modified mistakes marking for KB plant & COM CODE checking(petla_zlozenia_pol/ cztery_petle_zlozenia)
' Final tests And Release
' =======================================================================================================



' PFEB 2.92 2016-08-23
' =======================================================================================================
' osea scenario freeze top row added - by Mateusz
' com code rule modified for Osea - if blank  - by Paulina
' wyroznienie "no FU" kolor - by Paulina
' =======================================================================================================



' PFEB 2.93 2016-08-24
' =======================================================================================================
' osea scenario - legend modification (Planning Action / NOK for Osea) -by Paulina
' osea scenario - change of color adjustment (TT - cream to red) - by Paulina
' pre-input - hide option filter enabled
' input - buyer code added
' =======================================================================================================



' PFEB 2.94 2016-08-24
' =======================================================================================================
' input - buyer code added ZK7PCONT - by Mateusz
' Final tests And Release
' =======================================================================================================


' PFEB 2.95 2016-08-24
' =======================================================================================================
' zostalo dodane 5224 linie kodu w ramach ostatniej wersji 2.95
' dodany modul eksportu kodu w ramach wydzielenia logiki makra z samego excela
' wstepna inicjalizacja logiki sterujacej kontraktami dla preinputu i jego klasy"
' PreInputHandler
' w klasie tej znajduje sie sub sprawdz_czy
' jendak okazalo sie za klasa ta jest zbyt plaska by poradzic sobie z wyrafinowanymi warunkami
' ktore egzystuja na wielu plaszczyznach (nawet zagniezdzone ify to troche za duzo)
'
' rozwiazaniem ktore wydaje sie byc najlepszym to stworzenie odpowiednich obiektow
' ktore gromadzic beda stosowne info na temat jednego pn'u
' w jedna kolekcje ktora bedzie w stanie
' w odpowiedn sposob realizowac zadania nakreslone przez end usera danej klasy
'
'
' =======================================================================================================




' PFEB 2.96 2016-08-24
' =======================================================================================================
' ostatecznie w implementacji zamieszczona zostala dla logiki tworzenia inputu z preinputu
' dwie petle, a nie tak jak wczensiej jedna
' sub sprawdz_czy stal sie obsoletem, a zamiast niego dorzuclem obiekty: ContractItem zawierajace
' wlasciwie tylko publiczne pola, kolejnym obiektem juz nieco bardziej skomplikowanym jest PnForContractItem
' ktory zawiera jako element / pole kolekcje wlasnie wczesniej wspomnianych ajtemow
'
' sama klasa pre input handler wlasciwie sie nie zmienila (poza wspomnianym rodzieleniem jednej petli na dwie plus
' dodatkowe interacje po slowniku, ktory od tej wersji jest integralnym prywatnym polem tejze klasy.
' =======================================================================================================


' PFEB 2.96.3 2016-11-04
' =======================================================================================================
' zmiana weryfikacji waznosci kontraktu rowniez od strony graficznej - update kolorystyki
' =======================================================================================================



' PFEB 2.96.31 2016-11-07
' =======================================================================================================
' scenariusz O'sea - zmiana koloru dla oznaczenia flagi w przypadku gdy zostala ona zmieniona i nie zostala jeszcze przeprocesowana systemowo
' update tresci legendy scenariusza Osea z "will be set after COM CODES" na "changed - will be OK"
' =======================================================================================================



' PFEB 2.97 2016-11-18
' =======================================================================================================
' dodnie scenariusza BN / zdefiniowanie kolumn - by Paulina
' dodnaie zaciagania dodatkowych danych z ekranu ms9pv400:
' ms9pv400_SEC_STD_PACK
' ms9pv400_SEC_CNTR
' ms9pv400_SEC_CONT_WG
' ms9pv400_PRIM_DIM
'
' dodanie przycisku BTN Scenaio - by Paulina
' =======================================================================================================



' PFEB 2.97 1 2016-11-20
' =======================================================================================================
' dodnie oznaczen bledow dla scenariusza BN  - by Paulina
' dodanie oznczen bledow dla komponenowego scenariusza COM CODEs
' =======================================================================================================



' PFEB 2.97 2 2016-11-20
' =======================================================================================================
' korekcja kolorystyczna legendy scenariusza component i FMA  - by Paulina
' oznaczenie scenariusza BTN w panelu glownym
' test and release
' =======================================================================================================

' PFEB 2.97 3 2017-01-29
' =======================================================================================================
' change of description of contract validation from "EXPIRES THIS MONTH" to "EXPIRES WITHIN 30 DAYS"  - by Paulina
' test and release
' =======================================================================================================

' PFEB 2.98 2017-02-15
' =======================================================================================================
' DUNS added as a parameter for ROUTE verification in FMA scenario  - by Paulina
' ROUTE scren opener - DUNS parameter added
' test and release
' =======================================================================================================

' PFEB 2.98 1 2017-02-21
' =======================================================================================================
' CONTAINER added to FMA scenario  - by Paulina
' modified adjustment FMA
' test and release
' =======================================================================================================


' PFEB 2.99 2017-02-22
' =======================================================================================================
' OSEA scenario cplumn comp  chedule published added  - by Paulina
' modified adjustment OSEA
' =======================================================================================================


' PFEB 3.0 2017-02-22
' =======================================================================================================
' Wgeneral import logic added and implemented (WGenerelItem added / WgeneralModule)- by Mateusz
' Lgeneral import logic added (to be implemented)
' test and release
' =======================================================================================================




' PFEB 3.01 2017-02-23
' =======================================================================================================
' new MODULE - regular expressions
' =======================================================================================================


' PFEB 3.02 2017-02-23
' =======================================================================================================
' new MODULE - regular expressions - korekty logiki
' =======================================================================================================


' PFEB 3.03 2017-02-24
' =======================================================================================================
' new MODULE - regular expressions - korekty logiki
' + mocno rozpiany CheckMod z nowa logika sprawdzajaca kolumny com code
' oraz country code w subie com_code_loop
' =======================================================================================================

' PFEB 3.04 2017-03-03
' =======================================================================================================
' default TT logic added
' test and release
' =======================================================================================================


' PFEB 3.05 2017-03-09
' =======================================================================================================
' additional logic for pre-input and switching between contracts with diff dunses
' =======================================================================================================

' PFEB 3.06 2017-03-09
' =======================================================================================================
' pa dunses recognition
' =======================================================================================================
