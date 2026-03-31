# Pratiche Immigrazione

Navigatore pratiche per l'Ufficio Immigrazione — estrazione dal gestionale, checklist, note.

## Scarica l'ultima versione

Vai su **[Releases](../../releases)** e scarica `Pratiche_Immigrazione.zip` dall'ultima build.

Decomprimi e avvia `Pratiche.exe` — niente installazioni richieste.

## Requisiti

- Windows 10/11
- Firefox installato

## Struttura file dopo la decompressione

```
Pratiche.exe        ← applicazione
geckodriver.exe     ← driver Firefox (incluso nel pacchetto)
pratiche_NOME.xlsx  ← creato automaticamente al primo salvataggio
```

## Come si usa

1. Avvia `Pratiche.exe`
2. All'apertura vedi la lista pratiche (vuota al primo avvio)
3. Clicca **🌐 Browser** → si apre Firefox sul gestionale
4. Fai il login e naviga alla pratica
5. Clicca **⚡ Estrai** → anteprima dati
6. Conferma → pratica salvata, lista aggiornata
7. Doppio click sulla riga in lista → apre scheda + attività

## Build manuale

Per compilare da soli:

```
pip install pyinstaller customtkinter selenium openpyxl
pyinstaller --onefile --windowed --name Pratiche --collect-all customtkinter --collect-all selenium pratiche.py
```
