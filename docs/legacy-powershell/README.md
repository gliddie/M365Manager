# Legacy PowerShell Scripts

Bitte lege hier **alle bestehenden PowerShell-Skripte** der alten GUI ab
(einfach die `.ps1`-Dateien in diesen Ordner kopieren — gerne auch die
zugehörigen Modul-/Konfig-Dateien).

Ich (Claude) analysiere sie dann und leite daraus ab:

- Welche M365-/Azure-Admin-Funktionen die App abdecken muss
- Wie das aktuelle Logging in Textdateien funktioniert (damit wir es 1:1
  auf die SQL-Datenbank übertragen können)
- Welche Cmdlets/Module verwendet werden (Microsoft Graph, ExchangeOnline,
  MSOnline/AzureAD etc.)

## Ablage

```
docs/legacy-powershell/
├── README.md          <- diese Datei
└── (hier deine *.ps1 Skripte ablegen)
```

Wenn die Skripte drin sind, sag mir einfach Bescheid ("Skripte sind drin").
