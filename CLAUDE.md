# M365Manager

WPF-Desktop-App (.NET 10, MVVM via CommunityToolkit.Mvvm) zur Administration von Microsoft 365.
Löst eine Sammlung gewachsener PowerShell-Skripte ab, die unter `docs/legacy-powershell/` liegen
und weiterhin die fachliche Referenz sind — wer ein Feature portiert, liest zuerst das alte Skript.

## Projekte

| Projekt | Inhalt |
|---|---|
| `src/M365Manager` | WPF-UI: Views, ViewModels, Themes, Converters |
| `src/M365Manager.Core` | Fachlogik, Exchange/Graph/PowerShell-Zugriff, Settings, Notifications |
| `src/M365Manager.Data` | EF Core, SQL-Cache, Repositories, `Scripts/*.sql` |
| `src/ImportScripts` | Geplante Importe, die den SQL-Cache füllen |

## Build & Run

```powershell
dotnet build                                   # ganze Solution
pwsh scripts/Publish-Portable.ps1              # portable Zip auf den Desktop
pwsh scripts/Restore-Modules.ps1               # einmalig: gebündelte PS-Module holen
pwsh scripts/Test-SettingsLoad.ps1             # Prüfungen für SettingsService.Load
pwsh scripts/Test-Contrast.ps1                 # WCAG-Kontrast beider Farbpaletten
pwsh scripts/Test-AlertBox.ps1                 # AlertBox respektiert gebundene Visibility
```

**Es gibt kein Testprojekt.** Absicherung läuft über `dotnet build` plus manuellen Smoke-Test der
geänderten Screens. Bei XAML-Umbauten reicht ein grüner Build nicht — Binding-Fehler zeigen sich
erst zur Laufzeit, also die betroffene Seite wirklich öffnen und durchklicken.

Einzige Ausnahme ist `scripts/Test-SettingsLoad.ps1`: sechs Prüfungen gegen `SettingsService.Load`,
weil eine Regression dort teuer ist (ein unlesbarer Wert hatte früher alle Sektionen verworfen, und
der nächste `Save` schrieb die Defaults über die noch intakte Datei). Nach jeder Änderung an
`SettingsService`, `AppSettings` oder der Serialisierung laufen lassen. Das Skript arbeitet auf der
echten `settings.json` — es sichert sie vorher und stellt sie im `finally` wieder her.

`scripts/Test-Contrast.ps1` misst die WCAG-Kontraste beider Paletten gegen die Paare, die die UI
tatsächlich rendert. Nach jeder Änderung an `Themes/Colors.*.xaml` laufen lassen — beim ersten Lauf
sind vier Paare durchgefallen, darunter eines, dessen Verhältnis ein Code-Kommentar behauptet hatte,
ohne dass es je gemessen worden war.

Hinweis: `dotnet build --no-incremental` auf `src/M365Manager/M365Manager.csproj` schlägt mit
Fehlern in einem `_wpftmp`-Projekt fehl (WPF-Markup-Compile-Eigenheit, verliert die
ProjectReference). Kein echter Fehler — stattdessen `dotnet clean && dotnet build` auf der Solution.

## Zielgruppe & Hauptaufgaben

**Wer:** Der Nutzer und sein IT-Team arbeiten selbst als M365-Admins mit dem Tool. Jede Sitzung ist
**ticketgetrieben** — Änderungen und Neueinrichtungen kommen als ServiceNow-Tickets herein und
werden einzeln abgearbeitet. Deshalb hat fast jedes Formular ein Pflichtfeld „Ticket / task number",
und jede Aktion landet mit dieser Nummer im Audit-Log (`dbo.LogEntries`).

**Konsequenz für die UI:** Pro Sitzung wird *eine* Aufgabe erledigt, nicht explorativ gestöbert. Der
Weg von „Ticket gelesen" zu „Aktion ausgeführt" soll kurz und eindeutig sein. Ein Screen, der fünf
Formulare gleichzeitig zeigt, kostet in diesem Ablauf mehr, als er nützt.

### Priorisierung (bestätigt 2026-09-23)

**Häufig — täglich, Kern des Tagesgeschäfts:**

- **Groups** — Verteilerlisten, Security-, M365- und dynamische Gruppen
- **Shared Mailboxes** — anlegen, Owner wechseln, umbenennen
- **Teams** — Teams anlegen und verwalten
- **Rooms & Resources** — Teams Rooms und Ressourcen-Postfächer
- **New Hire** — Mitarbeiter für Teams-Telefonie aktivieren, Rufnummer zuweisen
- **Teams Policies** — Meeting-Policy-Gruppen pro Benutzer

**Selten:**

- **Trace Application** — Gast-Accounts für Geschäftspartner (Entra-Gasteinladungen)

**Unterstützend, keine Tagesaufgabe:**

- **Dashboard** — Verbindungsstatus
- **Logs** — Audit-Trail lesen und exportieren
- **Settings** — Einrichtung

Bei UI-Arbeit gilt: häufige Aufgaben prominent und mit wenigen Klicks; Seltenes darf in Menüs oder
Einstellungen wandern. Bei Zielkonflikten zwischen „selten" und „häufig" gewinnt „häufig".

## Oberfläche & Design-System

**Die GUI ist durchgehend Englisch**, auch wenn die Konversation deutsch läuft.

### Nur Tokens, keine Literale

Farb-, Abstands-, Radius- und Schriftwerte kommen **ausschließlich** aus:

| Datei | Inhalt |
|---|---|
| `Themes/Tokens.xaml` | Abstände, Radien, Schriftgrößen, Schatten, Layoutmaße — themenunabhängig |
| `Themes/Colors.Light.xaml` / `Colors.Dark.xaml` | Gleiche Schlüssel, andere Werte |
| `Themes/Styles.xaml` | Alle Basis-Komponenten, referenziert nur Tokens |

Ein hardcodiertes `#RRGGBB` oder ein frei gewählter Margin in einer View ist ein Fehler. Vor dem
Umbau: die Abstände sind **4 / 8 / 12 / 16 / 24 / 32 / 48**, dazwischen gibt es nichts. Für die
üblichen Formularabstände die benannten Thickness-Tokens nutzen (`GapLabel`, `GapHint`, `GapField`,
`GapSection`, `GapInline`).

**Farben immer über `DynamicResource`**, nie `StaticResource`: `ThemeService` tauscht das
Farb-Dictionary zur Laufzeit, eine `StaticResource` behält den beim Laden aufgelösten Wert.

### Rot hat zwei Bedeutungen — und drei Tokens

Das ist die zentrale Regel dieser Oberfläche:

- `BrandBrush` — **Fläche** in Markenrot unter weißer Schrift. Genau *ein* Primary-Button pro
  Screen. Sonst nichts.
- `BrandTextBrush` — Markenrot als **Text, Symbol, Akzentbalken oder Rahmen**. Im Dark Mode ein
  anderer Wert als `BrandBrush`, weil auf dunklem Grund kein einziges Rot beides kann.
- `DangerBrush` / `DangerFillBrush` — dieselbe Trennung für destruktive Aktionen.
- `OnBrandBrush` — Beschriftung auf einer Marken- oder Danger-Fläche. In beiden Paletten weiß;
  **nicht** `TextInverseBrush`, das im Dark Mode fast schwarz ist.

### Destruktive Aktionen

Im Ruhezustand **nicht rot**: `DangerButton` ist ein neutraler Rahmen mit Warndreieck und einem
Label aus Verb und Objekt („Delete group", nicht „Remove"). Rot erscheint erst auf dem
Bestätigen-Button des Modals, das der Button öffnet.

**Jede** destruktive Aktion geht über `IDialogService.ConfirmDestructive` — keine Inline-Checkbox,
kein `MessageBox`, und niemals ohne Rückfrage. Der Dialog nennt das Objekt beim Namen und sagt, was
mitgeht.

### Seitenaufbau

- Mehrere Aufgaben auf einer Seite → `TabStrip`, nicht gestapelt. Sichtbar ist genau ein Formular.
- Leitmuster der Arbeitsseiten: **Übersicht → Zeile wählen → verwalten**. Die Auswahl ist Komfort;
  die Identität muss weiterhin eintippbar bleiben, sonst sind Objekte außerhalb des SQL-Caches
  unerreichbar.
- Genau **eine** primäre Aktion pro Screen.
- Jede Liste braucht einen Leerzustand, jede Aktion eine Rückmeldung über `AlertBox` mit
  **explizit gesetzter** Severity — nie aus dem Text erraten.
- Formularbreite über `ContentWidth`, Tabellen dürfen die volle Breite nutzen.

### Basis-Komponenten

`Controls/AlertBox` (4 Stufen) · `Controls/ConfirmDialog` + `Services/DialogService` ·
`Controls/PowerShellConsole` (nur noch einmal in der Shell, nicht pro Seite) ·
`Services/ThemeService` · `Services/NavigationService` + `ITabbedPage` für Sprünge auf Seite + Tab.

Buttons: `PrimaryButton` · `SecondaryButton` · `GhostButton` · `DangerButton` ·
`DangerConfirmButton` (nur im Dialog), je mit `…Small`-Variante.

### AlertBox und Sichtbarkeit

`AlertBox` klappt sich bei leerem Text selbst zu — über einen **Style-Trigger**, nicht per
Zuweisung im Code. Der Unterschied ist nicht kosmetisch: eine Zuweisung schreibt einen lokalen
Wert, und der schlägt in WPF jede Bindung, die der Aufrufer im XAML gesetzt hat. Genau so ist eine
`Visibility="{Binding HasConflict...}"`-Bindung still gestorben und eine Warnung über das Entfernen
von Gruppen stand dauerhaft im Bild.

Wer an `AlertBox` oder an einer View arbeitet, die deren `Visibility` bindet:
`pwsh scripts/Test-AlertBox.ps1`. Der Test prüft beide Reihenfolgen — XAML bindet zuerst und setzt
`Text` danach, und nur diese Reihenfolge deckt den Fehler auf.

### Barrierefreiheit

WCAG AA ist Vorgabe, nicht Kür — `scripts/Test-Contrast.ps1` prüft es.

**Jedes** Eingabefeld, jede ComboBox und jedes Grid braucht `AutomationProperties.Name`; ohne den
liest ein Screenreader nichts Brauchbares vor. Der Name ist üblicherweise derselbe Text wie im
`FieldLabel` darüber.

**Tastatur:**

| Taste | Wirkung |
|---|---|
| `Ctrl+1` … `Ctrl+0` | springt zur 1. bis 10. Seite der Navigationsleiste |
| ``Ctrl+` `` | öffnet/schließt die PowerShell-Leiste |
| `Alt+<Buchstabe>` | Access Key eines Buttons (Unterstrich im `Content`) |
| `Enter` | löst die Hauptaktion aus (`IsDefault`, Suchfelder per `KeyBinding`) |
| `Esc` | bricht den Bestätigungsdialog ab |

Access Keys werden pro Seite vergeben — zwei gleichzeitig sichtbare Buttons dürfen sich keinen
Buchstaben teilen. Buttons auf verschiedenen Tabs dürfen, da immer nur ein Tab existiert. Die
Button-Templates setzen `RecognizesAccessKey`, der Unterstrich im `Content` genügt also.

Der Fokusring ist neutral (Textfarbe), nicht farbig: ein rotes Ringlein würde entweder mit der
Marke kollidieren oder als Warnung gelesen — beides ist nicht, was Fokus bedeutet.

## Konventionen

- **Benachrichtigungs-E-Mails sind immer Englisch**, auch wenn die Konversation deutsch läuft —
  die Empfänger sitzen in mehreren Ländern. Der Wortlaut stammt aus den Legacy-`.oft`-Vorlagen;
  `docs/mail-templates/README.md` ist das Register, welche Vorlage wo portiert ist.
- **Versand läuft über `INotificationMailService`**, nie direkt über Graph `/me/sendMail` — das
  Admin-Konto hat kein Postfach. Seit 2026-08-31 SMTP-Relay.
- **Fehlermeldungen an den Nutzer** über `ErrorText.Describe(ex)`, nicht `ex.Message` — EF versteckt
  die echte Ursache in der InnerException.
- **Neue Felder in `AppSettings`** müssen zusätzlich in `SettingsService.Save` ergänzt werden; die
  Methode baut das Objekt von Hand neu (wegen Passwort-Verschlüsselung) und verliert sonst das Feld.
- **Zugriffsgruppen nie aus einem Namen konstruieren** — erst über die tatsächlichen Berechtigungen
  suchen, dann auf abgeleitete Namen zurückfallen. Historisch von Hand umbenannte Gruppen brechen
  jede reine Namensableitung.
