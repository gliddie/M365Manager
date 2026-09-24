# Entra ID App Registration – Setup für M365Manager

Damit sich die Kollegen **mit ihrem eigenen Admin-Konto** (delegiert, interaktiv)
anmelden können, braucht M365Manager eine App-Registrierung in Entra ID.
Das ist eine **einmalige** Einrichtung durch einen Global- oder Application-Admin.

## Schritte

1. **Entra Admin Center** öffnen: https://entra.microsoft.com → **Identity → Applications → App registrations → New registration**
2. **Name:** `M365Manager` (frei wählbar)
3. **Supported account types:** *Accounts in this organizational directory only* (Single tenant)
4. **Redirect URI:**
   - Plattform: **Public client/native (mobile & desktop)**
   - Wert: `http://localhost`
5. **Register** klicken.

## Nach dem Anlegen

6. Auf der **Overview**-Seite notieren:
   - **Application (client) ID**  → kommt in M365Manager unter *Settings → Application (Client) ID*
   - **Directory (tenant) ID**    → kommt in M365Manager unter *Settings → Tenant ID*
7. **Authentication** öffnen → sicherstellen, dass unter *Advanced settings*
   **„Allow public client flows"** auf **Yes** steht.
8. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**:
   - Für den Start: `User.Read` (ist meist schon vorhanden)
   - Für die Teams-Funktion (Team erstellen) zusätzlich:
     - `Group.ReadWrite.All` — Team/Gruppe erstellen, Owner setzen
     - `Directory.ReadWrite.All` — Gastzugriff-Directory-Setting je Team
     - `Mail.Send` — Bestätigungsmail an den neuen Owner. **Nur nötig, wenn kein SMTP-Relay
       konfiguriert ist** (*Settings → Notification e-mail (SMTP)*). Ohne Relay verschickt die App
       die Mail als `/me/sendMail` unter dem angemeldeten Admin-Konto — das schlägt fehl, wenn
       dieses Konto keine lizenzierte Exchange-Mailbox hat. Mit Relay ist `Mail.Send` überflüssig.
9. **Für Exchange Online (Groups-Funktion):** **Add a permission** →
   Reiter **APIs my organization uses** → nach **`Office 365 Exchange Online`** suchen →
   **Delegated permissions** → **`Exchange.Manage`** auswählen → hinzufügen.
   (Damit kann die App mit deinem Browser-Login ein Exchange-Token holen — ohne Device-Code.)
9b. **Für SharePoint (Teams-Funktion, interne Teams sperren):** **Add a permission** →
    Reiter **Microsoft APIs** → **SharePoint** (falls nicht unter "APIs my organization uses"
    gelistet, dort taucht sie erst nach der ersten Nutzung auf) →
    **Delegated permissions** → **`AllSites.FullControl`** auswählen → hinzufügen.
    (Wird von `SharePointService`/PnP.PowerShell verwendet, um bei internen Teams die
    SharePoint-Site-Freigabe zu sperren. Das ist eine andere API als Microsoft Graph, daher
    ein eigener Berechtigungsname statt `Sites.FullControl.All`.)
10. **Grant admin consent for <Tenant>** klicken (empfohlen), damit die Kollegen beim
    ersten Login nicht jede Berechtigung einzeln bestätigen müssen.

> Falls `Exchange.Manage` nicht auftaucht: sag mir Bescheid, dann passen wir den Scope an.

## Wichtig zur Identität

- Die App-Registrierung ist nur der **Client** („Türöffner").
- Jeder Kollege meldet sich **interaktiv mit seinem eigenen Konto** an.
- Alle Änderungen laufen unter **seiner** Identität → Microsoft protokolliert sie auf ihn,
  und M365Manager schreibt seine UPN in die SQL-Log-Spalte `UserUpn`.
- MFA / Conditional Access funktionieren dabei normal.

## Exchange Online / SharePoint (eingebettete PowerShell)

Für die Exchange- und SharePoint-Funktionen nutzt M365Manager zusätzlich die Module
`ExchangeOnlineManagement` und `PnP.PowerShell` (eingebettete PowerShell).
`Connect-ExchangeOnline` bzw. `Connect-PnPOnline -Interactive` verwenden dieselbe
interaktive Anmeldung (App-Registrierung oben), sodass auch dort die Aktionen unter
dem Konto des jeweiligen Kollegen erfolgen.

Beide Module sind zu groß für git und werden lokal pro Maschine restauriert:

```
pwsh scripts/Restore-Modules.ps1
```

## Scheduled Import: Teams-Übersicht (`Import-TeamsToSql.ps1`)

Das Teams-Grid in der App liest ausschließlich aus SQL (`dbo.Teams`), nie live aus M365 - befüllt
wird das per Scheduled Task, analog zum bestehenden `Import-M365GroupsToSql.ps1` für
`dbo.Groups`. Dafür braucht ihr **zusätzlich zum Graph-App-Only-Zugang** einen eigenen
**Exchange-Online-App-Only-Zugang** (drei Felder kommen nur aus Exchange, nicht aus Graph:
Hidden from Address Book, Welcome Message, SharePoint Site URL):

1. Eigenes selbstsigniertes Zertifikat erzeugen (oder das vorhandene Graph-Zertifikat
   wiederverwenden, wenn ihr eine zweite App-Registrierung damit ausstatten wollt).
2. Neue App-Registrierung (oder die bestehende erweitern) mit dieser Zertifikats-Public-Key
   hochladen unter **Certificates & secrets**.
3. Die App-Registrierung braucht die Exchange-Online-Rolle **`Exchange Administrator`** oder
   **`Compliance Administrator`** zugewiesen (Entra ID → Roles and administrators), **nicht**
   über API-Permissions, sondern über eine Rollenzuweisung auf den Service Principal der App -
   das ist bei App-Only-Auth für Exchange Online so vorgesehen (kein Delegated/Application
   Permission-Konzept wie bei Graph).
4. `Import-TeamsToSql.ps1` mit `-ExoAppId`, `-ExoCertificateThumbprint`, `-Organization`
   (`<tenant>.onmicrosoft.com`) aufrufen, zusätzlich zu den Graph-Parametern (`-TenantId`,
   `-ClientId`, `-CertificateThumbprint`) und den SQL-Parametern.
5. Als Scheduled Task **einige Minuten nach** `Import-M365GroupsToSql.ps1` einplanen -
   `dbo.Teams` hat eine Fremdschlüsselbeziehung auf `dbo.Groups` und überspringt Teams, deren
   Gruppe dort noch nicht angekommen ist.
