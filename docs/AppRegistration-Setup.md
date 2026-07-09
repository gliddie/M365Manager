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
   - Später ergänzen wir die Rechte, die die einzelnen Admin-Funktionen brauchen
     (z. B. `User.ReadWrite.All`, `Group.ReadWrite.All`, `Directory.ReadWrite.All`).
9. Optional, aber empfohlen: **Grant admin consent for <Tenant>** klicken, damit die
   Kollegen beim ersten Login nicht jede Berechtigung einzeln bestätigen müssen.

## Wichtig zur Identität

- Die App-Registrierung ist nur der **Client** („Türöffner").
- Jeder Kollege meldet sich **interaktiv mit seinem eigenen Konto** an.
- Alle Änderungen laufen unter **seiner** Identität → Microsoft protokolliert sie auf ihn,
  und M365Manager schreibt seine UPN in die SQL-Log-Spalte `UserUpn`.
- MFA / Conditional Access funktionieren dabei normal.

## Exchange Online (später)

Für die Exchange-Funktionen nutzt M365Manager zusätzlich das Modul
`ExchangeOnlineManagement` (eingebettete PowerShell). `Connect-ExchangeOnline`
verwendet dieselbe interaktive Anmeldung, sodass auch dort die Aktionen unter
dem Konto des jeweiligen Kollegen erfolgen.
