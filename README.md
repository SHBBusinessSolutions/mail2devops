# Mail2DevOps – Outlook Add-in für Azure DevOps

Hängt E-Mails, ganze Verläufe oder **nur markierte Ausschnitte** als Diskussion (Kommentare) an Azure DevOps Work Items an – oder erstellt direkt ein neues PBI / User Story / Task / Bug.

Läuft in **neuem Outlook für Windows, klassischem Outlook (Windows/Mac) und Outlook im Web** (Office Web Add-in, kein VSTO).

## Funktionen

| | |
|---|---|
| **Ziel** | Bestehendes Item (Suche nach ID oder Titel, Filter *zuletzt geändert / mir zugewiesen / von mir erstellt*) oder neues Item (Typen werden je Projekt/Prozess geladen: PBI, User Story, Task, Bug, Feature, Epic …) |
| **Mehrere Organisationen** | Beliebig viele Org-URLs mit eigenem PAT, Standard-Projekt und API-Version. Admin kann Orgs zentral in `config.json` vorgeben. |
| **Ganze E-Mail** | Ein Kommentar mit Kopfzeile (Von/An/Cc/Datum/Betreff) und bereinigtem Inhalt |
| **Verlauf** | Erkennt die einzelnen Nachrichten im zitierten Verlauf (Outlook Desktop, Outlook Web/neues Outlook, Gmail, Thunderbird, Apple Mail, „-----Ursprüngliche Nachricht-----“) und postet **jede Nachricht als eigenen Kommentar** – chronologisch, einzeln abwählbar → echte Diskussion in DevOps |
| **Auswahl** ⭐ | Text in der Vorschau markieren → **＋ Übernehmen**. Mehrere Ausschnitte, farbig markiert, jeweils mit Absender und Datum zitiert. Beim Verfassen: Markierung direkt aus dem Editor. |
| **Bilder** | Eingebettete Bilder (cid:) werden als DevOps-Anhang hochgeladen und im Kommentar angezeigt |
| **Anhänge** | E-Mail-Anhänge und optional die Original-E-Mail als `.eml` ans Work Item hängen |
| **Neues Item** | Titel aus Betreff (ohne AW:/RE:/WG:), Tags, übergeordnetes Item, „mir zuweisen“, Inhalt in Diskussion und/oder Beschreibung (bei Bugs: Repro Steps) |
| **Komfort** | Work-Item-ID im Betreff (`#4711`, `AB#4711`) wird erkannt · Verläufe merken sich ihr Item · Task Pane anheftbar (wechselt mit der Auswahl) · Dark Mode |

## Bestehende Lösungen (Recherche 18.11.2026)

- **„Email to Work Item“ (Innova-Apps, Microsoft Marketplace)** – erstellt *neue* Work Items aus E-Mails und hängt die `.eml` an. Kein Anhängen als Kommentar an bestehende Items, keine Verlaufszerlegung, keine Ausschnitte.
- **Zapier / Zoho Flow** – Automatisierungen (Regel-basiert), kein interaktives Add-in; Zapier vor allem DevOps → Outlook.
- Das frühere *Team Companion* / TFS-Outlook-Integrationen sind eingestellt.

→ Die Kombination *bestehendes Item + Diskussion + Verlauf aufteilen + markierte Ausschnitte* gibt es nicht – daher dieses Add-in.

## Schnellstart (Entwicklung)

Voraussetzung: Node.js 18+.

```bash
npm install
npm run certs      # einmalig: vertrauenswürdiges localhost-Zertifikat installieren
npm start          # https://localhost:3000
```

**Demo ohne Outlook:** `https://localhost:3000/src/taskpane.html` im Browser öffnen – simulierte E-Mail und simuliertes DevOps, die ganze Oberfläche ist testbar. Eine echte Organisation kann dort ebenfalls hinzugefügt werden.

### In Outlook laden (Sideload)

- **Outlook im Web / neues Outlook:** <https://aka.ms/olksideload> öffnen → *Meine Add-Ins* → *Benutzerdefinierte Add-Ins* → *Aus Datei hinzufügen* → `manifest.xml`.
- **Klassisches Outlook (Windows):** `npm run sideload` (startet Server + Outlook) oder wie oben über Outlook im Web hinzufügen – es erscheint dann auch im Desktop.

Danach in einer E-Mail: Menüband **„An DevOps senden“** (bzw. *Apps* → Mail2DevOps).

## Personal Access Token

In Azure DevOps: *User settings → Personal access tokens → New Token* (Link direkt im Einstellungsformular).
Scopes: **Work Items: Read & Write** und **Project and Team: Read**.

PATs werden standardmässig nur lokal auf dem Gerät gespeichert (`localStorage` des Add-ins). Optional „geräteübergreifend“ → in den Roaming Settings des Postfachs (nur dieses Add-in kann sie lesen).

## Hosting über GitHub Pages (empfohlen, kostenlos)

Der Workflow [.github/workflows/pages.yml](.github/workflows/pages.yml) baut das Add-in bei jedem Push auf `main`, veröffentlicht es und stellt das fertige Manifest bereit.

1. Neues GitHub-Repository anlegen und das Projekt pushen. Für Pages im Free-Plan muss es **public** sein. Das ist unkritisch: Im Code stehen keine Geheimnisse, PATs bleiben bei den Benutzern. Org-URLs in `config.json` wären öffentlich sichtbar.
2. Im Repo **Settings → Pages → Source: GitHub Actions** wählen.
3. Nach dem ersten Lauf (Tab *Actions*) ist das Add-in erreichbar unter `https://<user>.github.io/<repo>/`.
4. Das passende Manifest herunterladen: `https://<user>.github.io/<repo>/manifest.xml`. Dann sideloaden oder im Admin Center verteilen.
5. Optional prüfen: `npm run check-host -- https://<user>.github.io/<repo>`

Updates: einfach pushen. Die URL bleibt gleich, das Manifest muss nicht neu verteilt werden (ausser es ändert sich selbst, z.B. die Version).

## An alle verteilen

1. Das Manifest von `https://<user>.github.io/<repo>/manifest.xml` herunterladen.
2. **Microsoft 365 Admin Center → Einstellungen → Integrierte Apps → Benutzerdefinierte Apps hochladen** und für Benutzer/Gruppen freigeben. Alternativ lädt jeder es selbst wie oben.
3. Optional `config.json` befüllen (Vorlage: `config.example.json`), damit Organisationen vorkonfiguriert sind und Benutzer nur noch ihren PAT eintragen. Nach dem Push gilt die Vorgabe für alle:
   ```json
   { "organizations": [ { "name": "Contoso", "url": "https://dev.azure.com/contoso", "defaultProject": "Webshop" } ] }
   ```

## Architektur

```
manifest.xml          Add-in-Manifest (Lesen + Verfassen, anheftbar)
config.json           optionale Admin-Vorgaben
src/taskpane.html     Oberfläche
src/app.js            UI-Logik, Ablauf beim Senden
src/parse.js          HTML bereinigen, Verlauf zerlegen, Kommentar-HTML bauen
src/ado.js            Azure DevOps REST (Projekte, Typen, WIQL, Work Items, Kommentare, Anhänge)
src/mail.js           Office.js-Zugriff (Body, Anhänge, .eml, Editor-Auswahl)
src/config.js         Speicherung Organisationen / PAT / Einstellungen
src/demo.js           Simulierte E-Mail + DevOps für den Browser-Test
scripts/serve.mjs     HTTPS-Dev-Server ohne Abhängigkeiten
```

Keine Build-Pipeline, kein Backend: Das Add-in spricht direkt mit der Azure DevOps REST-API (CORS mit PAT wird von Azure DevOps Services unterstützt). Daten verlassen den Rechner nur in Richtung der eigenen DevOps-Organisation.

Verwendete APIs: `wit/wiql`, `wit/workitems` (GET/POST/PATCH), `wit/workItems/{id}/comments` (7.1-preview.4), `wit/attachments`, `wit/workitemtypecategories`.

## Hinweise & Grenzen

- **Markieren beim Lesen:** Office.js erlaubt im Lesemodus keinen Zugriff auf die Markierung im Outlook-Lesebereich. Deshalb wird in der Vorschau des Add-ins markiert (komfortabler: mehrere Ausschnitte, Highlighting). Beim *Verfassen* wird die Editor-Markierung direkt übernommen.
- **Verlauf:** Die Zerlegung nutzt den zitierten Verlauf im aktuellen Mail-Body. Wurde der Verlauf beim Antworten gelöscht, ist er auch hier nicht vorhanden.
- **Azure DevOps Server (on-prem):** API-Version 7.0/6.0 wählbar. Der Server muss Browser-Zugriffe (CORS) vom Add-in-Host erlauben.
- **Requirement Sets:** Basis Mailbox 1.5; Anhänge/Bilder ab 1.8; `.eml` ab 1.14 – nicht unterstützte Optionen werden automatisch ausgeblendet/deaktiviert.
- **Ausbauidee:** Entra-ID-SSO (Nested App Authentication) statt PAT – erfordert eine App-Registrierung im Tenant.
