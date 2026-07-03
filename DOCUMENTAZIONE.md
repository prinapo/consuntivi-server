# Consuntivi Server — Documentazione

## 1. Architettura Generale

Sistema per la generazione automatica di consuntivi XLSX a partire da file SVS in SharePoint.

| Componente | Tecnologia | Ruolo |
|---|---|---|
| Backend (server) | Java 21, JDK HttpServer, Maven | Elabora file SVS, genera XLSX, carica su SharePoint |
| Frontend (SPFx) | TypeScript, SPFx v1.23.0, Heft | Pulsante "Genera Consuntivo" nella barra comandi SharePoint |
| Proxy | nginx + Let's Encrypt | HTTPS, proxy `/api/` → `127.0.0.1:8080` |
| Systemd | `consuntivi-server.service` | Gestione ciclo di vita del JAR |

URL backend: `https://development.sostienilsostegno.com/api/consolidaconsuntivi`  
Health: `https://development.sostienilsostegno.com/api/health` → `{"status":"ok","version":"1.1.0"}`

---

## 2. Struttura Directory

### Server (`/home/ubuntu/consuntivi-server/`)

```
consuntivi-server/
├── .env                          # Azure credentials (non in git)
├── pom.xml                       # Maven: dipendenze + shade plugin
├── DOCUMENTAZIONE.md             # Questo file
├── src/
│   ├── main/
│   │   ├── java/com/server/
│   │   │   ├── ConsolidaHandler.java     # 1187 righe — handler principale
│   │   │   ├── Main.java                # 33 righe — entry point HTTP
│   │   │   └── TestDownloadHandler.java # 130 righe — test endpoint
│   │   └── resources/
│   │       └── template-consuntivo.xlsx # Backup template (obsoleto)
└── target/
    └── consuntivi-server-1.1.0.jar      # JAR eseguibile (shaded)
```

### SPFx (`/home/ubuntu/consuntivi-heft/`)

```
consuntivi-heft/
├── src/
│   ├── index.ts                                    # Entry point
│   └── extensions/generaConsuntivo/
│       ├── GeneraConsuntivoCommandSet.ts            # 264 righe — main logic
│       └── GeneraConsuntivoCommandSet.manifest.json # Manifest SPFx
├── config/
│   ├── package-solution.json       # v1.2.2.0
│   ├── config.json                 # Bundle config
│   └── serve.json                  # Dev server config
├── sharepoint/assets/
│   ├── elements.xml                # CustomAction SharePoint
│   └── clientsideinstance.xml      # Instance registration
└── package.json                    # SPFx 1.23.0, Heft 1.2.17
```

---

## 3. Flusso di Elaborazione

```
┌─────────────────────────────────────────────────────────────────┐
│ SPFx (browser)                                                  │
│ 1. onListViewUpdated() → controlla regex /\/(.+)\/06\.Ordini$/i │
│ 2. Utente clicca "Genera Consuntivo"                            │
│ 3. confirm("Cancellare i precedenti?") → deleteOld=true/false   │
│ 4. fetch POST a /api/consolidaconsuntivi                        │
│ 5. Legge NDJSON streaming via ReadableStream.getReader()        │
└──────────────────────────┬──────────────────────────────────────┘
                           │ POST {"templateUrl":"...", "deleteOld":bool}
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│ Server (ConsolidaHandler.java)                                   │
│ 1. parseUrl(templateUrl) → host, sitePath, filePath              │
│ 2. resolveSiteId(host, site) → Graph API sites/{host}:{site}    │
│ 3. resolveProjectBasePath → cerca in:                           │
│    - {projectName} diretto                                       │
│    - Commesse/{year}/{projectName}                               │
│    - Documenti Condivisi/{projectName}                           │
│ 4. listSvsFiles → elenca SVS*.xlsx in 06.Ordini/RDA             │
│ 5. Per ogni file: readTabella2 → cerca Tabella2 nel foglio      │
│    - ❌ Tabella2 non trovata → skip                              │
│    - ✅ filtro Quantità ≠ 0 → valide                             │
│    - ⚠️ 0 valide (ma Tabella2 presente)                          │
│ 6. Calcola Mese (MM/yyyy) da Data (dd/MM/yyyy o serial)         │
│ 7. downloadTemplate() → scarica da ICTmanagement (35536 bytes)  │
│ 8. Scrive tutte le righe nel foglio, espande TabellaRDA         │
│ 9. deleteOldConsuntivi() → se richiesto, cancella precedenti     │
│ 10. uploadFile() → PUT /content diretto su SharePoint            │
│ 11. writeDone() → NDJSON evento "done"                          │
└─────────────────────────────────────────────────────────────────┘
```

### Eventi NDJSON

Il server streama eventi in formato `application/x-ndjson` (una riga JSON per evento):

| Evento | Campi | Descrizione |
|---|---|---|
| `log` | `message` | Messaggio di stato visibile all'utente |
| `progress` | `current, total, file, rows` | Avanzamento elaborazione SVS |
| `done` | `url, fileName, totalRows` | Completato con URL del file caricato |
| `error` | `message` | Errore (interrompe il flusso) |

---

## 4. ConsolidaHandler.java — Specifiche

### Costanti

```java
SERVER_VERSION = "1.1.0"
COLUMN_NAMES = ["RDA", "Data", "Cod", "Prodotto", "UM", "Colonna1",
                "PZ", "Quantità", "Prezzo", "Importo",
                "Fornitore", "Qt. Min. V.", "Note", "Mese"]
```

### Metodi Principali

| Metodo | Visibilità | Descrizione |
|---|---|---|
| `handle(HttpExchange)` | `public` | Entry point: gestisce richiesta POST/GET, stream NDJSON |
| `parseUrl(String url)` | `static` | Estrae hostname, sitePath, filePath da URL SharePoint |
| `resolveSiteId(String host, String site)` | `private` | Risolve sito SharePoint via Graph API |
| `resolveProjectBasePath(String siteId, String projectName)` | `private` | Trova cartella progetto in 3 pattern |
| `listSvsFiles(String siteId, String rdaPath)` | `private` | Elenca file SVS*.xlsx nella cartella RDA |
| `downloadTemplate()` | `private` | Scarica template da ICTmanagement via GET diretto |
| `readTabella2(String siteId, String itemId, String itemName)` | `private` | Scarica file SVS, cerca Tabella2, estrae righe |
| `uploadFile(String siteId, String parentFolder, String fileName, byte[] data)` | `private` | PUT diretto, gestisce 423 Locked |
| `deleteFileWithBypass(String siteId, String parentId, String fileName)` | `private` | DELETE con Prefer: bypass-shared-lock |
| `deleteOldConsuntivi(String siteId, String parentFolder, String projectName)` | `private` | Cancella vecchi consuntivi che matchano pattern |
| `writeEvent(OutputStream, String type, String message)` | `private` | Streama evento NDJSON |
| `writeProgress(OutputStream, int current, int total, String fileName, int rows)` | `private` | Streama progresso elaborazione |
| `writeDone(OutputStream, String url, String fileName, int totalRows)` | `private` | Streama evento completamento |

### Metodi Obsoleti (non usati nel flusso attuale)

| Metodo | Descrizione |
|---|---|
| `resolveSiteMetadata(String siteId)` | Recupera ContentTypeId e site GUID da Graph API |
| `patchTemplate(SiteMetadata meta)` | Patches customXml con site metadata |
| `modifyWithPoi(...)` | Vecchio metodo di generazione XLSX (sostituito da template) |
| `cleanContentTypesOrphans(byte[])` | Rimuove Override orfani da [Content_Types].xml |
| `stripCustomXml(byte[])` | Rimuove tutti i customXml dal ZIP |

### Gestione Upload

- `PUT /sites/{siteId}/drive/items/{parentId}:/{fileName}:/content` diretto
- **Delete prima del PUT** per evitare bug growth-hint di SharePoint (Apache POI Bug #65706)
- **423 Locked** → DELETE con header `Prefer: bypass-shared-lock` → retry PUT
- Nome file: `{YYYYMMDD_HHMM}_Consuntivo_{projectName}.xlsx` (CET/Europe/Rome)

---

## 5. Specifiche SPFx

### GeneraConsuntivoCommandSet.ts (v1.2.2.0)

**File**: `src/extensions/generaConsuntivo/GeneraConsuntivoCommandSet.ts` (264 righe)

**Dipendenze SPFx**:
- `@microsoft/sp-listview-extensibility` 1.23.0
- `@microsoft/sp-page-context` 1.23.0

**Componenti**:
- `BaseListViewCommandSet` — estensione barra comandi
- `GENERA_CONSUNTIVO` — pulsante con icona SVG blu (#0078D4)

**Visibilità**:
- Solo in cartelle che matchano regex `/\/([^/]+)\/06\.Ordini$/i`
- Lettura da `id` o `RootFolder` parametro URL
- Estratto `projectName` dal path

**Overlay DOM**:
- Non bloccante, posizionato in alto a destra
- Testo selezionabile, scrollabile
- Pulsante "Chiudi"
- Streaming NDJSON via `response.body.getReader()` + `TextDecoder` a buffer

**Flusso client**:
1. `confirm('Vuoi cancellare i consuntivi precedenti di questo progetto?')`
2. `fetch(FUNCTION_URL, { body: JSON.stringify({ templateUrl, deleteOld }) })`
3. Legge NDJSON in loop:
   - `log` → append testo
   - `progress` → update title + append ✅
   - `done` → mostra URL file
   - `error` → mostra errore

---

## 6. Configurazioni

### pom.xml

```xml
<groupId>com.server</groupId>
<artifactId>consuntivi-server</artifactId>
<version>1.1.0</version>
<properties>
    <java.version>21</java.version>
    <maven.compiler.source>21</maven.compiler.source>
</properties>
```

Dipendenze:
| Libreria | Versione | Uso |
|---|---|---|
| `microsoft-graph` | 5.80.0 | SharePoint API (drive, items, sites) |
| `azure-identity` | 1.10.4 | ClientSecretCredential per OAuth2 |
| `gson` | 2.10.1 | JSON parsing (Graph responses, NDJSON) |
| `poi-ooxml` | 5.2.5 | Lettura/scrittura XLSX, pivot tables |
| `dotenv-java` | 3.0.0 | Caricamento .env in sviluppo |

Build: `maven-shade-plugin` 3.5.1 → JAR eseguibile con `com.server.Main` come main class.

### systemd

```ini
[Unit]
Description=Consuntivi Server - Java HTTP backend for consuntivi generation

[Service]
Type=simple
User=ubuntu
WorkingDirectory=/home/ubuntu/consuntivi-server
EnvironmentFile=/home/ubuntu/consuntivi-server/.env
ExecStart=/usr/bin/java --add-opens java.base/java.time=ALL-UNNAMED
           -jar /home/ubuntu/consuntivi-server/target/consuntivi-server-1.1.0.jar
Restart=on-failure
RestartSec=5
```

Comandi:
```bash
sudo systemctl restart consuntivi-server
sudo journalctl -u consuntivi-server -f
```

### nginx

```nginx
location /consuntivi/ {
    alias /var/www/consuntivi/;
}
location /api/ {
    proxy_pass http://127.0.0.1:8080;
    proxy_buffering off;           # Necessario per streaming NDJSON
    proxy_read_timeout 300s;
}
location / {
    proxy_pass http://localhost:9000;
}
```

### Azure AD (Service Principal)

| Parametro | Valore |
|---|---|
| `AZURE_CLIENT_ID` | `9f424713-d07d-4807-8ebd-0dc5164b8e61` |
| `AZURE_TENANT_ID` | `fec3df64-6d43-46e7-81e6-8ed42be543bd` |
| Scope | `https://graph.microsoft.com/.default` |
| Permessi | `Sites.ReadWrite.All` (app-level) |

### SharePoint Site IDs

| Sito | Site ID |
|---|---|
| CommesseSVS2026 | `svssrl.sharepoint.com,205255dd-24ea-46c6-a79e-63b7587072c0,d5ab7167-c317-4bff-841e-a3c8512f3e70` |
| ICTmanagement (template) | `svssrl.sharepoint.com,96c00cdc-6492-44ab-94d9-2f7a25aabfd9,5820c7dd-ac32-4237-9d28-4b52de70fbd3` |

---

## 7. Schema Dati

### TabellaRDA (Excel Table su Foglio1)

Range dinamico: `A4:N{lastRow}` (header a riga 4, dati da riga 5)

| Colonna | Nome | Tipo | Note |
|---|---|---|---|
| A | RDA | String | Numero RDA documento |
| B | Data | Date/Serial | `dd/MM/yyyy` o serial Excel |
| C | Cod | String | Codice prodotto |
| D | Prodotto | String | Descrizione prodotto |
| E | UM | String | Unità di misura |
| F | Colonna1 | String | Informazione aggiuntiva |
| G | PZ | Number | Pezzi per collo |
| H | Quantità | Number | Quantità (filtrata ≠ 0) |
| I | Prezzo | Number | Prezzo unitario |
| J | Importo | Number | Importo totale |
| K | Fornitore | String | Nome fornitore |
| L | Qt. Min. V. | Number | Quantità minima vendibile |
| M | Note | String | Note opzionali |
| N | Mese | String | Derivato da Data: `MM/yyyy` |

La data viene processata nel loop (righe 257-276 di ConsolidaHandler.java):
- Se stringa: parse `dd/MM/yyyy`
- Se serial Double > 40000: Excel serial date (1900-01-01 + serial - 2)
- Altri casi: Mese vuoto

### File SVS (sorgente)

I file SVS contengono una tabella Excel chiamata **Tabella2** (cercata per nome, non per posizione). Il server:
1. Cerca `XSSFTable` con nome "Tabella2"
2. Legge righe dalla prima riga dati (headerRow + 1) all'ultima (endRef.getRow())
3. Filtra righe con `Quantità != 0`
4. Limita a 14 colonne

### Template ICTmanagement

- File: `TemplateConsuntiviRDA/Template pivot riassunto ordini.xlsx`
- 5 fogli: Foglio1 (dati) + 4 fogli pivot
- Tabella: TabellaRDA (`xl/tables/table1.xml`)
- Stile tabella: `TableStyleMedium5` (righe alterne)
- Pivot cache: referenzia TabellaRDA per nome (`worksheetSource name="TabellaRDA"`)
- 4 pivot table sui fogli 2-5

---

## 8. Versioni

| Componente | Versione | Note |
|---|---|---|
| Server JAR | `1.1.0` | `pom.xml`, `SERVER_VERSION`, `Main.java` health |
| SPFx Extension | `1.2.2.0` | `GeneraConsuntivoCommandSet.ts` + `package-solution.json` |
| Template ICTmanagement | — | Gestito dall'utente in Excel |

---

## 9. Build e Deploy

### Server

```bash
cd /home/ubuntu/consuntivi-server
mvn package -DskipTests
sudo systemctl restart consuntivi-server
sudo journalctl -u consuntivi-server -f
```

### SPFx

```bash
cd /home/ubuntu/consuntivi-heft
npm run build -- --clean --production
npm run package-solution
sudo cp sharepoint/solution/consuntivi-heft.sppkg /var/www/consuntivi/
```

Caricare `.sppkg` su App Catalog SharePoint → `https://development.sostienilsostegno.com/consuntivi/consuntivi-heft.sppkg`

### Debug

File generato salvato in: `/var/www/consuntivi/last-generated.xlsx`  
URL debug: `https://development.sostienilsostegno.com/consuntivi/last-generated.xlsx`

---

## 10. Upload File

### PUT diretto (metodo attuale)

1. DELETE preventivo del file con stesso nome (404 = non esiste, ignorato)
2. `PUT /sites/{siteId}/drive/items/{parentId}:/{fileName}:/content`
3. Header: `Authorization: Bearer {token}`, `Content-Type: application/octet-stream`
4. Se 201 → successo, parse risposta JSON in `DriveItem`
5. Se 423 Locked → DELETE con `Prefer: bypass-shared-lock` → retry PUT

### Perché non createUploadSession

Evitato `createUploadSession` per bug 409 `nameAlreadyExists` causato da sessioni stale.

---

## 11. Cancellazione Vecchi Consuntivi

Se `deleteOld=true`, prima di caricare il nuovo file:
1. Lista tutti i file in `06.Ordini` che matchano:
   - `Consuntivo_{projectName}.xlsx` (vecchio formato, senza timestamp)
   - `*_Consuntivo_{projectName}.xlsx` (nuovo formato con timestamp)
2. Per ogni match: `DELETE /sites/{siteId}/drive/items/{parentId}:/{file}` con `Prefer: bypass-shared-lock`

Il nuovo file non viene cancellato perché la cancellazione avviene **prima** dell'upload.

---

## 12. Note sul Fuso Orario

Il server utilizza `Europe/Rome` (CET/CEST) per:
- Nome file: `YYYYMMDD_HHMM_Consuntivo_{project}.xlsx`
- Riga info data: `dd/MM/yyyy`
- Riga info timestamp: `Ultimo aggiornamento: dd/MM/yyyy HH:mm`

Il server fisico è in UTC (VPS OVH), ma `ZonedDateTime.now(ZoneId.of("Europe/Rome"))` garantisce il fuso corretto.
