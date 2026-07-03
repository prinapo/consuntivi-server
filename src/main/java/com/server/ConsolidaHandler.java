package com.server;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URI;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.OffsetDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import java.io.StringWriter;

import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.requests.DriveItemCollectionPage;
import com.microsoft.graph.requests.GraphServiceClient;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonDeserializer;
import com.google.gson.JsonObject;

import io.github.cdimascio.dotenv.Dotenv;
import com.azure.identity.ClientSecretCredential;
import com.azure.identity.ClientSecretCredentialBuilder;
import com.microsoft.graph.authentication.TokenCredentialAuthProvider;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.http.GraphServiceException;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;

import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpHandler;

class ConsolidaHandler implements HttpHandler {

    private static final Logger LOG = Logger.getLogger(ConsolidaHandler.class.getName());

    private static final String SERVER_VERSION = "1.1.0";

    private static final Gson GSON = new GsonBuilder()
        .registerTypeAdapter(OffsetDateTime.class,
            (JsonDeserializer<OffsetDateTime>) (json, type, ctx) ->
                OffsetDateTime.parse(json.getAsString()))
        .create();

    private static final Set<String> LIBRARY_NAMES = new HashSet<>(Arrays.asList(
        "Documenti condivisi", "Shared Documents", "Documents"
    ));

    private static final List<String> COLUMN_NAMES = List.of(
        "RDA", "Data", "Cod", "Prodotto", "UM", "Colonna1",
        "PZ", "Quantità", "Prezzo", "Importo",
        "Fornitore", "Qt. Min. V.", "Note", "Mese"
    );

    private final GraphServiceClient graphClient;
    private final TokenCredentialAuthProvider authProvider;

    ConsolidaHandler() {
        var result = buildGraphClient();
        this.graphClient = result.client;
        this.authProvider = result.auth;
    }

    @Override
    public void handle(HttpExchange exchange) {
        try {
            String method = exchange.getRequestMethod().toUpperCase();

            if ("OPTIONS".equals(method)) {
                corsResponse(exchange, 204, "");
                return;
            }

            if (!"GET".equals(method) && !"POST".equals(method)) {
                corsResponse(exchange, 405, "{\"error\":\"Method not allowed\"}");
                return;
            }

            LOG.info("Avvio ConsolidaConsuntivi v" + SERVER_VERSION);

            String templateUrl = null;
            URI uri = exchange.getRequestURI();
            String query = uri.getQuery();
            if (query != null) {
                for (String param : query.split("&")) {
                    String[] kv = param.split("=", 2);
                    if (kv.length == 2 && "templateUrl".equals(kv[0])) {
                        templateUrl = java.net.URLDecoder.decode(kv[1], StandardCharsets.UTF_8);
                    }
                }
            }

            boolean deleteOld = false;

            if ("POST".equals(method)) {
                String body = new String(exchange.getRequestBody().readAllBytes(), StandardCharsets.UTF_8);
                if (body != null && !body.isBlank()) {
                    try {
                        JsonObject json = new Gson().fromJson(body, JsonObject.class);
                        if (json != null && json.has("templateUrl")) {
                            templateUrl = json.get("templateUrl").getAsString();
                        }
                        if (json != null && json.has("deleteOld")) {
                            deleteOld = json.get("deleteOld").getAsBoolean();
                        }
                    } catch (Exception e) {
                        // ignore, fallback to query param
                    }
                }
            }

            if (templateUrl == null || templateUrl.isBlank()) {
                corsResponse(exchange, 400, "{\"error\":\"Parametro templateUrl obbligatorio\"}");
                return;
            }

            LOG.info("Template URL: " + templateUrl);

            // Set up NDJSON streaming headers
            exchange.getResponseHeaders().add("Access-Control-Allow-Origin", "*");
            exchange.getResponseHeaders().add("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
            exchange.getResponseHeaders().add("Access-Control-Allow-Headers", "Content-Type");
            exchange.getResponseHeaders().add("Content-Type", "application/x-ndjson");
            exchange.sendResponseHeaders(200, 0);
            OutputStream os = exchange.getResponseBody();

            try {
                SharepointUrl parsed = parseUrl(templateUrl);
                if (parsed == null) {
                    writeEvent(os, "error", "URL SharePoint non valido");
                    return;
                }

                LOG.info("host=" + parsed.hostname + " site=" + parsed.sitePath + " file=" + parsed.filePath);

                writeEvent(os, "log", "Risolvo sito SharePoint...");
                String siteId = resolveSiteId(parsed.hostname, parsed.sitePath);
                LOG.info("Site ID: " + siteId);

                String projectName = parsed.filePath.split("/")[0];
                writeEvent(os, "log", "Cerco cartella progetto " + projectName + "...");
                String resolvedBase = resolveProjectBasePath(siteId, projectName);
                if (resolvedBase == null) {
                    writeEvent(os, "error", "Cartella progetto non trovata: " + projectName);
                    return;
                }

                String rdaPath = resolvedBase + "/06.Ordini/RDA";
                LOG.info("RDA path: " + rdaPath);

                writeEvent(os, "log", "Cerco file SVS in RDA...");
                List<DriveItem> svsFiles = listSvsFiles(siteId, rdaPath);
                LOG.info("Trovati " + svsFiles.size() + " file SVS");

                if (svsFiles.isEmpty()) {
                    writeEvent(os, "error", "Nessun file SVS trovato in RDA");
                    return;
                }

                writeEvent(os, "log", "Trovati " + svsFiles.size() + " file SVS. Inizio elaborazione...");

                int qtyIdx = -1;
                for (int i = 0; i < COLUMN_NAMES.size(); i++) {
                    if ("Quantità".equalsIgnoreCase(COLUMN_NAMES.get(i).trim())) {
                        qtyIdx = i;
                        break;
                    }
                }
                final int qIdx = qtyIdx;

                List<List<Object>> allRows = new ArrayList<>();
                int fileIndex = 0;
                for (DriveItem item : svsFiles) {
                    fileIndex++;
                    List<List<Object>> rows = readTabella2(siteId, item.id, item.name);
                    if (rows == null) {
                        writeEvent(os, "log", "\u274C " + fileIndex + "/" + svsFiles.size() + ": " + item.name + " - Tabella2 non trovata");
                        continue;
                    }
                    int rawCount = rows.size();

                    List<List<Object>> validRows = rows.stream()
                        .filter(row -> {
                            if (qIdx < 0 || qIdx >= row.size()) return true;
                            Object v = row.get(qIdx);
                            if (v == null) return false;
                            if (v instanceof Number n) return n.doubleValue() != 0;
                            if (v instanceof String s) {
                                String trimmed = s.trim();
                                return !trimmed.isEmpty() && !trimmed.equals("0") && !trimmed.equals("0.0");
                            }
                            return true;
                        })
                        .collect(Collectors.toList());

                    int validCount = validRows.size();
                    allRows.addAll(validRows);

                    if (validCount == 0 && rawCount > 0) {
                        LOG.warning("[DEBUG] file=" + item.name + " prime 3 righe col[" + qIdx + "]:");
                        for (int d = 0; d < Math.min(3, rows.size()); d++) {
                            List<Object> r = rows.get(d);
                            Object v = qIdx >= 0 && qIdx < r.size() ? r.get(qIdx) : "N/A";
                            LOG.warning("  row=" + d + " val=" + v + " type=" + (v != null ? v.getClass().getName() : "null"));
                        }
                    }

                    String icon = validCount > 0 ? "✅" : "⚠️";
                    writeEvent(os, "log", icon + " " + fileIndex + "/" + svsFiles.size() + ": " + item.name + " (" + rawCount + " raw \u2192 " + validCount + " valide)");
                }

                writeEvent(os, "log", "Totale " + allRows.size() + " righe con Quantità > 0. Generazione consuntivo...");

                if (allRows.isEmpty()) {
                    return;
                }

                DateTimeFormatter dateFmt = DateTimeFormatter.ofPattern("dd/MM/yyyy");
                DateTimeFormatter meseFmt = DateTimeFormatter.ofPattern("MM/yyyy");
                for (List<Object> row : allRows) {
                    String mese = "";
                    Object dataVal = row.size() > 1 ? row.get(1) : null;
                    if (dataVal != null) {
                        try {
                            if (dataVal instanceof String s && !s.isEmpty()) {
                                LocalDate d = LocalDate.parse(s.trim(), dateFmt);
                                mese = d.format(meseFmt);
                            } else if (dataVal instanceof Double serial && serial > 40000) {
                                LocalDate d = LocalDate.of(1900, 1, 1).plusDays((long) serial.doubleValue() - 2);
                                mese = d.format(meseFmt);
                            }
                        } catch (Exception e) {
                            LOG.fine("Mese parse error: " + dataVal);
                        }
                    }
                    row.add(mese);
                }

                String timestamp = ZonedDateTime.now(ZoneId.of("Europe/Rome")).format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm"));
                String parentFolder = resolvedBase + "/06.Ordini";
                String newName = ZonedDateTime.now(ZoneId.of("Europe/Rome")).format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmm"))
                    + "_Consuntivo_" + projectName + ".xlsx";

                writeEvent(os, "log", "Download template da ICTmanagement...");
                byte[] templateData = downloadTemplate();
                writeEvent(os, "log", "Template scaricato (" + templateData.length + " bytes). Scrittura dati...");

                XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(templateData));
                XSSFSheet sheet = wb.getSheetAt(0);

                // Trova l'ultima riga con almeno una cella non vuota (scansione dal basso)
                int lastDataRow = -1;
                for (int i = sheet.getLastRowNum(); i >= 0; i--) {
                    Row r = sheet.getRow(i);
                    if (r != null) {
                        for (int c = 0; c < 14; c++) {
                            Cell cell = r.getCell(c);
                            if (cell != null && cell.getCellType() != CellType.BLANK) {
                                lastDataRow = i;
                                break;
                            }
                        }
                        if (lastDataRow >= 0) break;
                    }
                }
                LOG.info("last data row found at " + (lastDataRow + 1));

                // Scrivi tutte le righe reali
                int rowIdx = lastDataRow + 1;
                for (List<Object> dataRow : allRows) {
                    Row excelRow = sheet.createRow(rowIdx);
                    for (int c = 0; c < Math.min(14, dataRow.size()); c++) {
                        Object val = dataRow.get(c);
                        Cell cell = excelRow.createCell(c);
                        if (val instanceof Number n) cell.setCellValue(n.doubleValue());
                        else cell.setCellValue(val != null ? val.toString() : "");
                    }
                    rowIdx++;
                }
                int totalWritten = allRows.size();
                LOG.info("scritte " + totalWritten + " righe da riga " + (lastDataRow + 2));

                // Espandi il range della TabellaRDA per includere tutte le nuove righe
                List<XSSFTable> tables = sheet.getTables();
                if (!tables.isEmpty()) {
                    XSSFTable table = tables.get(0);
                    String newRef = "A4:N" + (lastDataRow + totalWritten);
                    table.getCTTable().setRef(newRef);
                    LOG.info("expanded TabellaRDA ref to " + newRef);
                }

                // Aggiorna righe info (1-3) con progetto, data e timestamp
                Row infoRow1 = sheet.getRow(0);
                if (infoRow1 == null) infoRow1 = sheet.createRow(0);
                Cell c1 = infoRow1.getCell(0);
                if (c1 == null) c1 = infoRow1.createCell(0);
                c1.setCellValue(projectName);

                Row infoRow2 = sheet.getRow(1);
                if (infoRow2 == null) infoRow2 = sheet.createRow(1);
                Cell c2 = infoRow2.getCell(0);
                if (c2 == null) c2 = infoRow2.createCell(0);
                c2.setCellValue(ZonedDateTime.now(ZoneId.of("Europe/Rome")).format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));

                Row infoRow3 = sheet.getRow(2);
                if (infoRow3 == null) infoRow3 = sheet.createRow(2);
                Cell c3 = infoRow3.getCell(0);
                if (c3 == null) c3 = infoRow3.createCell(0);
                c3.setCellValue("Ultimo aggiornamento: " + timestamp);

                ByteArrayOutputStream buf = new ByteArrayOutputStream();
                wb.write(buf);
                wb.close();
                byte[] modifiedData = buf.toByteArray();

                // Cancella vecchi consuntivi PRIMA di caricare (così non cancella il nuovo)
                if (deleteOld) {
                    writeEvent(os, "log", "Cerco consuntivi precedenti da cancellare...");
                    int deletedCount = deleteOldConsuntivi(siteId, parentFolder, projectName);
                    writeEvent(os, "log", "Cancellati " + deletedCount + " consuntivi precedenti");
                }

                writeEvent(os, "log", "Caricamento su SharePoint...");
                DriveItem uploaded = uploadFile(siteId, parentFolder, newName, modifiedData);
                String effectiveName = uploaded.name;
                LOG.info("Upload completato: " + uploaded.id + " -> " + effectiveName);

                String outputUrl = uploaded.webUrl;
                writeDone(os, outputUrl, effectiveName, allRows.size());

            } catch (ClientException e) {
                LOG.log(Level.SEVERE, "Graph API error: " + e.getMessage(), e);
                writeEvent(os, "error", "Errore Graph API: " + e.getMessage());
            } catch (Exception e) {
                LOG.log(Level.SEVERE, "Errore: " + e.getMessage(), e);
                writeEvent(os, "error", e.getMessage());
            } finally {
                os.close();
            }

        } catch (Exception e) {
            LOG.log(Level.SEVERE, "Errore iniziale: " + e.getMessage(), e);
            corsResponse(exchange, 500, "{\"error\":\"" + safeJson(e.getMessage()) + "\"}");
        }
    }

    private void corsResponse(HttpExchange exchange, int code, String body) {
        try {
            exchange.getResponseHeaders().add("Access-Control-Allow-Origin", "*");
            exchange.getResponseHeaders().add("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
            exchange.getResponseHeaders().add("Access-Control-Allow-Headers", "Content-Type");
            exchange.getResponseHeaders().add("Content-Type", "application/json");
            byte[] bytes = body.getBytes(StandardCharsets.UTF_8);
            exchange.sendResponseHeaders(code, bytes.length);
            if (bytes.length > 0) {
                try (OutputStream os = exchange.getResponseBody()) {
                    os.write(bytes);
                }
            }
        } catch (Exception e) {
            LOG.log(Level.SEVERE, "Error writing response", e);
        }
    }

    private void writeEvent(OutputStream os, String type, String message) throws IOException {
        JsonObject data = new JsonObject();
        data.addProperty("event", type);
        data.addProperty("message", message);
        byte[] line = (data.toString() + "\n").getBytes(StandardCharsets.UTF_8);
        os.write(line);
        os.flush();
    }

    private void writeProgress(OutputStream os, int current, int total, String fileName, int rows) throws IOException {
        JsonObject data = new JsonObject();
        data.addProperty("event", "progress");
        data.addProperty("current", current);
        data.addProperty("total", total);
        data.addProperty("file", fileName);
        data.addProperty("rows", rows);
        byte[] line = (data.toString() + "\n").getBytes(StandardCharsets.UTF_8);
        os.write(line);
        os.flush();
    }

    private void writeDone(OutputStream os, String url, String fileName, int totalRows) throws IOException {
        JsonObject data = new JsonObject();
        data.addProperty("event", "done");
        data.addProperty("url", url);
        data.addProperty("fileName", fileName);
        data.addProperty("totalRows", totalRows);
        byte[] line = (data.toString() + "\n").getBytes(StandardCharsets.UTF_8);
        os.write(line);
        os.flush();
    }

    static class SharepointUrl {
        String hostname;
        String sitePath;
        String filePath;
    }

    static SharepointUrl parseUrl(String url) {
        try {
            URL u = new URL(url);
            String path = u.getPath();
            int idx = path.indexOf("/sites/");
            if (idx < 0) idx = path.indexOf("/teams/");
            if (idx < 0) return null;
            int next = path.indexOf("/", idx + 7);
            if (next < 0) return null;
            SharepointUrl p = new SharepointUrl();
            p.hostname = u.getHost();
            p.sitePath = path.substring(0, next);
            String rawPath = path.substring(next + 1);
            int libSlash = rawPath.indexOf("/");
            if (libSlash > 0) {
                String firstSegment = java.net.URLDecoder.decode(rawPath.substring(0, libSlash), StandardCharsets.UTF_8);
                if (LIBRARY_NAMES.contains(firstSegment)) {
                    rawPath = rawPath.substring(libSlash + 1);
                }
            }
            p.filePath = java.net.URLDecoder.decode(rawPath, StandardCharsets.UTF_8);
            return p;
        } catch (Exception e) {
            return null;
        }
    }

    private static String parentPath(String path) {
        int i = path.lastIndexOf("/");
        return i >= 0 ? path.substring(0, i) : "";
    }

    private static String safeJson(String s) {
        return s != null
            ? s.replace("\\", "\\\\").replace("\"", "\\\"").replace("\n", " ")
            : "";
    }

    private record GraphClientResult(GraphServiceClient client, TokenCredentialAuthProvider auth) {}

    private GraphClientResult buildGraphClient() {
        Dotenv dotenv = Dotenv.configure().ignoreIfMissing().load();
        String cid = getEnv("AZURE_CLIENT_ID", dotenv);
        String tid = getEnv("AZURE_TENANT_ID", dotenv);
        String sec = getEnv("AZURE_CLIENT_SECRET", dotenv);

        ClientSecretCredential cred = new ClientSecretCredentialBuilder()
            .clientId(cid).clientSecret(sec).tenantId(tid).build();

        List<String> scopes = Arrays.asList("https://graph.microsoft.com/.default");
        TokenCredentialAuthProvider auth = new TokenCredentialAuthProvider(scopes, cred);

        return new GraphClientResult(
            GraphServiceClient.builder().authenticationProvider(auth).buildClient(),
            auth
        );
    }

    private static String getEnv(String key, Dotenv dotenv) {
        String v = System.getenv(key);
        return (v != null && !v.isBlank()) ? v : dotenv.get(key);
    }

    private String resolveSiteId(String host, String site) {
        return graphClient.sites(host + ":" + site).buildRequest().get().id;
    }

    private List<DriveItem> listSvsFiles(String siteId, String rdaPath) {
        List<DriveItem> items = new ArrayList<>();
        try {
            LOG.info("listSvsFiles: cerco in -> " + rdaPath);
            DriveItemCollectionPage page = graphClient.sites(siteId).drive().root()
                .itemWithPath(rdaPath).children().buildRequest().get();
            if (page == null || page.getCurrentPage() == null) {
                LOG.warning("listSvsFiles: page is null");
                return items;
            }
            LOG.info("listSvsFiles: trovati " + page.getCurrentPage().size() + " elementi in RDA");
            for (DriveItem di : page.getCurrentPage()) {
                boolean isFolder = di.folder != null;
                LOG.info("  -> " + di.name + " (folder: " + isFolder + ")");
                if (di.name != null && di.name.startsWith("SVS") && di.name.endsWith(".xlsx")) {
                    items.add(di);
                }
            }
        } catch (Exception e) {
            LOG.warning("listSvsFiles(" + rdaPath + "): " + e.getMessage());
        }
        return items;
    }

    private byte[] downloadTemplate() throws Exception {
        String ictSiteId = "svssrl.sharepoint.com,96c00cdc-6492-44ab-94d9-2f7a25aabfd9,5820c7dd-ac32-4237-9d28-4b52de70fbd3";
        String encodedPath = "TemplateConsuntiviRDA/Template%20pivot%20riassunto%20ordini.xlsx";
        String apiUrl = graphClient.getServiceRoot()
            + "/sites/" + ictSiteId
            + "/drive/root:/" + encodedPath + ":/content";

        URL url = URI.create(apiUrl).toURL();
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");

        String token = authProvider.getAuthorizationTokenAsync(url).get();
        conn.setRequestProperty("Authorization", "Bearer " + token);

        int code = conn.getResponseCode();
        LOG.info("downloadTemplate: GET response code=" + code);
        if (code < 200 || code >= 300) {
            String errorBody;
            try (InputStream es = conn.getErrorStream()) {
                errorBody = es != null ? new String(readAllBytes(es), StandardCharsets.UTF_8) : "no body";
            }
            throw new IOException("GET template failed: " + code + " " + errorBody);
        }

        byte[] data = readAllBytes(conn.getInputStream());
        LOG.info("downloadTemplate: downloaded " + data.length + " bytes");
        return data;
    }

    private DriveItem uploadFile(String siteId, String parentFolder, String fileName, byte[] data) throws Exception {
                LOG.info("uploadFile: salvo in -> " + parentFolder + "/" + fileName + " (" + data.length + " bytes)");
                DriveItem parent = graphClient.sites(siteId).drive().root()
            .itemWithPath(parentFolder).buildRequest().get();
        LOG.info("uploadFile: parent itemId=" + parent.id);

        String encodedName = URLEncoder.encode(fileName, StandardCharsets.UTF_8)
            .replace("+", "%20");

        // Always delete first to avoid SharePoint adding [trash] and customXml entries
        LOG.info("uploadFile: deleting existing file first...");
        try {
            deleteFileWithBypass(siteId, parent.id, fileName);
            LOG.info("uploadFile: previous file deleted");
        } catch (Exception e) {
            LOG.info("uploadFile: no previous file to delete (" + e.getMessage() + ")");
        }

        String apiUrl = graphClient.getServiceRoot()
            + "/sites/" + siteId
            + "/drive/items/" + parent.id + ":/" + encodedName + ":/content";

        URL url = URI.create(apiUrl).toURL();
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setDoOutput(true);
        conn.setRequestMethod("PUT");
        conn.setRequestProperty("Content-Type", "application/octet-stream");
        conn.setRequestProperty("Content-Length", String.valueOf(data.length));

        String token = authProvider.getAuthorizationTokenAsync(url).get();
        conn.setRequestProperty("Authorization", "Bearer " + token);

        try (OutputStream os = conn.getOutputStream()) {
            os.write(data);
        }

        int code = conn.getResponseCode();
        LOG.info("uploadFile: PUT response code=" + code);

        if (code >= 200 && code < 300) {
            String jsonResponse;
            try (InputStream is = conn.getInputStream()) {
                jsonResponse = new String(readAllBytes(is), StandardCharsets.UTF_8);
            }
            LOG.info("uploadFile: response size=" + GSON.fromJson(jsonResponse, JsonObject.class).get("size"));
            return GSON.fromJson(jsonResponse, DriveItem.class);
        }

        String errorBody;
        try (InputStream es = conn.getErrorStream()) {
            errorBody = es != null ? new String(readAllBytes(es), StandardCharsets.UTF_8) : "no body";
        }
        throw new IOException("PUT failed: " + code + " " + errorBody);
    }

    private void deleteFileWithBypass(String siteId, String parentId, String fileName) throws Exception {
        String encodedName = URLEncoder.encode(fileName, StandardCharsets.UTF_8)
            .replace("+", "%20");
        String apiUrl = graphClient.getServiceRoot()
            + "/sites/" + siteId
            + "/drive/items/" + parentId + ":/" + encodedName;

        URL url = URI.create(apiUrl).toURL();
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("DELETE");
        conn.setRequestProperty("Prefer", "bypass-shared-lock");

        String token = authProvider.getAuthorizationTokenAsync(url).get();
        conn.setRequestProperty("Authorization", "Bearer " + token);

        int code = conn.getResponseCode();
        LOG.info("deleteFileWithBypass: DELETE response code=" + code);
        if (code < 200 || code >= 300) {
            String errorBody;
            try (InputStream es = conn.getErrorStream()) {
                errorBody = es != null ? new String(readAllBytes(es), StandardCharsets.UTF_8) : "no body";
            }
            throw new IOException("DELETE with bypass failed: " + code + " " + errorBody);
        }
    }

    private int deleteOldConsuntivi(String siteId, String parentFolder, String projectName) {
        int deleted = 0;
        try {
            DriveItem parent = graphClient.sites(siteId).drive().root()
                .itemWithPath(parentFolder).buildRequest().get();
            DriveItemCollectionPage page = graphClient.sites(siteId).drive().root()
                .itemWithPath(parentFolder).children().buildRequest().get();
            if (page == null || page.getCurrentPage() == null) {
                return 0;
            }
            String exactOld = "Consuntivo_" + projectName + ".xlsx";
            String tsSuffix = "_Consuntivo_" + projectName + ".xlsx";
            for (DriveItem item : page.getCurrentPage()) {
                if (item.name != null && item.name.endsWith(".xlsx")
                    && (item.name.equals(exactOld) || item.name.endsWith(tsSuffix))) {
                    LOG.info("deleteOldConsuntivi: cancellando " + item.name);
                    deleteFileWithBypass(siteId, parent.id, item.name);
                    deleted++;
                }
            }
        } catch (Exception e) {
            LOG.warning("deleteOldConsuntivi: " + e.getMessage());
        }
        return deleted;
    }

    private String resolveProjectBasePath(String siteId, String projectName) {
        List<String> candidates = new ArrayList<>();
        candidates.add(projectName);
        String year = "20" + projectName.substring(0, 2);
        candidates.add("Commesse/" + year + "/" + projectName);
        candidates.add("Documenti Condivisi/" + projectName);
        for (String path : candidates) {
            LOG.info("resolveProjectBasePath: provo -> " + path);
            try {
                var item = graphClient.sites(siteId).drive().root().itemWithPath(path).buildRequest().get();
                if (item != null && item.folder != null) {
                    LOG.info("resolveProjectBasePath: trovato -> " + path);
                    return path;
                }
            } catch (Exception e) {
                LOG.info("  non trovato");
            }
        }
        return null;
    }

    private List<List<Object>> readTabella2(String siteId, String itemId, String itemName) {
        List<List<Object>> result = new ArrayList<>();
        try {
            LOG.info("readTabella2: scarico " + itemName + " (" + itemId + ")");
            byte[] fileBytes;
            try (InputStream is = graphClient.sites(siteId).drive()
                    .items(itemId).content().buildRequest().get()) {
                fileBytes = readAllBytes(is);
            }
            LOG.info("readTabella2: scaricati " + fileBytes.length + " bytes");
            XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(fileBytes));
            XSSFSheet xssfSheet = wb.getSheetAt(0);
            LOG.info("readTabella2: foglio='" + xssfSheet.getSheetName() + "' righe=" + xssfSheet.getLastRowNum());

            XSSFTable table = null;
            for (XSSFTable t : xssfSheet.getTables()) {
                if ("Tabella2".equalsIgnoreCase(t.getName())) {
                    table = t;
                    break;
                }
            }

            if (table == null) {
                LOG.warning("readTabella2(" + itemName + "): tabella Tabella2 non trovata, saltato");
                wb.close();
                return null;
            }

            CellReference startRef = table.getStartCellReference();
            CellReference endRef = table.getEndCellReference();
            int headerRow = startRef.getRow();
            int lastDataRow = endRef.getRow();
            int lastCol = Math.min(endRef.getCol(), COLUMN_NAMES.size() - 1);
            LOG.info("readTabella2: Tabella2 da riga=" + headerRow + " a riga=" + lastDataRow + " ultimaCol=" + lastCol);

            for (int r = headerRow + 1; r <= lastDataRow; r++) {
                Row row = xssfSheet.getRow(r);
                if (row == null) continue;
                List<Object> rowData = new ArrayList<>();
                boolean allEmpty = true;
                for (int c = 0; c <= lastCol; c++) {
                    Cell cell = row.getCell(c);
                    if (cell == null) {
                        rowData.add("");
                        continue;
                    }
                    Object val = switch (cell.getCellType()) {
                        case STRING -> { String s = cell.getStringCellValue().trim(); if (!s.isEmpty()) allEmpty = false; yield s; }
                        case NUMERIC -> { double d = cell.getNumericCellValue(); if (d != 0) allEmpty = false; yield d; }
                        case BOOLEAN -> { allEmpty = false; yield cell.getBooleanCellValue(); }
                        case FORMULA -> {
                            CellType resultType = cell.getCachedFormulaResultType();
                            switch (resultType) {
                                case NUMERIC -> { double d = cell.getNumericCellValue(); if (d != 0) allEmpty = false; yield d; }
                                case STRING -> { String s = cell.getStringCellValue().trim(); if (!s.isEmpty()) allEmpty = false; yield s; }
                                case BOOLEAN -> { allEmpty = false; yield cell.getBooleanCellValue(); }
                                default -> { yield ""; }
                            }
                        }
                        default -> "";
                    };
                    rowData.add(val);
                }
                if (!allEmpty) {
                    result.add(rowData);
                }
            }
            wb.close();
            LOG.info("readTabella2: lette " + result.size() + " righe da " + itemName);
        } catch (Exception e) {
            LOG.warning("readTabella2(" + itemName + "): " + e.getMessage());
        }
        return result;
    }

    private record SiteMetadata(String contentTypeId, String siteGuid) {}

    private SiteMetadata resolveSiteMetadata(String siteId) throws Exception {
        String siteGuid = siteId.split(",")[2];

        String apiUrl = graphClient.getServiceRoot()
            + "/sites/" + siteId + "/drive/list/contentTypes";
        URL url = URI.create(apiUrl).toURL();
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        String token = authProvider.getAuthorizationTokenAsync(url).get();
        conn.setRequestProperty("Authorization", "Bearer " + token);
        conn.setRequestProperty("Accept", "application/json");

        String ctId = null;
        int code = conn.getResponseCode();
        if (code >= 200 && code < 300) {
            String body;
            try (InputStream is = conn.getInputStream()) {
                body = new String(readAllBytes(is), StandardCharsets.UTF_8);
            }
            com.google.gson.JsonObject json = GSON.fromJson(body, com.google.gson.JsonObject.class);
            for (com.google.gson.JsonElement elem : json.getAsJsonArray("value")) {
                com.google.gson.JsonObject ct = elem.getAsJsonObject();
                String name = ct.has("name") ? ct.get("name").getAsString() : "";
                if ("Document".equals(name)) {
                    ctId = ct.get("id").getAsString();
                    break;
                }
            }
            if (ctId == null) {
                com.google.gson.JsonObject first = json.getAsJsonArray("value").get(0).getAsJsonObject();
                ctId = first.get("id").getAsString();
            }
        } else {
            LOG.warning("resolveSiteMetadata: contentTypes API returned " + code);
            ctId = "0x010100CD60E95D8782964D9E9800742678F266";
        }

        LOG.info("resolveSiteMetadata: contentTypeId=" + ctId + ", siteGuid=" + siteGuid);
        return new SiteMetadata(ctId, siteGuid);
    }

    private byte[] patchTemplate(SiteMetadata meta) throws Exception {
        ByteArrayOutputStream result = new ByteArrayOutputStream();
        Map<String, byte[]> entries = new HashMap<>();

        try (InputStream tplIs = getClass().getResourceAsStream("/template-consuntivo.xlsx");
             ZipInputStream zis = new ZipInputStream(tplIs)) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                entries.put(entry.getName(), zis.readAllBytes());
            }
        }

        if (entries.containsKey("docProps/custom.xml")) {
            String customXml = new String(entries.get("docProps/custom.xml"), StandardCharsets.UTF_8);
            customXml = customXml.replace(
                "0x010100E37D1D34D5473C4C8E9B3C97F48F20EB",
                meta.contentTypeId()
            );
            entries.put("docProps/custom.xml", customXml.getBytes(StandardCharsets.UTF_8));
            LOG.info("patchTemplate: patched docProps/custom.xml with contentTypeId=" + meta.contentTypeId());
        }

        if (entries.containsKey("customXml/item2.xml")) {
            String item2 = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>\n"
                + "<p:properties xmlns:p=\"http://schemas.microsoft.com/office/2006/metadata/properties\" "
                + "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" "
                + "xmlns:pc=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\">\n"
                + "<documentManagement>\n"
                + "<TaxCatchAll xmlns=\"" + meta.siteGuid() + "\" xsi:nil=\"true\"/>\n"
                + "</documentManagement>\n"
                + "</p:properties>";
            entries.put("customXml/item2.xml", item2.getBytes(StandardCharsets.UTF_8));
            LOG.info("patchTemplate: patched customXml/item2.xml with siteGuid=" + meta.siteGuid());
        }

        if (entries.containsKey("customXml/item3.xml")) {
            String item3 = new String(entries.get("customXml/item3.xml"), StandardCharsets.UTF_8);
            item3 = item3.replace(
                "ma:contentTypeID=\"0x010100E37D1D34D5473C4C8E9B3C97F48F20EB\"",
                "ma:contentTypeID=\"" + meta.contentTypeId() + "\""
            );
            entries.put("customXml/item3.xml", item3.getBytes(StandardCharsets.UTF_8));
            LOG.info("patchTemplate: patched customXml/item3.xml with contentTypeId=" + meta.contentTypeId());
        }

        if (!entries.containsKey("customXml/item4.xml")) {
            entries.put("customXml/item4.xml",
                ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                + "<FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\">"
                + "<Display>DocumentLibraryForm</Display>"
                + "<Edit>DocumentLibraryForm</Edit>"
                + "<New>DocumentLibraryForm</New>"
                + "</FormTemplates>").getBytes(StandardCharsets.UTF_8));
            entries.put("customXml/_rels/item4.xml.rels",
                ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps\" Target=\"itemProps4.xml\"/>"
                + "</Relationships>").getBytes(StandardCharsets.UTF_8));
            entries.put("customXml/itemProps4.xml",
                ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>\n"
                + "<ds:datastoreItem ds:itemID=\"{4000734F-E707-476E-8A5E-69BBC22E439F}\" xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\"/>")
                .getBytes(StandardCharsets.UTF_8));
            LOG.info("patchTemplate: added customXml/item4.xml (FormTemplates)");
        }

        if (entries.containsKey("[Content_Types].xml")) {
            String ct = new String(entries.get("[Content_Types].xml"), StandardCharsets.UTF_8);
            if (!ct.contains("itemProps4")) {
                ct = ct.replace(
                    "</Types>",
                    "<Override PartName=\"/customXml/itemProps4.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.customXmlProperties+xml\"/></Types>"
                );
                entries.put("[Content_Types].xml", ct.getBytes(StandardCharsets.UTF_8));
                LOG.info("patchTemplate: added itemProps4 to [Content_Types].xml");
            }
        }

        if (entries.containsKey("xl/_rels/workbook.xml.rels")) {
            String rels = new String(entries.get("xl/_rels/workbook.xml.rels"), StandardCharsets.UTF_8);
            if (!rels.contains("item4.xml")) {
                rels = rels.replace(
                    "</Relationships>",
                    "<Relationship Id=\"rId14\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\" Target=\"../customXml/item4.xml\"/></Relationships>"
                );
                entries.put("xl/_rels/workbook.xml.rels", rels.getBytes(StandardCharsets.UTF_8));
                LOG.info("patchTemplate: added item4.xml to workbook.xml.rels");
            }
        }

        try (ZipOutputStream zos = new ZipOutputStream(result)) {
            for (Map.Entry<String, byte[]> e : entries.entrySet()) {
                zos.putNextEntry(new ZipEntry(e.getKey()));
                zos.write(e.getValue());
                zos.closeEntry();
            }
        }

        return result.toByteArray();
    }

    private byte[] modifyWithPoi(List<List<Object>> rows,
                                  List<String> columnNames, String projectName,
                                  String timestamp, String siteId) throws Exception {
        InputStream tpl = getClass().getResourceAsStream("/template-consuntivo.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(tpl);
        tpl.close();
        XSSFSheet targetSheet = wb.getSheetAt(0); // Foglio1
        LOG.info("STEP-INIT: template-pivot loaded, sheet=" + targetSheet.getSheetName());

        // Rows 1-3: metadata (Progetto, Data, Ultimo aggiornamento)
        Row projectRow = targetSheet.getRow(0);
        if (projectRow == null) projectRow = targetSheet.createRow(0);
        Cell pc = projectRow.getCell(0);
        if (pc == null) pc = projectRow.createCell(0);
        pc.setCellValue("Progetto: " + projectName);

        Row dateRow = targetSheet.getRow(1);
        if (dateRow == null) dateRow = targetSheet.createRow(1);
        Cell dc = dateRow.getCell(0);
        if (dc == null) dc = dateRow.createCell(0);
        dc.setCellValue("Data: " + timestamp.split(" ")[0]);

        Row updateRow = targetSheet.getRow(2);
        if (updateRow == null) updateRow = targetSheet.createRow(2);
        Cell uc = updateRow.getCell(0);
        if (uc == null) uc = updateRow.createCell(0);
        uc.setCellValue("Ultimo aggiornamento: " + timestamp);

        // Row 4 (index 3): headers — already correct in template, skip

        // Clear existing data rows (5+)
        int lastDataRow = targetSheet.getLastRowNum();
        for (int r = 4; r <= lastDataRow; r++) {
            Row row = targetSheet.getRow(r);
            if (row != null) {
                targetSheet.removeRow(row);
            }
        }
        LOG.info("STEP-CLEAR: cleared rows 4+ (was " + (lastDataRow + 1) + " rows total)");

        // Column formats for data rows
        String[] colFormats = {
            "@",           // 0  RDA
            "DD/MM/YYYY",  // 1  Data
            "@",           // 2  Cod
            "@",           // 3  Prodotto
            "@",           // 4  UM
            null,          // 5  Colonna1 (generale)
            "#,##0",       // 6  PZ
            "#,##0",       // 7  Quantità
            "#,##0.00 €",  // 8  Prezzo
            "#,##0.00 €",  // 9  Importo
            "@",           // 10 Fornitore
            "#,##0",       // 11 Qt. Min. V.
            "@",           // 12 Note
            "@"            // 13 Mese
        };
        CellStyle[] colStyles = new CellStyle[colFormats.length];
        for (int i = 0; i < colFormats.length; i++) {
            if (colFormats[i] != null) {
                CellStyle cs = wb.createCellStyle();
                cs.setDataFormat(wb.createDataFormat().getFormat(colFormats[i]));
                colStyles[i] = cs;
            }
        }

        int dataStartRow = 4; // row index 4 = row 5 in Excel (first data row)
        for (int ri = 0; ri < rows.size(); ri++) {
            Row excelRow = targetSheet.createRow(dataStartRow + ri);
            List<Object> dataRow = rows.get(ri);
            int cols = Math.min(columnNames.size(), dataRow.size());
            for (int ci = 0; ci < cols && ci < colFormats.length; ci++) {
                Cell cell = excelRow.createCell(ci);
                Object val = dataRow.get(ci);
                if (val == null) {
                    cell.setCellValue("");
                } else if (val instanceof Number n) {
                    cell.setCellValue(n.doubleValue());
                } else if (val instanceof Boolean b) {
                    cell.setCellValue(b);
                } else {
                    cell.setCellValue(val.toString());
                }
                if (colStyles[ci] != null) cell.setCellStyle(colStyles[ci]);
            }
        }
        LOG.info("STEP-DATA: wrote " + rows.size() + " data rows starting at row " + (dataStartRow + 1));

        byte[] rawBytes = dumpWorkbook(wb);
        LOG.info("STEP-WRITE: wb.write done, size=" + rawBytes.length);

        byte[] cleaned = cleanContentTypesOrphans(rawBytes);
        LOG.info("STEP-CLEAN: after cleanContentTypesOrphans, size=" + cleaned.length);

        try {
            java.nio.file.Files.write(
                java.nio.file.Paths.get("/var/www/consuntivi/last-generated.xlsx"),
                cleaned
            );
            LOG.info("STEP-DEBUG: saved last-generated.xlsx (" + cleaned.length + " bytes)");
        } catch (Exception e) {
            LOG.warning("STEP-DEBUG: failed to save debug file: " + e.getMessage());
        }

        wb.close();
        return cleaned;
    }

    private static byte[] dumpWorkbook(XSSFWorkbook wb) throws IOException {
        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        wb.write(buf);
        return buf.toByteArray();
    }

    static byte[] cleanContentTypesOrphans(byte[] xlsx) throws IOException {
        Map<String, byte[]> entries = new HashMap<>();
        String contentTypesName = "[Content_Types].xml";

        try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(xlsx))) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                entries.put(entry.getName(), zis.readAllBytes());
            }
        }

        if (!entries.containsKey(contentTypesName)) {
            LOG.warning("STEP-CLEAN: [Content_Types].xml not found, skipping");
            return xlsx;
        }

        try {
            Document doc = DocumentBuilderFactory.newInstance()
                .newDocumentBuilder()
                .parse(new ByteArrayInputStream(entries.get(contentTypesName)));

            NodeList overrides = doc.getElementsByTagName("Override");
            java.util.List<Element> toRemove = new java.util.ArrayList<>();
            int total = overrides.getLength();

            for (int i = 0; i < total; i++) {
                Element override = (Element) overrides.item(i);
                String partName = override.getAttribute("PartName");
                String zipPath = partName.startsWith("/") ? partName.substring(1) : partName;
                if (!entries.containsKey(zipPath)) {
                    toRemove.add(override);
                    LOG.info("STEP-CLEAN: removing orphan Override: " + partName);
                }
            }

            for (Element el : toRemove) {
                el.getParentNode().removeChild(el);
            }

            StringWriter sw = new StringWriter();
            TransformerFactory.newInstance().newTransformer()
                .transform(new DOMSource(doc), new StreamResult(sw));
            entries.put(contentTypesName, sw.toString().getBytes(StandardCharsets.UTF_8));
            LOG.info("STEP-CLEAN: removed " + toRemove.size() + " orphan Overrides from " + total + " total");

        } catch (Exception e) {
            LOG.warning("STEP-CLEAN: failed to parse [Content_Types].xml: " + e.getMessage());
            return xlsx;
        }

        ByteArrayOutputStream result = new ByteArrayOutputStream();
        try (ZipOutputStream zos = new ZipOutputStream(result)) {
            for (Map.Entry<String, byte[]> e : entries.entrySet()) {
                zos.putNextEntry(new ZipEntry(e.getKey()));
                zos.write(e.getValue());
                zos.closeEntry();
            }
        }
        return result.toByteArray();
    }

    static byte[] stripCustomXml(byte[] xlsx) throws IOException {
        Map<String, byte[]> entries = new LinkedHashMap<>();

        try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(xlsx))) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                entries.put(entry.getName(), zis.readAllBytes());
            }
        }

        // Remove all customXml entries
        int removed = 0;
        var it = entries.keySet().iterator();
        while (it.hasNext()) {
            String name = it.next();
            if (name.startsWith("customXml/")) {
                it.remove();
                removed++;
            }
        }
        LOG.info("STEP-STRIP: removed " + removed + " customXml entries");

        // Remove customXml Overrides from [Content_Types].xml
        String contentTypesName = "[Content_Types].xml";
        if (entries.containsKey(contentTypesName)) {
            try {
                Document doc = DocumentBuilderFactory.newInstance()
                    .newDocumentBuilder()
                    .parse(new ByteArrayInputStream(entries.get(contentTypesName)));

                NodeList overrides = doc.getElementsByTagName("Override");
                java.util.List<Element> toRemove = new java.util.ArrayList<>();
                for (int i = 0; i < overrides.getLength(); i++) {
                    Element override = (Element) overrides.item(i);
                    String partName = override.getAttribute("PartName");
                    if (partName.contains("customXml")) {
                        toRemove.add(override);
                    }
                }
                for (Element el : toRemove) {
                    el.getParentNode().removeChild(el);
                }
                if (!toRemove.isEmpty()) {
                    StringWriter sw = new StringWriter();
                    TransformerFactory.newInstance().newTransformer()
                        .transform(new DOMSource(doc), new StreamResult(sw));
                    entries.put(contentTypesName, sw.toString().getBytes(StandardCharsets.UTF_8));
                    LOG.info("STEP-STRIP: removed " + toRemove.size() + " customXml Overrides from Content_Types");
                }
            } catch (Exception e) {
                LOG.warning("STEP-STRIP: failed to clean Content_Types: " + e.getMessage());
            }
        }

        // Remove customXml relationships from xl/_rels/workbook.xml.rels
        String workbookRelsName = "xl/_rels/workbook.xml.rels";
        if (entries.containsKey(workbookRelsName)) {
            try {
                Document doc = DocumentBuilderFactory.newInstance()
                    .newDocumentBuilder()
                    .parse(new ByteArrayInputStream(entries.get(workbookRelsName)));

                NodeList rels = doc.getElementsByTagName("Relationship");
                java.util.List<Element> toRemove = new java.util.ArrayList<>();
                for (int i = 0; i < rels.getLength(); i++) {
                    Element rel = (Element) rels.item(i);
                    String target = rel.getAttribute("Target");
                    if (target.contains("customXml")) {
                        toRemove.add(rel);
                    }
                }
                for (Element el : toRemove) {
                    el.getParentNode().removeChild(el);
                }
                if (!toRemove.isEmpty()) {
                    StringWriter sw = new StringWriter();
                    TransformerFactory.newInstance().newTransformer()
                        .transform(new DOMSource(doc), new StreamResult(sw));
                    entries.put(workbookRelsName, sw.toString().getBytes(StandardCharsets.UTF_8));
                    LOG.info("STEP-STRIP: removed " + toRemove.size() + " customXml relationships from workbook.xml.rels");
                }
            } catch (Exception e) {
                LOG.warning("STEP-STRIP: failed to clean workbook.xml.rels: " + e.getMessage());
            }
        }

        ByteArrayOutputStream result = new ByteArrayOutputStream();
        try (ZipOutputStream zos = new ZipOutputStream(result)) {
            for (Map.Entry<String, byte[]> e : entries.entrySet()) {
                zos.putNextEntry(new ZipEntry(e.getKey()));
                zos.write(e.getValue());
                zos.closeEntry();
            }
        }
        return result.toByteArray();
    }

    private static byte[] readAllBytes(InputStream is) throws Exception {
        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        byte[] b = new byte[8192];
        int n;
        while ((n = is.read(b)) != -1) buf.write(b, 0, n);
        return buf.toByteArray();
    }

}
