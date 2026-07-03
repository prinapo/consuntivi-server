package com.server;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.logging.Logger;

import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpHandler;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class TestDownloadHandler implements HttpHandler {

    private static final Logger LOG = Logger.getLogger(TestDownloadHandler.class.getName());

    @Override
    public void handle(HttpExchange exchange) throws IOException {
        try {
            doHandle(exchange);
        } catch (Exception e) {
            LOG.severe("test-download error: " + e.getMessage());
            String err = "{\"error\":\"" + e.getMessage() + "\"}";
            exchange.getResponseHeaders().set("Content-Type", "application/json");
            exchange.sendResponseHeaders(500, err.length());
            exchange.getResponseBody().write(err.getBytes());
            exchange.getResponseBody().close();
        }
    }

    private void doHandle(HttpExchange exchange) throws Exception {
        LOG.info("test-download: generating test workbook");

        InputStream tpl = getClass().getResourceAsStream("/template-consuntivo.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(tpl);
        tpl.close();
        XSSFSheet targetSheet = wb.getSheetAt(0); // Foglio1

        // Update metadata rows
        Row projectRow = targetSheet.getRow(0);
        if (projectRow == null) projectRow = targetSheet.createRow(0);
        Cell pc = projectRow.getCell(0);
        if (pc == null) pc = projectRow.createCell(0);
        pc.setCellValue("Progetto: Test Download");

        Row dateRow = targetSheet.getRow(1);
        if (dateRow == null) dateRow = targetSheet.createRow(1);
        Cell dc = dateRow.getCell(0);
        if (dc == null) dc = dateRow.createCell(0);
        dc.setCellValue("Data: " + LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));

        Row updateRow = targetSheet.getRow(2);
        if (updateRow == null) updateRow = targetSheet.createRow(2);
        Cell uc = updateRow.getCell(0);
        if (uc == null) uc = updateRow.createCell(0);
        uc.setCellValue("Ultimo aggiornamento: " + LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss")));

        // Clear existing data rows (5+)
        int lastDataRow = targetSheet.getLastRowNum();
        for (int r = 4; r <= lastDataRow; r++) {
            Row row = targetSheet.getRow(r);
            if (row != null) {
                targetSheet.removeRow(row);
            }
        }

        String[] colFormats = {
            "@", "@", "@", "@", "@", null, "#,##0", "#,##0", "#,##0.00 €", "#,##0.00 €", "@", "#,##0", "@", "@"
        };
        CellStyle[] colStyles = new CellStyle[colFormats.length];
        for (int i = 0; i < colFormats.length; i++) {
            if (colFormats[i] != null) {
                CellStyle cs = wb.createCellStyle();
                cs.setDataFormat(wb.createDataFormat().getFormat(colFormats[i]));
                colStyles[i] = cs;
            }
        }

        String[][] fakeData = {
            {"RDA-001", "15/06/2026", "C001", "Profilo Guida U 40/75/40", "MT", "", "10", "50", "12.50", "625.00", "Fornitore A", "5", "", "06/2026"},
            {"RDA-001", "15/06/2026", "C002", "Paraspigolo 2004", "PZ", "", "100", "200", "3.75", "750.00", "Fornitore A", "20", "", "06/2026"},
            {"RDA-002", "20/06/2026", "C003", "Stucco EVOPLUS 60", "KG", "", "1", "150", "8.90", "1335.00", "Fornitore B", "10", "", "06/2026"},
            {"RDA-002", "20/06/2026", "C004", "Nastro Garza Rete", "MT", "", "90", "30", "2.10", "63.00", "Fornitore B", "5", "", "06/2026"},
            {"RDA-003", "25/06/2026", "C005", "Vite V.R. 3.5x35", "PZ", "", "1000", "500", "0.08", "40.00", "Fornitore C", "100", "", "06/2026"},
            {"RDA-003", "25/06/2026", "C006", "Tassello Nylon 6x30", "PZ", "", "100", "300", "0.12", "36.00", "Fornitore C", "20", "", "06/2026"},
            {"RDA-004", "01/07/2026", "C007", "Profilo Montante C 48/74/50", "MT", "", "6", "80", "15.00", "1200.00", "Fornitore A", "6", "", "07/2026"},
            {"RDA-004", "01/07/2026", "C008", "Profilo Guida U 28/27/28", "MT", "", "6", "60", "11.25", "675.00", "Fornitore A", "6", "", "07/2026"},
            {"RDA-005", "05/07/2026", "C009", "Lastra PREGYPLAC PLUS 13mm", "PZ", "", "1", "40", "22.50", "900.00", "Fornitore D", "1", "", "07/2026"},
            {"RDA-005", "05/07/2026", "C010", "Tasello TEKS 4.2x12.7", "PZ", "", "1000", "200", "0.05", "10.00", "Fornitore D", "50", "", "07/2026"},
        };

        int dataStartRow = 4;
        for (int ri = 0; ri < fakeData.length; ri++) {
            Row excelRow = targetSheet.createRow(dataStartRow + ri);
            for (int ci = 0; ci < fakeData[ri].length; ci++) {
                Cell cell = excelRow.createCell(ci);
                String val = fakeData[ri][ci];
                if (ci == 6 || ci == 7 || ci == 11) {
                    cell.setCellValue(Integer.parseInt(val));
                } else if (ci == 8 || ci == 9) {
                    cell.setCellValue(Double.parseDouble(val));
                } else {
                    cell.setCellValue(val);
                }
                if (colStyles[ci] != null) cell.setCellStyle(colStyles[ci]);
            }
        }

        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        wb.write(buf);
        wb.close();

        byte[] cleaned = ConsolidaHandler.cleanContentTypesOrphans(buf.toByteArray());

        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        String fileName = "test_download_" + timestamp + ".xlsx";

        exchange.getResponseHeaders().set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        exchange.getResponseHeaders().set("Content-Disposition", "attachment; filename=\"" + fileName + "\"");
        exchange.sendResponseHeaders(200, cleaned.length);
        exchange.getResponseBody().write(cleaned);
        exchange.getResponseBody().close();

        LOG.info("test-download: served " + fileName + " (" + cleaned.length + " bytes)");
    }
}
