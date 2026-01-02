package MA_MSG_Suite_INB;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.concurrent.atomic.AtomicLong;

/**
 * LPNgenerator
 * Generates unique sequential 8-digit iLPN IDs starting with '8'.
 * Stores and retrieves the last LPN value from the "LPNgenerator" sheet
 * in the Excel file located at  ExcelReaderIB.LPN_EXCEL_PATH.
 */
public class LPNgenerator {

    private static final String EXCEL_PATH =  ExcelReaderIB.LPN_EXCEL_PATH;

    private static final String SHEET_NAME = "LPNgenerator";
    private static final long START_NUMBER = 80000000L;

    private static final AtomicLong counter = new AtomicLong();

    static {
        counter.set(readLastLpnFromExcel());
    }

    /**
     * Returns the next sequential LPN number (8 digits, starts with 8).
     * Example: 80000001, 80000002, 80000003, ...
     */
    public static synchronized String getNextLpn() {
        long next = counter.incrementAndGet();
        saveLastLpnToExcel(next);
        return String.valueOf(next);
    }

    /**
     * Reads the last used LPN value from the Excel file.
     * If sheet or value not found, starts from START_NUMBER.
     */
    private static long readLastLpnFromExcel() {
        File file = new File(EXCEL_PATH);
        if (!file.exists()) {
            System.out.println("⚠ Excel file not found. Starting from " + START_NUMBER);
            return START_NUMBER;
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheet(SHEET_NAME);
            if (sheet == null) {
                System.out.println("ℹ Sheet 'LPNgenerator' not found. Creating new one.");
                createSheetWithDefaultValue(workbook, file);
                return START_NUMBER;
            }

            Row row = sheet.getRow(1);
            if (row == null || row.getCell(0) == null) {
                System.out.println("ℹ No existing LPN value found. Starting from " + START_NUMBER);
                return START_NUMBER;
            }

            DataFormatter formatter = new DataFormatter();
            String value = formatter.formatCellValue(row.getCell(0)).trim();
            if (value.matches("\\d{8}")) {
                return Long.parseLong(value);
            } else {
                System.out.println("⚠ Invalid LPN value in Excel. Resetting to " + START_NUMBER);
                return START_NUMBER;
            }

        } catch (Exception e) {
            e.printStackTrace();
            return START_NUMBER;
        }
    }

    /**
     * Writes the new last LPN value to the Excel file.
     * Automatically creates the sheet if it doesn't exist.
     */
    private static void saveLastLpnToExcel(long lastLpn) {
        File file = new File(EXCEL_PATH);
        Workbook workbook = null;
        FileInputStream fis = null;

        try {
            if (file.exists()) {
                fis = new FileInputStream(file);
                workbook = WorkbookFactory.create(fis);
            } else {
                workbook = new XSSFWorkbook();
            }

            Sheet sheet = workbook.getSheet(SHEET_NAME);
            if (sheet == null) {
                sheet = workbook.createSheet(SHEET_NAME);
                Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("LastLPN");
            }

            Row row = sheet.getRow(1);
            if (row == null) row = sheet.createRow(1);
            Cell cell = row.getCell(0);
            if (cell == null) cell = row.createCell(0);
            cell.setCellValue(lastLpn);

            if (fis != null) fis.close(); // ensure not locked
            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (fis != null) fis.close();
                if (workbook != null) workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * Creates the LPNgenerator sheet with a starting value.
     */
    private static void createSheetWithDefaultValue(Workbook workbook, File file) {
        try {
            Sheet sheet = workbook.createSheet(SHEET_NAME);
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("LastLPN");

            Row data = sheet.createRow(1);
            data.createCell(0).setCellValue(START_NUMBER);

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
            System.out.println("✅ Created 'LPNgenerator' sheet and initialized counter at " + START_NUMBER);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
/*
    // Optional: reset function if you ever want to restart the sequence manually
    public static synchronized void resetLpnCounter() {
        counter.set(START_NUMBER);
        saveLastLpnToExcel(START_NUMBER);
        System.out.println("✅ LPN counter reset to: " + START_NUMBER);
    }

    // For testing
    public static void main(String[] args) {
        System.out.println("Generated LPN: " + LPNgenerator.getNextLpn());
        System.out.println("Generated LPN: " + LPNgenerator.getNextLpn());
        System.out.println("Generated LPN: " + LPNgenerator.getNextLpn());
    }

 */
}
