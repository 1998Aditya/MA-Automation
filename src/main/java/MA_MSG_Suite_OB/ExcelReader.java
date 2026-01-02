package MA_MSG_Suite_OB;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ExcelReader {

    private static final String FILE_PATH = ExcelReaderOB.LOGIN_EXCEL_PATH; // <-- set once
    private static final String SHEET_NAME = "Login"; // <-- set once

    private Workbook workbook;
    private Sheet sheet;

    public ExcelReader() throws IOException {
        FileInputStream fis = new FileInputStream(FILE_PATH);
        workbook = new XSSFWorkbook(fis);
        sheet = workbook.getSheet(SHEET_NAME);
        if (sheet == null) {
            throw new RuntimeException("Sheet not found: " + SHEET_NAME);
        }
    }

    /** Get cell value by row and column index */
    public String getCellValue(int rowIndex, int colIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) return "";
        Cell cell = row.getCell(colIndex);
        return cell != null ? cell.toString().trim() : "";
    }

    /** Get cell value by header name (first row assumed as header) */
    public String getCellValueByHeader(int rowIndex, String headerName) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) throw new RuntimeException("No header row found.");

        Map<String, Integer> headerMap = new HashMap<>();
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            if (cell != null) {
                headerMap.put(cell.getStringCellValue().trim(), i);
            }
        }

        Integer colIndex = headerMap.get(headerName);
        if (colIndex == null) throw new RuntimeException("Header not found: " + headerName);

        return getCellValue(rowIndex, colIndex);
    }

    public void close() throws IOException {
        workbook.close();
    }
}