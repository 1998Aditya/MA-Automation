package MA_MSG_Suite_OB;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class DocPathManager {
    private static String docPathLocal;
    private static XWPFDocument sharedDocument;

    public synchronized static String getOrCreateDocPath(String filePath, String testcase) {
        if (docPathLocal == null || docPathLocal.isEmpty()) {
            docPathLocal = buildDocPath(filePath, testcase);
            sharedDocument = new XWPFDocument(); // create once
        }
        return docPathLocal;
    }

    public synchronized static XWPFDocument getSharedDocument() {
        if (sharedDocument == null) {
            sharedDocument = new XWPFDocument();
        }
        return sharedDocument;
    }

    public synchronized static void saveSharedDocument() {
        if (docPathLocal != null && sharedDocument != null) {
            try (FileOutputStream out = new FileOutputStream(docPathLocal)) {
                sharedDocument.write(out);
                System.out.println("Document saved at: " + docPathLocal);
            } catch (IOException e) {
                System.out.println("Error saving document: " + e.getMessage());
            }
        }
    }



    public static String buildDocPath(String excelPathStr, String baseName) {
        Path excelPath = Paths.get(excelPathStr);
        Path parent = excelPath.getParent() != null ? excelPath.getParent() : Paths.get(".");
        String stamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        String unique = baseName + "_" + stamp + ".docx";
        return parent.resolve(unique).toString();
    }

}
