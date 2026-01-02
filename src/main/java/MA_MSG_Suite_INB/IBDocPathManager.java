package MA_MSG_Suite_INB;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openqa.selenium.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

public class IBDocPathManager {

    private static String docPathLocal;
    private static XWPFDocument sharedDocument;
    private static String currentTestcase;

    // ---------------- RESET ----------------
    public synchronized static void reset() {
        docPathLocal = null;
        sharedDocument = null;
        currentTestcase = null;
    }

    // ---------------- DOC PATH ----------------
    public synchronized static String getOrCreateDocPath(String excelPath, String testcase) {
        if (docPathLocal == null || !testcase.equals(currentTestcase)) {
            docPathLocal = buildDocPath(excelPath, testcase);
            sharedDocument = new XWPFDocument();
            currentTestcase = testcase;
        }
        return docPathLocal;
    }

    // ---------------- SHARED DOC ----------------
    public synchronized static XWPFDocument getSharedDocument() {
        if (sharedDocument == null) {
            sharedDocument = new XWPFDocument();
        }
        return sharedDocument;
    }

    // ---------------- SAVE DOC ----------------
    public synchronized static void saveSharedDocument() {
        if (docPathLocal == null || sharedDocument == null) return;

        try (FileOutputStream out = new FileOutputStream(docPathLocal)) {
            sharedDocument.write(out);
            System.out.println("Document saved at: " + docPathLocal);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // ---------------- BUILD PATH ----------------
    private static String buildDocPath(String excelPathStr, String baseName) {
        Path excelPath = Paths.get(excelPathStr);
        Path parent = excelPath.getParent() != null ? excelPath.getParent() : Paths.get(".");
        String stamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        return parent.resolve(baseName + "_" + stamp + ".docx").toString();
    }

    // ---------------- SCREENSHOT (FULL PAGE) ----------------
    public static void captureScreenshot(WebDriver driver, String fileName) {
        if (driver == null) {
            throw new RuntimeException("WebDriver is NULL while capturing screenshot");
        }

        try {
            File srcFile = ((TakesScreenshot) driver)
                    .getScreenshotAs(OutputType.FILE);

            try (FileInputStream fis = new FileInputStream(srcFile)) {
                XWPFDocument document = getSharedDocument();

                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();

                run.setText("Screenshot: " + fileName);
                run.addBreak();

                run.addPicture(
                        fis,
                        Document.PICTURE_TYPE_PNG,
                        fileName + ".png",
                        Units.toEMU(500),
                        Units.toEMU(300)
                );
            }

            System.out.println("Screenshot added: " + fileName);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ---------------- CARD ROW SCREENSHOTS ----------------
    public static void captureAllCardsScreenshots(WebDriver driver)
            throws InterruptedException, IOException {

        if (driver == null) {
            throw new RuntimeException("WebDriver is NULL while capturing cards");
        }

        XWPFDocument document = getSharedDocument();
        List<WebElement> rows =
                driver.findElements(By.cssSelector("[role='main'] card-view"));

        int i = 1;
        for (WebElement row : rows) {
            ((JavascriptExecutor) driver)
                    .executeScript("arguments[0].scrollIntoView({block:'center'});", row);

            Thread.sleep(500);
            captureScreenshotRow(row, i, document);
            Thread.sleep(800);
            i++;
        }
    }

    // ---------------- ELEMENT SCREENSHOT ----------------
    private static void captureScreenshotRow(
            WebElement ele, int index, XWPFDocument document) {

        try {
            File srcFile = ele.getScreenshotAs(OutputType.FILE);

            try (FileInputStream fis = new FileInputStream(srcFile)) {
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();

                run.setText("Card Row Screenshot: " + index);
                run.addBreak();

                run.addPicture(
                        fis,
                        Document.PICTURE_TYPE_PNG,
                        index + ".png",
                        Units.toEMU(500),
                        Units.toEMU(120)
                );
            }

            System.out.println("Row screenshot added: " + index);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
