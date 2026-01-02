package MA_MSG_Suite_OB;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openqa.selenium.*;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class DocPathManager {
    private static String docPathLocal;
    private static XWPFDocument sharedDocument;
    private static String currentTestcase; // track testcase

    public synchronized static void reset() {
        docPathLocal = null;
        sharedDocument = null;
        currentTestcase = null;
    }
    public static WebDriver driver;




    public synchronized static String getOrCreateDocPath(String filePath, String testcase) {
        // If no path yet OR testcase changed → regenerate
        if (docPathLocal == null || docPathLocal.isEmpty() ||
                currentTestcase == null || !currentTestcase.equals(testcase)) {

            docPathLocal = buildDocPath(filePath, testcase);
            sharedDocument = new XWPFDocument(); // new doc for new testcase
            currentTestcase = testcase; // update tracker
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


    public static void captureScreenshot(String fileName,WebDriver driver) {
        try {
            File srcFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
            try (FileInputStream fis = new FileInputStream(srcFile)) {
                XWPFDocument document = DocPathManager.getSharedDocument();
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setText("Screenshot: " + fileName);
                run.addBreak();
                run.addPicture(fis,
                        Document.PICTURE_TYPE_PNG,
                        fileName + ".png",
                        Units.toEMU(500),
                        Units.toEMU(300));
            }
            System.out.println("Screenshot added to document: " + fileName);
        } catch (Exception e) {
            System.out.println("Error capturing screenshot: " + e.getMessage());
        }
    }
    public static void captureAllCardsScreenshots(WebDriver driver) throws InterruptedException, IOException {
        XWPFDocument document = DocPathManager.getSharedDocument(); // shared doc
        List<WebElement> rows = driver.findElements(By.cssSelector("[role='main'] card-view"));
        int i = 1;
        for (WebElement row : rows) {
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'});", row);
            Thread.sleep(500);
            captureScreenshotRow(row, i, document);
            Thread.sleep(800);
            i++;
        }
    }

    public static void captureScreenshotRow(WebElement ele, int i, XWPFDocument document) {
        try {
            File srcFile = ele.getScreenshotAs(OutputType.FILE);
            try (FileInputStream fis = new FileInputStream(srcFile)) {
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setText("Card Row Screenshot: " + i);
                run.addBreak();
                run.addPicture(fis, Document.PICTURE_TYPE_PNG, i + ".png", Units.toEMU(500), Units.toEMU(100));
            }
            System.out.println("Row screenshot added: " + i);
        } catch (Exception e) {
            System.out.println("Error capturing row screenshot: " + e.getMessage());
        }
    }






//    public static void captureAll1(WebDriver driver) throws InterruptedException, IOException {
//        XWPFDocument document = DocPathManager.getSharedDocument(); // shared doc
//        JavascriptExecutor js = (JavascriptExecutor) driver;
//        WebDriverWait wait = new WebDriverWait(driver, java.time.Duration.ofSeconds(10));
//
//        // 1) Locate the ACE editor inside modal-content (exact container you shared)
//        WebElement editor = wait.until(ExpectedConditions.visibilityOfElementLocated(
//                By.cssSelector("modal-content .app-ace-editor")
//        ));
//
//        // 2) Choose a viewport element to screenshot (prefer scroller; fallback to content)
//        WebElement viewport = null;
//        try {
//            viewport = editor.findElement(By.cssSelector(".ace_scroller"));
//        } catch (NoSuchElementException e) {
//            viewport = editor.findElement(By.cssSelector(".ace_content"));
//        }
//
//        // 3) Locate the vertical scrollbar in this editor
//        WebElement vScroll = wait.until(ExpectedConditions.visibilityOfElementLocated(
//                By.cssSelector("modal-content .app-ace-editor .ace_scrollbar-v")
//        ));
//
//        // Bring the editor into view once
//        js.executeScript("arguments[0].scrollIntoView({block:'center'});", editor);
//        Thread.sleep(250);
//
//        // 4) Get scroll metrics from the vertical scrollbar (safe Number casting)
//        long clientHeight = ((Number) js.executeScript("return arguments[0].clientHeight;", vScroll)).longValue();
//        long innerHeight = ((Number) js.executeScript(
//                "var inner=arguments[0].querySelector('.ace_scrollbar-inner');" +
//                        "return inner?inner.offsetHeight:arguments[0].scrollHeight;",
//                vScroll)).longValue();
//
//        long maxScrollTop = Math.max(0, innerHeight - clientHeight);
//
//        // 5) Step size: use the viewport height so each capture is a new slice
//        long step = ((Number) js.executeScript("return arguments[0].clientHeight;", viewport)).longValue();
//        if (step <= 0) step = clientHeight; // fallback
//
//        // Optional: small overlap to avoid gaps between slices (e.g., 10px)
//        long overlap = 10;
//        if (step > overlap) step = step - overlap;
//
//        // 6) Scroll to top first
//        js.executeScript("arguments[0].scrollTop = 0;", vScroll);
//        Thread.sleep(300);
//
//        // 7) Loop through the container height and grab a screenshot each time
//        for (long y = 1; y <= maxScrollTop; y += step) {
//            long targetY = Math.min(y, maxScrollTop);
//
//            js.executeScript("arguments[0].scrollTop = arguments[1];", vScroll, targetY);
//            Thread.sleep(300); // allow render to settle
//
//            // Screenshot the visible editor viewport (your existing method)
//            captureScreenshot1(viewport,document);
//
//            Thread.sleep(200); // pacing
//        }
//
//
//
//
//        js.executeScript("arguments[0].scrollTop = arguments[1];", vScroll, maxScrollTop);
//        captureScreenshot1(viewport,document);
//
//        // 8) Save once at the end
//       // saveDocument("PayloadScreenshots");
//       // saveSharedDocument();
//
//
//
//    }





    public static void captureAll1(String filePath,WebDriver driver) throws InterruptedException, IOException {

        XWPFDocument document = DocPathManager.getSharedDocument();
        JavascriptExecutor js = (JavascriptExecutor) driver;
        WebDriverWait wait = new WebDriverWait(driver, java.time.Duration.ofSeconds(10));

        // 1) Locate ACE editor
        WebElement editor = wait.until(ExpectedConditions.visibilityOfElementLocated(
                By.cssSelector("modal-content .app-ace-editor")
        ));

        // 2) Determine viewport
        WebElement viewport;
        try {
            viewport = editor.findElement(By.cssSelector(".ace_scroller"));
        } catch (NoSuchElementException e) {
            viewport = editor.findElement(By.cssSelector(".ace_content"));
        }

        // 3) Try to locate scrollbar — if missing → captureScreenshot() and exit
        WebElement vScroll = null;
        try {
            vScroll = wait.until(ExpectedConditions.visibilityOfElementLocated(
                    By.cssSelector("modal-content .app-ace-editor .ace_scrollbar-v")
            ));
        } catch (Exception e) {
            System.out.println("⚠️ Scrollbar not found or not visible. Capturing fallback screenshot.");
          //  captureScreenshot(filePath,driver);   // <-- your fallback call
            return;                // stop execution safely
        }

        // Bring editor into view
        js.executeScript("arguments[0].scrollIntoView({block:'center'});", editor);
        Thread.sleep(250);

        // 4) Scroll metrics
        long clientHeight = ((Number) js.executeScript("return arguments[0].clientHeight;", vScroll)).longValue();
        long innerHeight = ((Number) js.executeScript(
                "var inner=arguments[0].querySelector('.ace_scrollbar-inner');" +
                        "return inner?inner.offsetHeight:arguments[0].scrollHeight;",
                vScroll)).longValue();

        long maxScrollTop = Math.max(0, innerHeight - clientHeight);

        // 5) Step size
        long step = ((Number) js.executeScript("return arguments[0].clientHeight;", viewport)).longValue();
        if (step <= 0) step = clientHeight;

        long overlap = 10;
        if (step > overlap) step -= overlap;

        // 6) Scroll to top
        js.executeScript("arguments[0].scrollTop = 0;", vScroll);
        Thread.sleep(300);

        // 7) Loop through scrollable area
        for (long y = 1; y <= maxScrollTop; y += step) {
            long targetY = Math.min(y, maxScrollTop);

            js.executeScript("arguments[0].scrollTop = arguments[1];", vScroll, targetY);
            Thread.sleep(300);

            captureScreenshot1(viewport, document);
            Thread.sleep(200);
        }

        // Final bottom capture
        js.executeScript("arguments[0].scrollTop = arguments[1];", vScroll, maxScrollTop);
        captureScreenshot1(viewport, document);

        // Save happens outside
    }



    public static void captureScreenshot1(WebElement ele,XWPFDocument document) {
        try {
            File srcFile = ele.getScreenshotAs(OutputType.FILE);
            try (FileInputStream fis = new FileInputStream(srcFile)) {
                int wPx = ele.getSize().getWidth();
                int hPx = ele.getSize().getHeight();

                int maxWEmu = Units.toEMU(600);   // ~6 inches
                int maxHEmu = Units.toEMU(1000);  // ~8 inches

                int widthEmu = Units.pixelToEMU(wPx);
                int heightEmu = Units.pixelToEMU(hPx);

                if (widthEmu > maxWEmu) widthEmu = maxWEmu;
                if (heightEmu > maxHEmu) heightEmu = maxHEmu;

                XWPFParagraph p = document.createParagraph();
                XWPFRun r = p.createRun();
                r.setText("Screenshot:");
                r.addBreak();
                r.addPicture(fis, Document.PICTURE_TYPE_PNG, "img.png", widthEmu, heightEmu);

                System.out.println("Screenshot added.");
            }
        } catch (Exception e) {
            System.out.println("Error: " + e.getMessage());
        }
    }



}

