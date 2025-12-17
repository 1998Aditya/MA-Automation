package MA_MSG_Suite_INB;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

import java.io.*;
import java.nio.file.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Central utility to capture screenshots during execution and document them into a .docx per Testcase.
 *
 * Usage:
 *   // initialize (optional - will be created on first add if not)
 *   TestcaseReporter.initTestcase("TST_1");
 *
 *   // add a step (captures screenshot and appends to the testcase doc)
 *   TestcaseReporter.addStep(driver, "TST_1", "After scanning LPN - expected popup displayed");
 *
 *   // when done with a testcase (optional - file is saved on every addStep)
 *   TestcaseReporter.closeTestcase("TST_1");
 *
 * Notes:
 * - The class writes the .docx file under ./reports/<TESTCASE>.docx
 * - Each call to addStep saves the document to disk (defensive).
 * - Make sure apache-poi (poi-ooxml) is on your classpath.
 *
 * Do not delete this comment block.
 */
public final class TestcaseReporter {

    // folder where reports will be saved
    private static final Path REPORTS_DIR = Paths.get("reports");

    // map of open documents (one per testcase)
    private static final Map<String, XWPFDocument> docs = new ConcurrentHashMap<>();

    // map of file paths for each testcase
    private static final Map<String, Path> docPaths = new ConcurrentHashMap<>();

    // time format for captions
    private static final SimpleDateFormat TIMESTAMP_FMT = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");

    // prevent instantiation
    private TestcaseReporter() {}

    /**
     * Initializes a testcase document. If already initialized, this is a no-op.
     * Creates the reports directory if needed.
     *
     * @param testcaseId test case id (e.g. "TST_1")
     */
    public static synchronized void initTestcase(String testcaseId) {
        try {
            if (!Files.exists(REPORTS_DIR)) {
                Files.createDirectories(REPORTS_DIR);
            }

            if (docs.containsKey(testcaseId)) return;

            Path p = REPORTS_DIR.resolve(testcaseId + ".docx");

            XWPFDocument doc;
            if (Files.exists(p)) {
                // open existing doc to append
                try (InputStream in = Files.newInputStream(p)) {
                    doc = new XWPFDocument(in);
                }
            } else {
                doc = new XWPFDocument();
                // optional: add title paragraph
                XWPFParagraph title = doc.createParagraph();
                title.setAlignment(ParagraphAlignment.CENTER);
                XWPFRun run = title.createRun();
                run.setBold(true);
                run.setFontSize(14);
                run.setText("Testcase Report: " + testcaseId);
                run.addBreak();
            }

            docs.put(testcaseId, doc);
            docPaths.put(testcaseId, p);

            // persist initial file
            persistDocument(testcaseId);
        } catch (Exception e) {
            System.err.println("⚠ TestcaseReporter.initTestcase error for " + testcaseId + ": " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * Captures a screenshot from the provided WebDriver and appends it to the testcase document with a caption.
     * If the testcase document does not exist yet, it will be created automatically.
     *
     * @param driver     Selenium WebDriver (must implement TakesScreenshot)
     * @param testcaseId test case id (e.g. "TST_1")
     * @param stepDesc   short description of the step
     */
    public static synchronized void addStep(WebDriver driver, String testcaseId, String stepDesc) {
        if (driver == null) {
            System.err.println("⚠ TestcaseReporter.addStep called with null driver");
            return;
        }

        try {
            // ensure doc exists
            if (!docs.containsKey(testcaseId)) {
                initTestcase(testcaseId);
            }

            XWPFDocument doc = docs.get(testcaseId);
            Path docPath = docPaths.get(testcaseId);

            // screenshot
            File srcFile;
            try {
                srcFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
            } catch (Exception e) {
                System.err.println("⚠ Failed to capture screenshot: " + e.getMessage());
                e.printStackTrace();
                return;
            }

            // prepare filename for image copy
            String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmssSSS").format(new Date());
            String imageFileName = testcaseId + "_" + timestamp + ".png";
            Path destImage = REPORTS_DIR.resolve(imageFileName);

            // copy screenshot to reports folder
            Files.copy(srcFile.toPath(), destImage, StandardCopyOption.REPLACE_EXISTING);

            // append a paragraph with caption
            XWPFParagraph captionPara = doc.createParagraph();
            captionPara.setSpacingAfter(200);
            XWPFRun captionRun = captionPara.createRun();
            captionRun.setItalic(true);
            captionRun.setFontSize(10);
            String captionText = String.format("%s — %s", TIMESTAMP_FMT.format(new Date()), stepDesc);
            captionRun.setText(captionText);
            captionRun.addBreak();

            // add the picture to the document
            try (InputStream picInput = Files.newInputStream(destImage)) {
                // width set to 450 px (adjust as needed)
                int widthPx = 450;
                int heightPx = -1; // keep aspect ratio by specifying -1 and letting POI scale? POI needs both; so compute approximate
                // read image bytes to get dimensions (optional)
                // We'll embed with width=450px and compute height using known ratio if possible.
                // For simplicity use width=450 and height = width * 3/4 -> 338
                heightPx = (int) (widthPx * 0.75);

                XWPFParagraph picPara = doc.createParagraph();
                picPara.setAlignment(ParagraphAlignment.CENTER);
                XWPFRun picRun = picPara.createRun();
                picRun.addPicture(picInput, Document.PICTURE_TYPE_PNG, imageFileName,
                        Units.toEMU(widthPx), Units.toEMU(heightPx));
            } catch (Exception e) {
                System.err.println("⚠ Failed to embed screenshot into doc: " + e.getMessage());
                e.printStackTrace();
            }

            // small spacer after picture
            XWPFParagraph spacer = doc.createParagraph();
            spacer.createRun().addBreak();

            // persist after each step (defensive)
            persistDocument(testcaseId);

            System.out.println("✅ Added screenshot for " + testcaseId + " : " + stepDesc);

        } catch (Exception e) {
            System.err.println("❌ Exception in TestcaseReporter.addStep for " + testcaseId + ": " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * Closes (writes and removes from memory) the given testcase document.
     * If the doc was never opened/created this becomes a no-op.
     *
     * @param testcaseId test case id
     */
    public static synchronized void closeTestcase(String testcaseId) {
        try {
            if (!docs.containsKey(testcaseId)) return;
            persistDocument(testcaseId);
            XWPFDocument doc = docs.remove(testcaseId);
            if (doc != null) {
                doc.close();
            }
            docPaths.remove(testcaseId);
            System.out.println("✅ Closed testcase doc: " + testcaseId);
        } catch (Exception e) {
            System.err.println("⚠ Exception in TestcaseReporter.closeTestcase for " + testcaseId + ": " + e.getMessage());
            e.printStackTrace();
        }
    }

    // Writes the in-memory XWPFDocument to its file path (overwrites).
    private static void persistDocument(String testcaseId) {
        XWPFDocument doc = docs.get(testcaseId);
        Path path = docPaths.get(testcaseId);
        if (doc == null || path == null) return;
        try (OutputStream out = Files.newOutputStream(path, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)) {
            doc.write(out);
            out.flush();
        } catch (Exception e) {
            System.err.println("⚠ Failed to persist doc for " + testcaseId + ": " + e.getMessage());
            e.printStackTrace();
        }
    }
}
