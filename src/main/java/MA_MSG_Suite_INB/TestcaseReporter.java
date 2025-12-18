package MA_MSG_Suite_INB;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openqa.selenium.*;
import org.openqa.selenium.support.ui.WebDriverWait;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.*;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.*;
import java.util.List;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Central utility to capture FULL-PAGE screenshots and document them into a .docx per Testcase.
 *
 * One DOCX is generated per testcase (TST_1.docx, TST_2.docx, etc.)
 * Screenshots are appended step-by-step.
 *
 * ⚠ DO NOT DELETE THIS COMMENT BLOCK
 */
public final class TestcaseReporter {

    // 📁 Reports directory
    private static final Path REPORTS_DIR = Paths.get("reports");

    // 📘 One document per testcase
    private static final Map<String, XWPFDocument> DOCS = new ConcurrentHashMap<>();
    private static final Map<String, Path> DOC_PATHS = new ConcurrentHashMap<>();

    private static final SimpleDateFormat TS_FMT =
            new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");

    private TestcaseReporter() {}

    // =========================================================
    // ✅ INIT TESTCASE
    // =========================================================
    public static synchronized void initTestcase(String testcaseId) {
        try {
            Files.createDirectories(REPORTS_DIR);

            if (DOCS.containsKey(testcaseId)) return;

            Path docPath = REPORTS_DIR.resolve(testcaseId + ".docx");
            XWPFDocument doc = new XWPFDocument();

            XWPFParagraph title = doc.createParagraph();
            title.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun run = title.createRun();
            run.setBold(true);
            run.setFontSize(14);
            run.setText("Execution Evidence : " + testcaseId);
            run.addBreak();

            DOCS.put(testcaseId, doc);
            DOC_PATHS.put(testcaseId, docPath);

            persist(testcaseId);

        } catch (Exception e) {
            throw new RuntimeException("Failed to init testcase: " + testcaseId, e);
        }
    }

    // =========================================================
    // ✅ ADD STEP (FULL PAGE SCREENSHOT)
    // =========================================================
    public static synchronized void addStep(WebDriver driver, String testcaseId, String stepDesc) {
        if (driver == null) return;

        try {
            if (!DOCS.containsKey(testcaseId)) {
                initTestcase(testcaseId);
            }

            XWPFDocument doc = DOCS.get(testcaseId);

            // 📸 Capture full-page screenshot
            BufferedImage image = captureFullPage(driver);

            Path imgPath = REPORTS_DIR.resolve(
                    testcaseId + "_" + System.currentTimeMillis() + ".png"
            );
            ImageIO.write(image, "png", imgPath.toFile());

            // 📝 Caption
            XWPFParagraph caption = doc.createParagraph();
            XWPFRun capRun = caption.createRun();
            capRun.setItalic(true);
            capRun.setFontSize(10);
            capRun.setText(TS_FMT.format(new Date()) + " — " + stepDesc);
            capRun.addBreak();

            // 🖼 Embed image
            try (InputStream is = Files.newInputStream(imgPath)) {
                XWPFParagraph picPara = doc.createParagraph();
                picPara.setAlignment(ParagraphAlignment.CENTER);
                XWPFRun picRun = picPara.createRun();
                picRun.addPicture(
                        is,
                        Document.PICTURE_TYPE_PNG,
                        imgPath.getFileName().toString(),
                        Units.toEMU(450),
                        Units.toEMU(300)
                );
            }

            persist(testcaseId);

            System.out.println("📸 Screenshot captured: " + stepDesc);

        } catch (Exception e) {
            throw new RuntimeException("Failed to add step for " + testcaseId, e);
        }
    }

    // =========================================================
    // ✅ CLOSE TESTCASE  (THIS WAS MISSING BEFORE)
    // =========================================================
    public static synchronized void closeTestcase(String testcaseId) {
        try {
            if (!DOCS.containsKey(testcaseId)) return;

            persist(testcaseId);
            DOCS.get(testcaseId).close();

            DOCS.remove(testcaseId);
            DOC_PATHS.remove(testcaseId);

            System.out.println("📄 Testcase document closed: " + testcaseId);

        } catch (Exception e) {
            throw new RuntimeException("Failed to close testcase: " + testcaseId, e);
        }
    }

    // =========================================================
    // 🔧 FULL PAGE SCREENSHOT WITH AUTO SCROLL
    // =========================================================
    private static BufferedImage captureFullPage(WebDriver driver) throws Exception {
        JavascriptExecutor js = (JavascriptExecutor) driver;

        long totalHeight = (long) js.executeScript("return document.body.scrollHeight");
        long viewportHeight = (long) js.executeScript("return window.innerHeight");

        List<BufferedImage> images = new ArrayList<>();

        for (int y = 0; y < totalHeight; y += viewportHeight) {
            js.executeScript("window.scrollTo(0, arguments[0])", y);
            Thread.sleep(300);

            File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
            images.add(ImageIO.read(src));
        }

        int width = images.get(0).getWidth();
        int height = images.stream().mapToInt(BufferedImage::getHeight).sum();

        BufferedImage full = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = full.createGraphics();

        int currentY = 0;
        for (BufferedImage img : images) {
            g.drawImage(img, 0, currentY, null);
            currentY += img.getHeight();
        }
        g.dispose();

        js.executeScript("window.scrollTo(0,0)");
        return full;
    }

    // =========================================================
    // 💾 SAVE DOC
    // =========================================================
    private static void persist(String testcaseId) throws IOException {
        try (OutputStream os = Files.newOutputStream(
                DOC_PATHS.get(testcaseId),
                StandardOpenOption.CREATE,
                StandardOpenOption.TRUNCATE_EXISTING
        )) {
            DOCS.get(testcaseId).write(os);
        }
    }
}
