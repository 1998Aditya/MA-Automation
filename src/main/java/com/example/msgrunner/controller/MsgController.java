package com.example.msgrunner.controller;

import MA_MSG_Suite_INB.*;
import MA_MSG_Suite_OB.*;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.springframework.web.bind.annotation.*;

import java.io.FileInputStream;
import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicBoolean;

@RestController
@RequestMapping("/api")
public class MsgController {

    private final ConcurrentHashMap<String, Boolean> cancelFlags = new ConcurrentHashMap<>(); // outbound cancel
    private final ExecutorService executor = Executors.newSingleThreadExecutor();
    private final ConcurrentHashMap<String, String> jobs = new ConcurrentHashMap<>();

    // Track submitted Futures so we can cancel them
    private final ConcurrentHashMap<String, Future<?>> runningTasks = new ConcurrentHashMap<>();

    // --- Stop store for inbound stop requests (thread-safe) ---
    private static final ConcurrentHashMap<String, AtomicBoolean> stopRequests = new ConcurrentHashMap<>();

    public static void requestStop(String jobId) {
        if (jobId == null) return;
        stopRequests.computeIfAbsent(jobId, k -> new AtomicBoolean(false)).set(true);
    }

    public static boolean isStopRequested(String jobId) {
        if (jobId == null) return false;
        AtomicBoolean b = stopRequests.get(jobId);
        return b != null && b.get();
    }

    public static void clearStopRequest(String jobId) {
        if (jobId == null) return;
        stopRequests.remove(jobId);
    }

    @PostMapping("/run")
    public Map<String, String> runInbound(@RequestBody Map<String, Object> payload) {
        String jobId = UUID.randomUUID().toString();
        jobs.put(jobId, "queued");
        @SuppressWarnings("unchecked")
        List<String> steps = (List<String>) payload.get("steps");
        String env = (String) payload.get("env");
        if (steps == null || steps.isEmpty()) {
            return Map.of("error", "No steps provided");
        }

        // Submit and keep the Future
        Future<?> f = executor.submit(() -> {
            try {
                jobs.put(jobId, "running");
                // clear any previous stop flag just in case
                clearStopRequest(jobId);
                executeInboundSteps(jobId, steps, env);
                // If executeInboundSteps returns normally and stop wasn't requested, set finished
                if (!isStopRequested(jobId)) {
                    jobs.put(jobId, "finished");
                } else {
                    jobs.put(jobId, "stopped");
                }
            } catch (InterruptedException ie) {
                // Thread interrupted (likely via cancel(true))
                jobs.put(jobId, "cancelled_by_interrupt");
                Thread.currentThread().interrupt();
            } catch (Throwable e) {
                jobs.put(jobId, "failed:" + e.getMessage());
                e.printStackTrace();
            } finally {
                // ensure cleanup
                runningTasks.remove(jobId);
                clearStopRequest(jobId);
            }
        });

        runningTasks.put(jobId, f);
        return Map.of("jobId", jobId);
    }

    @PostMapping("/outbound/run")
    public Map<String, String> runOutbound(@RequestBody Map<String, Object> payload) {
        String jobId = UUID.randomUUID().toString();
        jobs.put(jobId, "queued");
        @SuppressWarnings("unchecked")
        List<String> steps = (List<String>) payload.get("steps");
        String env = (String) payload.get("env");
        String filePath = ExcelReaderOB.DATA_EXCEL_PATH;

        Future<?> f = executor.submit(() -> {
            System.out.println("Executor picked up job " + jobId);
            try {
                jobs.put(jobId, "running");
                System.out.println("Calling executeOutboundSteps with steps=" + steps);
                executeOutboundSteps(jobId, filePath, steps, env);
                jobs.put(jobId, "finished");
            } catch (InterruptedException ie) {
                jobs.put(jobId, "cancelled_by_interrupt");
                Thread.currentThread().interrupt();
            } catch (Throwable e) {
                jobs.put(jobId, "failed:" + e.getMessage());
                e.printStackTrace();
            } finally {
                runningTasks.remove(jobId);
            }
        });

        runningTasks.put(jobId, f);
        System.out.println("Received steps: " + steps + ", env: " + env);
        return Map.of("jobId", jobId);
    }

    @GetMapping("/status/{jobId}")
    public Map<String, String> getStatus(@PathVariable String jobId) {
        return Map.of("jobId", jobId, "status", jobs.getOrDefault(jobId, "unknown"));
    }

    /**
     * Legacy/outbound cancel endpoint (unchanged logic), extended to cancel the running future as well.
     * Retained to avoid changing outbound behavior; it still sets cancelFlags used by executeOutboundSteps.
     */
    @PostMapping("/cancel/{jobId}")
    public Map<String, String> cancelJob(@PathVariable String jobId) {
        cancelFlags.put(jobId, true);
        jobs.put(jobId, "cancelled");
        // also attempt to cancel running task/future if present
        Future<?> f = runningTasks.get(jobId);
        if (f != null) {
            boolean ok = f.cancel(true); // interrupt if running
            System.out.println("Attempted to cancel future for job " + jobId + ", result=" + ok);
        }
        return Map.of("status", "cancelled");
    }

    /**
     * New: Stop endpoint for inbound runs. Expects JSON body like: { "jobId": "<id>" }.
     * This sets the stop flag for inbound execution and cancels the future if possible.
     */
    @PostMapping("/stop")
    public Map<String, String> stopJob(@RequestBody Map<String, String> payload) {
        String jobId = payload != null ? payload.get("jobId") : null;
        if (jobId == null || jobId.trim().isEmpty()) {
            return Map.of("error", "jobId required");
        }

        // Set stop flag (task checks this)
        requestStop(jobId);
        jobs.put(jobId, "stopping");

        // Try to cancel the future if it's running
        Future<?> f = runningTasks.get(jobId);
        if (f != null) {
            boolean cancelled = f.cancel(true); // will send interrupt to worker thread
            System.out.println("Stop requested for job " + jobId + ", cancel returned=" + cancelled);
        } else {
            System.out.println("Stop requested for job " + jobId + " but no running future found.");
        }

        return Map.of("status", "stop_requested");
    }

    // --- The executeInboundSteps and executeOutboundSteps methods are preserved except that they throw InterruptedException
    // --- when interrupted and respect isStopRequested(jobId) as before. We only changed orchestration; not step logic.
    private void executeInboundSteps(String jobId, List<String> steps, String env) throws Exception {
        WebDriver driver = null;

        // 🔹 Validators (ADDED – no existing code changed)
        LPNlvlValidation lpnValidator = new LPNlvlValidation();
        ItemlvlValidation itemValidator = new ItemlvlValidation();

        class Helpers {
            private Map<String, List<Row>> groupRows(Sheet sheet) {
                Map<String, List<Row>> map = new LinkedHashMap<>();
                if (sheet == null) return map;
                Iterator<Row> it = sheet.iterator();
                if (it.hasNext()) it.next();
                while (it.hasNext()) {
                    Row row = it.next();
                    if (row == null) continue;
                    Cell c = row.getCell(0);
                    if (c == null) continue;
                    String tc = c.toString().trim();
                    if (!tc.matches("TST_\\d+")) continue;
                    map.computeIfAbsent(tc, x -> new ArrayList<>()).add(row);
                }
                return map;
            }

            private LinkedHashSet<String> collectTestcasesFromWorkbook(String dataExcelPath) {
                LinkedHashSet<String> orderedTcs = new LinkedHashSet<>();
                try (FileInputStream fis = new FileInputStream(new File(dataExcelPath));
                     Workbook wb = WorkbookFactory.create(fis)) {

                    for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                        Sheet sh = wb.getSheetAt(i);
                        Map<String, List<Row>> grouped = groupRows(sh);
                        orderedTcs.addAll(grouped.keySet());
                    }
                } catch (Exception e) {
                    System.out.println("⚠ Unable to collect testcases from workbook: " + e.getMessage());
                }
                return orderedTcs;
            }
        }

        Helpers helpers = new Helpers();

        try {
            if (isStopRequested(jobId)) {
                System.out.println("Stop requested before start for jobId=" + jobId + " — aborting.");
                return;
            }

            LinkedHashSet<String> testcases = helpers.collectTestcasesFromWorkbook(ExcelReaderIB.DATA_EXCEL_PATH);

            if (testcases.isEmpty()) {
                System.out.println("No Testcase entries found — running steps once (old behaviour).");
                for (String step : steps) {

                    if (Thread.currentThread().isInterrupted() || isStopRequested(jobId)) return;

                    switch (step) {
                        case "step1": {
                            WebDriverManager.chromedriver().setup();
                            ChromeOptions o = new ChromeOptions();
                            o.addArguments("--start-maximized");
                            driver = new ChromeDriver(o);
                            driver.manage().window().maximize();
                            URL_Login s1 = new URL_Login(driver, env);
                            s1.execute();
                            break;
                        }
                        case "step2": {
                            Thread.sleep(10000);
                            MSG_LPN_ASN_Creation s2 = new MSG_LPN_ASN_Creation();
                            s2.execute();
                            if (driver != null) {
                                lpnValidator.execute(driver, System.getProperty("testcase"));
                            }
                            break;
                        }
                        case "step3": {
                            Thread.sleep(10000);
                            MSG_Item_ASN_Creation s3 = new MSG_Item_ASN_Creation();
                            s3.execute();
                            if (driver != null) {
                                itemValidator.execute(driver, System.getProperty("testcase"));
                            }
                            break;
                        }
                        case "step4": {
                            Thread.sleep(10000);
                            MSG_ReportRCV_itemlvl s4 = new MSG_ReportRCV_itemlvl();
                            s4.execute();
                            if (driver != null) {
                                itemValidator.execute(driver, System.getProperty("testcase"));
                            }
                            break;
                        }
                        case "step5": {
                            Thread.sleep(10000);
                            MSG_ReportRCV_LPNLvl s5 = new MSG_ReportRCV_LPNLvl();
                            s5.execute();
                            if (driver != null) {
                                lpnValidator.execute(driver, System.getProperty("testcase"));
                            }
                            break;
                        }
                        case "step6": {
                            Thread.sleep(10000);
                            GetConditionCode s6 = new GetConditionCode();
                            s6.execute();
                            break;
                        }
                        case "step7": {
                            Thread.sleep(10000);
                            BoxDelivered s7 = new BoxDelivered();
                            s7.execute();
                            break;
                        }
                        case "step8": {
                            Thread.sleep(10000);
                            RemoveConditionCode s8 = new RemoveConditionCode();
                            s8.execute();
                            break;
                        }
                        case "step9": {
                            if (driver == null) {
                                WebDriverManager.chromedriver().setup();
                                ChromeOptions o2 = new ChromeOptions();
                                o2.addArguments("--start-maximized");
                                driver = new ChromeDriver(o2);
                            }
                            InductiLPN_MFS s9 = new InductiLPN_MFS(driver);
                            s9.execute();
                            break;
                        }
                        case "step10": {
                            Thread.sleep(10000);
                            iLPNToted s10 = new iLPNToted();
                            s10.execute();
                            if (driver != null) {
                                itemValidator.execute(driver, System.getProperty("testcase"));
                                lpnValidator.execute(driver, System.getProperty("testcase"));
                            }
                            break;
                        }
                        case "step11": {
                            Manual_Item_rcv s11 = new Manual_Item_rcv(driver);
                            s11.execute();
                            if (driver != null) {
                                itemValidator.execute(driver, System.getProperty("testcase"));
                            }
                            break;
                        }
                        case "step12": {
                            Manual_LPN_rcv s12 = new Manual_LPN_rcv(driver);
                            s12.execute();
                            if (driver != null) {
                                lpnValidator.execute(driver, System.getProperty("testcase"));
                            }
                            break;
                        }
                        case "step13": {
                            Manual_pallet_putaway s13 = new Manual_pallet_putaway(driver);
                            s13.execute();
                            if (driver != null) {
                                itemValidator.execute(driver, System.getProperty("testcase"));
                                lpnValidator.execute(driver, System.getProperty("testcase"));
                            }
                            break;
                        }
                    }
                }
            } else {
                System.out.println("Detected Testcases: " + testcases + " — running in Testcase sequence.");
                for (String tc : testcases) {

                    System.out.println("\n===== Starting Testcase: " + tc + " =====");
                    System.setProperty("testcase", tc);

                    for (String step : steps) {

                        switch (step) {
                            case "step2": {
                                Thread.sleep(10000);
                                MSG_LPN_ASN_Creation s2 = new MSG_LPN_ASN_Creation();
                                s2.execute();
                                if (driver != null) {
                                    lpnValidator.execute(driver, tc);
                                }
                                break;
                            }
                            case "step3": {
                                Thread.sleep(10000);
                                MSG_Item_ASN_Creation s3 = new MSG_Item_ASN_Creation();
                                s3.execute();
                                if (driver != null) {
                                    itemValidator.execute(driver, tc);
                                }
                                break;
                            }
                            case "step4": {
                                Thread.sleep(10000);
                                MSG_ReportRCV_itemlvl s4 = new MSG_ReportRCV_itemlvl();
                                s4.execute();
                                if (driver != null) {
                                    itemValidator.execute(driver, tc);
                                }
                                break;
                            }
                            case "step5": {
                                Thread.sleep(10000);
                                MSG_ReportRCV_LPNLvl s5 = new MSG_ReportRCV_LPNLvl();
                                s5.execute();
                                if (driver != null) {
                                    lpnValidator.execute(driver, tc);
                                }
                                break;
                            }
                            case "step10": {
                                Thread.sleep(10000);
                                iLPNToted s10 = new iLPNToted();
                                s10.execute();
                                if (driver != null) {
                                    itemValidator.execute(driver, tc);
                                    lpnValidator.execute(driver, tc);
                                }
                                break;
                            }
                            case "step11": {
                                Manual_Item_rcv s11 = new Manual_Item_rcv(driver);
                                s11.execute();
                                if (driver != null) {
                                    itemValidator.execute(driver, tc);
                                }
                                break;
                            }
                            case "step12": {
                                Manual_LPN_rcv s12 = new Manual_LPN_rcv(driver);
                                s12.execute();
                                if (driver != null) {
                                    lpnValidator.execute(driver, tc);
                                }
                                break;
                            }
                            case "step13": {
                                Manual_pallet_putaway s13 = new Manual_pallet_putaway(driver);
                                s13.execute();
                                if (driver != null) {
                                    itemValidator.execute(driver, tc);
                                    lpnValidator.execute(driver, tc);
                                }
                                break;
                            }
                        }
                    }

                    System.clearProperty("testcase");
                    System.out.println("===== Finished Testcase: " + tc + " =====\n");
                }
            }
        } finally {
            clearStopRequest(jobId);
        }
    }
    // executeOutboundSteps left mostly unchanged; we still honor cancelFlags and also support future.cancel
    private void executeOutboundSteps(String jobId, String filePath, List<String> steps, String env) throws Exception {
        System.out.println("Executing outbound steps for job " + jobId);

        if (steps.contains("Auto")) {
            System.out.println("From excel we are taking");

            FileInputStream fis = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet finalSheet = workbook.getSheet("Final");

            int tcIdx = -1, fromIdx = -1, orderIdx = -1, uptoIdx = -1;
            Row header = finalSheet.getRow(0);
            for (Cell cell : header) {
                String val = cell.getStringCellValue().trim();
                if (val.equalsIgnoreCase("TestCase")) tcIdx = cell.getColumnIndex();
                if (val.equalsIgnoreCase("From")) fromIdx = cell.getColumnIndex();
                if (val.equalsIgnoreCase("Order")) orderIdx = cell.getColumnIndex();
                if (val.equalsIgnoreCase("Upto")) uptoIdx = cell.getColumnIndex();
            }

            for (int i = 1; i <= finalSheet.getLastRowNum(); i++) {
                Row row = finalSheet.getRow(i);
                if (row == null) continue;

                String tcVal = getCellValue(row, tcIdx);
                String fromVal = getCellValue(row, fromIdx);
                String orderVal = getCellValue(row, orderIdx);
                String uptoVal = getCellValue(row, uptoIdx);

                List<String> currentSteps = new ArrayList<>();
                if (fromVal.equalsIgnoreCase("userInventory")) currentSteps.add("stepOB2");
                else if (fromVal.equalsIgnoreCase("systemInventory")) currentSteps.add("stepOB3");
                else if (fromVal.equalsIgnoreCase("B2B")) currentSteps.add("stepOB4");
                else if (fromVal.equalsIgnoreCase("B2C")) currentSteps.add("stepOB5");
                else if (fromVal.equalsIgnoreCase("wave") && uptoVal.equalsIgnoreCase("pack")) {
                    currentSteps.add("stepOB6"); currentSteps.add("stepOB7"); currentSteps.add("stepOB8");
                } else if (fromVal.equalsIgnoreCase("pick") && uptoVal.equalsIgnoreCase("pack")) {
                    currentSteps.add("stepOB7"); currentSteps.add("stepOB8");
                }

                if (orderVal.equalsIgnoreCase("B2B")) currentSteps.add("stepOB4");
                else if (orderVal.equalsIgnoreCase("B2C")) currentSteps.add("stepOB5");

                if (uptoVal.equalsIgnoreCase("wave")) currentSteps.add("stepOB6");
                else if (uptoVal.equalsIgnoreCase("pick")) { currentSteps.add("stepOB6"); currentSteps.add("stepOB7"); }
                else if (uptoVal.equalsIgnoreCase("pack")) { currentSteps.add("stepOB6"); currentSteps.add("stepOB7"); currentSteps.add("stepOB8"); }

                currentSteps = new ArrayList<>(new LinkedHashSet<>(currentSteps));

                System.out.println("Executing " + tcVal + " with steps: " + currentSteps);
                FileInputStream fis2 = new FileInputStream(filePath);
                Workbook workbook2 = new XSSFWorkbook(fis2);

                Sheet sheet1 = workbook2.getSheet("Data Creation");
                Sheet sheet2 = workbook2.getSheet("Order Creation");
                Sheet sheet3 = workbook2.getSheet("OPS_Tab");
                Sheet sheet4 = workbook2.getSheet("Tasks");
                Sheet sheet5 = workbook2.getSheet("ECOM Order");
                Sheet sheet6 = workbook2.getSheet("Outbound");
                Map<String, List<Row>> d1 = groupRows(sheet1);
                Map<String, List<Row>> d2 = groupRows(sheet2);
                Map<String, List<Row>> d3 = groupRows(sheet3);
                Map<String, List<Row>> d4 = groupRows(sheet4);
                Map<String, List<Row>> d5 = groupRows(sheet5);
                Map<String, List<Row>> d6 = groupRows(sheet6);

                if (currentSteps.contains("stepOB2")) {
                    Main2_CreateInventory.execute(tcVal, filePath, "Data Creation", env);
                }
                if (currentSteps.contains("stepOB3")) {
                    Main3_UserDirectedPutaway.execute(tcVal, filePath, "Data Creation", env);
                }
                if (currentSteps.contains("stepOB4")) {
                    Main4_B2BOrderCreation.main(tcVal, filePath,env);
                }
                if (currentSteps.contains("stepOB5")) {
                    Main5_EcomOrderCreation.main(tcVal, filePath,env);
                }
                if (currentSteps.contains("stepOB6")) {
                    Main6a_WaveToReleaseTask.main(filePath, tcVal, env);
                }
                if (currentSteps.contains("stepOB7")) {
                    Main7_MakePickCart.main(filePath, tcVal, env, "Tasks");
                }
                if (currentSteps.contains("stepOB8")) {
                    Main8_ECOMPackPickCart.main(filePath, tcVal, env);
                }

                workbook2.close();
                fis2.close();
            }

            workbook.close();
            fis.close();
        }

        if (steps.contains("stepOB1")) {
            WebDriverManager.chromedriver().setup();
            ChromeOptions o = new ChromeOptions();
            o.addArguments("--start-maximized");
            WebDriver driver = new ChromeDriver(o);
            driver.manage().window().maximize();
            Main1_URL_Login1 s = new Main1_URL_Login1(driver, env);
            s.execute();
            System.out.println("Outbound Login Done");
            return;
        }

        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet1 = workbook.getSheet("Data Creation");
        Sheet sheet2 = workbook.getSheet("Order Creation");
        Sheet sheet3 = workbook.getSheet("OPS_Tab");
        Sheet sheet4 = workbook.getSheet("Tasks");
        Sheet sheet5 = workbook.getSheet("ECOM Order");
        Sheet sheet6 = workbook.getSheet("Outbound");
        Map<String, List<Row>> d1 = groupRows(sheet1);
        Map<String, List<Row>> d2 = groupRows(sheet2);
        Map<String, List<Row>> d3 = groupRows(sheet3);
        Map<String, List<Row>> d4 = groupRows(sheet4);
        Map<String, List<Row>> d5 = groupRows(sheet5);
        Map<String, List<Row>> d6 = groupRows(sheet6);
        Set<String> cases = new LinkedHashSet<>();
        cases.addAll(d1.keySet());
        cases.addAll(d2.keySet());
        cases.addAll(d3.keySet());
        cases.addAll(d4.keySet());
        cases.addAll(d5.keySet());
        cases.addAll(d6.keySet());

        // Reset DocPathManager state at the start of this run
        DocPathManager.reset();

        for (String tc : cases) {
            if (cancelFlags.getOrDefault(jobId, false) || Thread.currentThread().isInterrupted()) {
                System.out.println("Job " + jobId + " cancelled.");
                return;
            }

            if (steps.contains("stepOB2")) {
                if (d1.get(tc) != null) {
                    Main2_CreateInventory.execute(tc, filePath, "Data Creation", env);
                }
            }
            if (steps.contains("stepOB3")) {
                if (d1.get(tc) != null) {
                    Main3_UserDirectedPutaway.execute(tc, filePath, "Data Creation", env);
                }
            }
            if (steps.contains("stepOB4")) {
                if (d2.get(tc) != null) {
                    Main4_B2BOrderCreation.main(tc, filePath,env);
                    System.out.println("------B2BOrderCreation complete" + tc);
                }
            }
            if (steps.contains("stepOB5")) {
                if (d5.get(tc) != null) {
                    Main5_EcomOrderCreation.main(tc, filePath,env);
                    System.out.println("------EcomOrderCreation complete" + tc);
                }
            }
            if (steps.contains("stepOB6")) {
                if (d3.get(tc) != null) {
                    Main6a_WaveToReleaseTask.main(filePath, tc, env);
                }
            }
            if (steps.contains("stepOB7")) {
                Map<String, List<Row>> e4 = loadSheet(filePath, "Tasks"); // reload fresh
                if (e4.get(tc) != null) {
                    Main7_MakePickCart.main(filePath, tc, env, "Tasks");
                    System.out.println("---------------MakePickCart complete" + tc);
                }
            }

            if (steps.contains("stepOB8")) {
                Map<String, List<Row>> e5 = loadSheet(filePath, "Tasks"); // reload fresh
                if (e5.get(tc) != null) {
                    Main8_B2BPackPickCart.main(filePath, tc, env);
                    System.out.println("---------------B2B Packing complete" + tc);
                }
            }
            if (steps.contains("stepOB9")) {
                Map<String, List<Row>> e6 = loadSheet(filePath, "Tasks"); // reload fresh
                if (e6.get(tc) != null) {
                    Main8_ECOMPackPickCart.main(filePath, tc, env);
                    System.out.println("---------------ECOM Packing complete" + tc);
                }
            }
            if (steps.contains("stepOB10")) {
                Map<String, List<Row>> e7 = loadSheet(filePath, "Tasks"); // reload fresh
                if (e7.get(tc) != null) {
                    System.out.println("Manual Palletisation started for "+tc);
                    Main9_Palletisation.main(filePath, tc, env);
                    System.out.println("---------------Manual Palletisation complete" + tc);
                    Thread.sleep(800);
                }}
            if (steps.contains("stepOB11")) {
                if (d6.get(tc) != null) {
                    Main10_OutboundPutaway.main(filePath, tc, env);
                    Thread.sleep(800);
                }
            }
            if (steps.contains("stepOB12")) {
                if (d6.get(tc) != null) {
                    Main11_Shipment.main(filePath, "Outbound", tc, env);
                    Thread.sleep(800);
                }
            }
            if (steps.contains("stepOB5A")) {
                Map<String, List<Row>> e8 = loadSheet(filePath, "Tasks"); // reload fresh
                if (e8.get(tc) != null) {
                    System.out.println("NewOutboundOrder Started");
                    String docPathLocal = DocPathManager.getOrCreateDocPath(filePath, tc);
                    Main100_MHEJournalScreenshot.main(tc,filePath, env,"NewOutboundOrder",docPathLocal);
                    Thread.sleep(800);
                }}

            if (steps.contains("stepOB1A")) {
                Map<String, List<Row>> e8 = loadSheet(filePath, "Tasks"); // reload fresh
                if (e8.get(tc) != null) {
                    System.out.println("OrderReadyForPacking started");
                    MainA1_OrderReadyForPacking.main(tc,filePath, "OrderReadyForPacking",env);
                    Thread.sleep(800);
                }
            }
            if (steps.contains("stepOB2A")) {
                Map<String, List<Row>> e8 = loadSheet(filePath, "Tasks"); // reload fresh
                if (e8.get(tc) != null) {
                    System.out.println("oLPNPrepared Started");
                    MainA2_oLPNPrepared.main(tc,filePath, "oLPNPrepared",env);
                    Thread.sleep(800);
                }}
            if (steps.contains("stepOB3A")) {
                Map<String, List<Row>> e8 = loadSheet(filePath, "Tasks"); // reload fresh
                if (e8.get(tc) != null) {
                    System.out.println("oLPNPalletised: Not implemented yet");
                    MainA3_oLPNPalletised.main(filePath,tc, "oLPNPalletised",env);
                    Thread.sleep(800);
                }}
            if (steps.contains("stepOB4A")) {
                Map<String, List<Row>> e8 = loadSheet(filePath, "Tasks"); // reload fresh
                if (e8.get(tc) != null) {
                    System.out.println("PalletReady: Not implemented yet");
                    MainA4_PalletReady.main(filePath,tc, "PalletReady",env);
                    Thread.sleep(800);
                }}


        }

        workbook.close();
        fis.close();
    }

    private Map<String, List<Row>> groupRows(Sheet sheet) {
        Map<String, List<Row>> map = new LinkedHashMap<>();
        if (sheet == null) return map;
        Iterator<Row> it = sheet.iterator();
        if (it.hasNext()) it.next();
        while (it.hasNext()) {
            Row row = it.next();
            if (row == null) continue;
            Cell c = row.getCell(0);
            if (c == null) continue;
            String tc = c.toString().trim();
            if (!tc.matches("TST_\\d+")) continue;
            map.computeIfAbsent(tc, x -> new ArrayList<>()).add(row);
        }
        return map;
    }

    private String getCellValue(Row row, int index) {
        if (row == null || index < 0) return "";
        Cell cell = row.getCell(index);
        if (cell == null) return "";
        cell.setCellType(CellType.STRING);
        return cell.getStringCellValue().trim();
    }

    private Map<String, List<Row>> loadSheet(String filePath, String sheetName) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet(sheetName);
            return groupRows(sheet);
        }
    }
}
