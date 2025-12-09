//package com.example.msgrunner.controller;
//
//import MA_MSG_Suite_INB.*;
//import MA_MSG_Suite_OB.*;
//import io.github.bonigarcia.wdm.WebDriverManager;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.openqa.selenium.WebDriver;
//import org.openqa.selenium.chrome.ChromeDriver;
//import org.openqa.selenium.chrome.ChromeOptions;
//import org.springframework.web.bind.annotation.*;
//import java.io.FileInputStream;
//import java.util.*;
//import java.util.concurrent.*;
//
//@RestController
//@RequestMapping("/api")
//public class MsgController {
//
//    private final ConcurrentHashMap<String,Boolean> cancelFlags=new ConcurrentHashMap<>();
//    private final ExecutorService executor=Executors.newSingleThreadExecutor();
//    private final ConcurrentHashMap<String,String> jobs=new ConcurrentHashMap<>();
//
//    @PostMapping("/run")
//    public Map<String,String> runInbound(@RequestBody Map<String,Object> payload){
//        String jobId=UUID.randomUUID().toString();
//        jobs.put(jobId,"queued");
//        @SuppressWarnings("unchecked")
//        List<String> steps=(List<String>)payload.get("steps");
//        String env=(String)payload.get("env");
//        if(steps==null||steps.isEmpty()){
//            return Map.of("error","No steps provided");
//        }
//        executor.submit(()->{
//            try{
//                jobs.put(jobId,"running");
//                executeInboundSteps(jobId,steps,env);
//                jobs.put(jobId,"finished");
//            }catch(Throwable e){
//                jobs.put(jobId,"failed:"+e.getMessage());
//                e.printStackTrace();
//            }
//        });
//        return Map.of("jobId",jobId);
//    }
//
//    @PostMapping("/outbound/run")
//    public Map<String,String> runOutbound(@RequestBody Map<String,Object> payload){
//        String jobId=UUID.randomUUID().toString();
//        jobs.put(jobId,"queued");
//        System.out.println("Executor picked up job " + jobId);
//        @SuppressWarnings("unchecked")
//        List<String> steps=(List<String>)payload.get("steps");
//        String env=(String)payload.get("env");
//        String filePath="C:\\Users\\2210420\\IdeaProjects\\msg-runner\\OOdata.xlsx";
//        executor.submit(()->{
//            try{
//                jobs.put(jobId,"running");
//                executeOutboundSteps(jobId,filePath,steps,env);
//                jobs.put(jobId,"finished");
//            }catch(Throwable e){
//                jobs.put(jobId,"failed:"+e.getMessage());
//                e.printStackTrace();
//            }
//        });
//        return Map.of("jobId",jobId);
//    }
//
//    @GetMapping("/status/{jobId}")
//    public Map<String,String> getStatus(@PathVariable String jobId){
//        return Map.of("jobId",jobId,"status",jobs.getOrDefault(jobId,"unknown"));
//    }
//
//    @PostMapping("/cancel/{jobId}")
//    public Map<String,String> cancelJob(@PathVariable String jobId){
//        cancelFlags.put(jobId,true);
//        jobs.put(jobId,"cancelled");
//        return Map.of("status","cancelled");
//    }
//
//    private void executeInboundSteps(String jobId,List<String> steps,String env)throws Exception{
//        WebDriver driver=null;
//        try{
//            for(String step:steps){
//                switch(step){
//
//                    case "step1":{
//                        WebDriverManager.chromedriver().setup();
//                        ChromeOptions o=new ChromeOptions();
//                        o.addArguments("--start-maximized");
//                        driver=new ChromeDriver(o);
//                        driver.manage().window().maximize();
//                        URL_Login s1=new URL_Login(driver,env);
//                        s1.execute();
//                        break;
//                    }
//
//                    case "step3":{
//                        MSG_Item_ASN_Creation s3=new MSG_Item_ASN_Creation();
//                        s3.execute();
//                        break;
//                    }
//
//                    case "step4":{
//                        MSG_ReportRCV_itemlvl s4=new MSG_ReportRCV_itemlvl();
//                        s4.execute();
//                        break;
//                    }
//
//                    case "step5":{
//                        MSG_ReportRCV_LPNLvl s5=new MSG_ReportRCV_LPNLvl();
//                        s5.execute();
//                        break;
//                    }
//
//                    case "step6":{
//                        GetConditionCode s6=new GetConditionCode();
//                        s6.execute();
//                        break;
//                    }
//
//                    case "step7":{
//                        BoxDelivered s7=new BoxDelivered();
//                        s7.execute();
//                        break;
//                    }
//
//                    case "step8":{
//                        RemoveConditionCode s8=new RemoveConditionCode();
//                        s8.execute();
//                        break;
//                    }
//
//                    case "step9":{
//                        if(driver==null){
//                            WebDriverManager.chromedriver().setup();
//                            ChromeOptions o2=new ChromeOptions();
//                            o2.addArguments("--start-maximized");
//                            driver=new ChromeDriver(o2);
//                        }
//                        InductiLPN_MFS s9=new InductiLPN_MFS(driver);
//                        s9.execute();
//                        break;
//                    }
//
//                    case "step10":{
//                        iLPNToted s10=new iLPNToted();
//                        s10.execute();
//                        break;
//                    }
//
//                    default:
//                        System.out.println("Unknown inbound step:"+step);
//                }
//                Thread.sleep(1000);
//            }
//        }finally{
//            if(driver!=null){
//                driver.quit();
//                System.out.println("ChromeDriver closed.");
//            }
//        }
//    }
//
//    private void executeOutboundSteps(String jobId,String filePath,List<String> steps,String env)throws Exception{
//
//
//
//        if(steps.contains("stepOB1")){
//            WebDriverManager.chromedriver().setup();
//            ChromeOptions o=new ChromeOptions();
//            o.addArguments("--start-maximized");
//            WebDriver driver=new ChromeDriver(o);
//            driver.manage().window().maximize();
//            Main1_URL_Login1 s=new Main1_URL_Login1(driver,env);
//            s.execute();
//            System.out.println("Outbound Login Done");
//            return;
//        }
//
//        FileInputStream fis=new FileInputStream(filePath);
//        Workbook workbook=new XSSFWorkbook(fis);
//
//        Sheet sheet1=workbook.getSheet("Data Creation");
//        Sheet sheet2=workbook.getSheet("Order Creation");
//        Sheet sheet3=workbook.getSheet("OPS_Tab");
//        Sheet sheet4=workbook.getSheet("Tasks");
//        Sheet sheet5=workbook.getSheet("ECOM Order");
//        Sheet sheet6=workbook.getSheet("Outbound");
//
//        Map<String,List<Row>> d1=groupRows(sheet1);
//        Map<String,List<Row>> d2=groupRows(sheet2);
//        Map<String,List<Row>> d3=groupRows(sheet3);
//        Map<String,List<Row>> d4=groupRows(sheet4);
//        Map<String,List<Row>> d5=groupRows(sheet5);
//        Map<String,List<Row>> d6=groupRows(sheet6);
//
//        Set<String> cases=new LinkedHashSet<>();
//        cases.addAll(d1.keySet());
//        cases.addAll(d2.keySet());
//        cases.addAll(d3.keySet());
//        cases.addAll(d4.keySet());
//        cases.addAll(d5.keySet());
//        cases.addAll(d6.keySet());
//
//        for(String tc:cases){
//
//                if(cancelFlags.getOrDefault(jobId,false)){
//                    System.out.println("Job "+jobId+" cancelled.");
//                    return;
//                }
//
//            if(steps.contains("stepOB2")){
//                List<Row> rows=d1.get(tc);
//                if(rows!=null){
//                    for(Row r:rows){
//                        Main2_CreateInventory.execute(r,filePath,"Data Creation",env);
//                    }
//                }
//            }
//            if(steps.contains("stepOB3")){
//                List<Row> rows=d1.get(tc);
//                if(rows!=null){
//                    for(Row r:rows){
//                        Main3_UserDirectedPutaway.execute(r,filePath,"Data Creation",env);
//                    }
//                }
//            }
//
//            if(steps.contains("stepOB4")){
//                if(d3.get(tc)!=null){
//                        Main4_B2BOrderCreation.main(tc,filePath);
//
//                }
//            }
//
//            if(steps.contains("stepOB5")){
//                if(d5.get(tc)!=null) {
//                    Main5_EcomOrderCreation.main(tc, filePath);
//                }
//            }
//
//            if(steps.contains("stepOB6")){
//                if(d3.get(tc)!=null){
//                    Main6a_WaveToReleaseTask.main(filePath,tc,env);
//                }
//            }
//
//            if(steps.contains("stepOB7")){
//                if(d3.get(tc)!=null){
//                    Main7_MakePickCart.main(filePath,tc,env);
//                }
//            }
//
//            if(steps.contains("stepOB8")){
//                if(d3.get(tc)!=null){
//                    Main8_Packing.main(filePath,tc,env);
//                }
//            }
//
//            if(steps.contains("stepOB9")){
//                System.out.println("Palletisation: Not implemented yet");
//                Thread.sleep(800);
//            }
//
//            if(steps.contains("stepOB10")){
//                if (d6.get(tc) != null) {
//
//                    Main10_OutboundPutaway.main(filePath, tc, env);
//                    Thread.sleep(800);
//                }
//            }
//
//            if(steps.contains("stepOB11")) {
//                if (d6.get(tc) != null) {
//                    Main11_Shipment.main(filePath, "Outbound", tc,env);
//                    Thread.sleep(800);
//                }
//            }
//
//        }
//
//        workbook.close();
//        fis.close();
//    }
//
//    private Map<String,List<Row>> groupRows(Sheet sheet){
//        Map<String,List<Row>> map=new LinkedHashMap<>();
//        if(sheet==null)return map;
//        Iterator<Row> it=sheet.iterator();
//        if(it.hasNext())it.next();
//        while(it.hasNext()){
//            Row row=it.next();
//            if(row==null)continue;
//            Cell c=row.getCell(0);
//            if(c==null)continue;
//            String tc=c.toString().trim();
//            if(!tc.matches("TST_\\d+"))continue;
//            map.computeIfAbsent(tc,x->new ArrayList<>()).add(row);
//        }
//        return map;
//    }
//
//}






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
import java.util.*;
import java.util.concurrent.*;
@RestController
@RequestMapping("/api")
public class MsgController {
    private final ConcurrentHashMap<String,Boolean> cancelFlags=new ConcurrentHashMap<>();
    private final ExecutorService executor=Executors.newSingleThreadExecutor();
    private final ConcurrentHashMap<String,String> jobs=new ConcurrentHashMap<>();
    @PostMapping("/run")
    public Map<String,String> runInbound(@RequestBody Map<String,Object> payload){
        String jobId=UUID.randomUUID().toString();
        jobs.put(jobId,"queued");
        @SuppressWarnings("unchecked")
        List<String> steps=(List<String>)payload.get("steps");
        String env=(String)payload.get("env");
        if(steps==null||steps.isEmpty()){
            return Map.of("error","No steps provided");
        }
        executor.submit(()->{
            try{
                jobs.put(jobId,"running");
                executeInboundSteps(jobId,steps,env);
                jobs.put(jobId,"finished");
            }catch(Throwable e){
                jobs.put(jobId,"failed:"+e.getMessage());
                e.printStackTrace();
            }
        });
        return Map.of("jobId",jobId);
    }
    @PostMapping("/outbound/run")
    public Map<String,String> runOutbound(@RequestBody Map<String,Object> payload){
        String jobId=UUID.randomUUID().toString();
        jobs.put(jobId,"queued");
        @SuppressWarnings("unchecked")
        List<String> steps=(List<String>)payload.get("steps");
        String env=(String)payload.get("env");
        String filePath="C:\\Users\\2210420\\IdeaProjects\\msg-runner\\OOdata.xlsx";
//
// executor.submit(()->{
// try{
// jobs.put(jobId,"running");
// executeOutboundSteps(jobId,filePath,steps,env);
// jobs.put(jobId,"finished");
// }catch(Throwable e){
// jobs.put(jobId,"failed:"+e.getMessage());
// e.printStackTrace();
// }
// });
        System.out.println("Executor isShutdown=" + executor.isShutdown() + ", isTerminated=" + executor.isTerminated());
        executor.submit(() -> {
            System.out.println("Executor picked up job " + jobId);
            try {
                jobs.put(jobId,"running");
                System.out.println("Calling executeOutboundSteps with steps=" + steps);
                executeOutboundSteps(jobId,filePath,steps,env);
                jobs.put(jobId,"finished");
            } catch(Throwable e) {
                jobs.put(jobId,"failed:"+e.getMessage());
                e.printStackTrace();
            }
        });
        System.out.println("Received steps: " + steps + ", env: " + env);
        return Map.of("jobId",jobId);
    }
    @GetMapping("/status/{jobId}")
    public Map<String,String> getStatus(@PathVariable String jobId){
        return Map.of("jobId",jobId,"status",jobs.getOrDefault(jobId,"unknown"));
    }
    @PostMapping("/cancel/{jobId}")
    public Map<String,String> cancelJob(@PathVariable String jobId){
        cancelFlags.put(jobId,true);
        jobs.put(jobId,"cancelled");
        return Map.of("status","cancelled");
    }
    private void executeInboundSteps(String jobId,List<String> steps,String env)throws Exception{
        WebDriver driver=null;
        try{
            for(String step:steps){
                switch(step){
                    case "step1":{
                        WebDriverManager.chromedriver().setup();
                        ChromeOptions o=new ChromeOptions();
                        o.addArguments("--start-maximized");
                        driver=new ChromeDriver(o);
                        driver.manage().window().maximize();
                        URL_Login s1=new URL_Login(driver,env);
                        s1.execute();
                        break;
                    }
                    case "step3":{
                        MSG_Item_ASN_Creation s3=new MSG_Item_ASN_Creation();
                        s3.execute();
                        break;
                    }
                    case "step4":{
                        MSG_ReportRCV_itemlvl s4=new MSG_ReportRCV_itemlvl();
                        s4.execute();
                        break;
                    }
                    case "step5":{
                        MSG_ReportRCV_LPNLvl s5=new MSG_ReportRCV_LPNLvl();
                        s5.execute();
                        break;
                    }
                    case "step6":{
                        GetConditionCode s6=new GetConditionCode();
                        s6.execute();
                        break;
                    }
                    case "step7":{
                        BoxDelivered s7=new BoxDelivered();
                        s7.execute();
                        break;
                    }
                    case "step8":{
                        RemoveConditionCode s8=new RemoveConditionCode();
                        s8.execute();
                        break;
                    }
                    case "step9":{
                        if(driver==null){
                            WebDriverManager.chromedriver().setup();
                            ChromeOptions o2=new ChromeOptions();
                            o2.addArguments("--start-maximized");
                            driver=new ChromeDriver(o2);
                        }
                        InductiLPN_MFS s9=new InductiLPN_MFS(driver);
                        s9.execute();
                        break;
                    }
                    case "step10":{
                        iLPNToted s10=new iLPNToted();
                        s10.execute();
                        break;
                    }
                    default:
                        System.out.println("Unknown inbound step:"+step);
                }
                Thread.sleep(1000);
            }
        }finally{
            if(driver!=null){
                driver.quit();
                System.out.println("ChromeDriver closed.");
            }
        }
    }
    private void executeOutboundSteps(String jobId,String filePath,List<String> steps,String env)throws Exception{
        System.out.println("Executing outbound steps for job " + jobId);

        if (steps.contains("Auto")) {
            System.out.println("From excel we are taking");

            FileInputStream fis = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet finalSheet = workbook.getSheet("Final");

            // Find column indexes
            int tcIdx = -1, fromIdx = -1, orderIdx = -1, uptoIdx = -1;
            Row header = finalSheet.getRow(0);
            for (Cell cell : header) {
                String val = cell.getStringCellValue().trim();
                if (val.equalsIgnoreCase("TestCase")) tcIdx = cell.getColumnIndex();
                if (val.equalsIgnoreCase("From")) fromIdx = cell.getColumnIndex();
                if (val.equalsIgnoreCase("Order")) orderIdx = cell.getColumnIndex();
                if (val.equalsIgnoreCase("Upto")) uptoIdx = cell.getColumnIndex();
            }

            // Iterate rows
            for (int i = 1; i <= finalSheet.getLastRowNum(); i++) {
                Row row = finalSheet.getRow(i);
                if (row == null) continue;

                String tcVal   = getCellValue(row, tcIdx);
                String fromVal = getCellValue(row, fromIdx);
                String orderVal= getCellValue(row, orderIdx);
                String uptoVal = getCellValue(row, uptoIdx);

                // Build steps for this testcase
                List<String> currentSteps = new ArrayList<>();

                // --- From column logic ---
                if (fromVal.equalsIgnoreCase("userInventory")) {
                    currentSteps.add("stepOB2");
                } else if (fromVal.equalsIgnoreCase("systemInventory")) {
                    currentSteps.add("stepOB3");
                } else if (fromVal.equalsIgnoreCase("B2B")) {
                    currentSteps.add("stepOB4");
                } else if (fromVal.equalsIgnoreCase("B2C")) {
                    currentSteps.add("stepOB5");
                } else if (fromVal.equalsIgnoreCase("wave") && uptoVal.equalsIgnoreCase("pack")) {
                    currentSteps.add("stepOB6");
                    currentSteps.add("stepOB7");
                    currentSteps.add("stepOB8");
                } else if (fromVal.equalsIgnoreCase("pick") && uptoVal.equalsIgnoreCase("pack")) {
                    currentSteps.add("stepOB7");
                    currentSteps.add("stepOB8");
                }

                // --- Order column logic ---
                if (orderVal.equalsIgnoreCase("B2B")) {
                    currentSteps.add("stepOB4");
                } else if (orderVal.equalsIgnoreCase("B2C")) {
                    currentSteps.add("stepOB5");
                }

                // --- Upto column logic ---
                if (uptoVal.equalsIgnoreCase("wave")) {
                    currentSteps.add("stepOB6");
                } else if (uptoVal.equalsIgnoreCase("pick")) {
                    currentSteps.add("stepOB6");
                    currentSteps.add("stepOB7");
                } else if (uptoVal.equalsIgnoreCase("pack")) {
                    currentSteps.add("stepOB6");
                    currentSteps.add("stepOB7");
                    currentSteps.add("stepOB8");
                }

                // Deduplicate while preserving order
                currentSteps = new ArrayList<>(new LinkedHashSet<>(currentSteps));

                // Now execute the classes for this testcase
                System.out.println("Executing " + tcVal + " with steps: " + currentSteps);
                FileInputStream fis2 = new FileInputStream(filePath);
                Workbook workbook2 = new XSSFWorkbook(fis2);

                Sheet sheet1=workbook2.getSheet("Data Creation");
                Sheet sheet2=workbook2.getSheet("Order Creation");
                Sheet sheet3=workbook2.getSheet("OPS_Tab");
                Sheet sheet4=workbook2.getSheet("Tasks");
                Sheet sheet5=workbook2.getSheet("ECOM Order");
                Sheet sheet6=workbook2.getSheet("Outbound");
                Map<String,List<Row>> d1=groupRows(sheet1);
                Map<String,List<Row>> d2=groupRows(sheet2);
                Map<String,List<Row>> d3=groupRows(sheet3);
                Map<String,List<Row>> d4=groupRows(sheet4);
                Map<String,List<Row>> d5=groupRows(sheet5);
                Map<String,List<Row>> d6=groupRows(sheet6);
                Set<String> cases=new LinkedHashSet<>();
                cases.addAll(d1.keySet());
                cases.addAll(d2.keySet());
                cases.addAll(d3.keySet());
                cases.addAll(d4.keySet());
                cases.addAll(d5.keySet());
                cases.addAll(d6.keySet());


                System.out.println("Executing2 " + tcVal + " with steps: " + currentSteps);


                    // Replace your old "steps.contains(...)" checks with "step.equals(...)"
                    if (currentSteps.contains("stepOB2")) {
                        System.out.println("Main2_CreateInventory ");

                            //    Main2_CreateInventory.execute(r, filePath, "Data Creation", env);


                    }
                System.out.println("Executing3 " + tcVal + " with steps: " + currentSteps);

                if (currentSteps.contains("stepOB3")) {
                    System.out.println("Main2_CreateInventory ");

                       Main3_UserDirectedPutaway.execute(tcVal, filePath, "Data Creation", env);


                    }
                System.out.println("Executing4 " + tcVal + " with steps: " + currentSteps);

                if (currentSteps.contains("stepOB4")) {

                            Main4_B2BOrderCreation.main(tcVal, filePath);
                            System.out.println("------B2BOrderCreation complete"+tcVal);

                    }
                    if (currentSteps.contains("stepOB5")) {

                            Main5_EcomOrderCreation.main(tcVal, filePath);
                            System.out.println("------EcomOrderCreation complete"+tcVal);

                    }
                    if (currentSteps.contains("stepOB6")) {

                            Main6a_WaveToReleaseTask.main(filePath, tcVal, env);

                    }
                    if (currentSteps.contains("stepOB7")) {

                            Main7_MakePickCart.main(filePath, tcVal, env);

                    }

                    if (currentSteps.contains("stepOB8")) {

                            Main8_Packing.main(filePath, tcVal, env);

                    }

            }

            workbook.close();
            fis.close();
        }






        if(steps.contains("stepOB1")){
            WebDriverManager.chromedriver().setup();
            ChromeOptions o=new ChromeOptions();
            o.addArguments("--start-maximized");
            WebDriver driver=new ChromeDriver(o);
            driver.manage().window().maximize();
            Main1_URL_Login1 s=new Main1_URL_Login1(driver,env);
            s.execute();
            System.out.println("Outbound Login Done");
            return;
        }
        FileInputStream fis=new FileInputStream(filePath);
        Workbook workbook=new XSSFWorkbook(fis);
        Sheet sheet1=workbook.getSheet("Data Creation");
        Sheet sheet2=workbook.getSheet("Order Creation");
        Sheet sheet3=workbook.getSheet("OPS_Tab");
        Sheet sheet4=workbook.getSheet("Tasks");
        Sheet sheet5=workbook.getSheet("ECOM Order");
        Sheet sheet6=workbook.getSheet("Outbound");
        Map<String,List<Row>> d1=groupRows(sheet1);
        Map<String,List<Row>> d2=groupRows(sheet2);
        Map<String,List<Row>> d3=groupRows(sheet3);
        Map<String,List<Row>> d4=groupRows(sheet4);
        Map<String,List<Row>> d5=groupRows(sheet5);
        Map<String,List<Row>> d6=groupRows(sheet6);
        Set<String> cases=new LinkedHashSet<>();
        cases.addAll(d1.keySet());
        cases.addAll(d2.keySet());
        cases.addAll(d3.keySet());
        cases.addAll(d4.keySet());
        cases.addAll(d5.keySet());
        cases.addAll(d6.keySet());
        for(String tc:cases) {
            if (cancelFlags.getOrDefault(jobId, false)) {
                System.out.println("Job " + jobId + " cancelled.");
                return;
            }



            if (steps.contains("stepOB2")) {
                List<Row> rows = d1.get(tc);
                if (rows != null) {
                    for (Row r : rows) {
                        Main2_CreateInventory.execute(r, filePath, "Data Creation", env);
                    }
                }
            }
            if (steps.contains("stepOB3")) {
                if (d2.get(tc) != null) {
                        Main3_UserDirectedPutaway.execute(tc, filePath, "Data Creation", env);

                }
            }
            if (steps.contains("stepOB4")) {
                if (d2.get(tc) != null) {
                    Main4_B2BOrderCreation.main(tc, filePath);
                    System.out.println("------B2BOrderCreation complete"+tc);
                }
            }
            if (steps.contains("stepOB5")) {
                if (d5.get(tc) != null) {
                    Main5_EcomOrderCreation.main(tc, filePath);
                    System.out.println("------EcomOrderCreation complete"+tc);
                }
            }
            if (steps.contains("stepOB6")) {
                if (d3.get(tc) != null) {
                    Main6a_WaveToReleaseTask.main(filePath, tc, env);
                }
            }
            if (steps.contains("stepOB7")) {
                if (d4.get(tc) != null) {
                    Main7_MakePickCart.main(filePath, tc, env);
                }
            }
            if (steps.contains("stepOB1A")) {
                if (d4.get(tc) != null) {
                    System.out.println("OrderReadyForPacking: Not implemented yet");
                    MainA1_OrderReadyForPacking.main(filePath, "OrderReadyForPacking");
                    Thread.sleep(800);
                }
            }
            if (steps.contains("stepOB8")) {
                if (d4.get(tc) != null) {
                    Main8_Packing.main(filePath, tc, env);
                }
            }
            if (steps.contains("stepOB2A")) {
                System.out.println("oLPNPrepared: Not implemented yet");
                MainA2_oLPNPrepared.main(filePath,"oLPNPrepared");
                Thread.sleep(800);

            }

            if (steps.contains("stepOB9")) {
                System.out.println("Palletisation: Not implemented yet");
                Thread.sleep(800);
            }
            if (steps.contains("stepOB3A")) {
                System.out.println("oLPNPalletised: Not implemented yet");
                MainA3_oLPNPalletised.main(filePath,"oLPNPalletised");
                Thread.sleep(800);
            }
            if (steps.contains("stepOB4A")) {
                System.out.println("PalletReady: Not implemented yet");
                MainA4_PalletReady.main(filePath,"PalletReady");
                Thread.sleep(800);
            }


            if (steps.contains("stepOB10")) {
                if (d6.get(tc) != null) {

                    Main10_OutboundPutaway.main(filePath, tc, env);
                    Thread.sleep(800);
                }
            }

            if (steps.contains("stepOB11")) {
                if (d6.get(tc) != null) {
                    Main11_Shipment.main(filePath, "Outbound", tc, env);
                    Thread.sleep(800);
                }
            }

        }






//            if(steps.contains("stepOB9")){
//                System.out.println("Palletisation: Not implemented yet");
//                Thread.sleep(800);
//            }
//            if(steps.contains("stepOB10")){
//                System.out.println("Outbound Putaway: Not implemented yet");
//                Main10_OutboundPutaway.main(filePath,tc,env);
//                Thread.sleep(800);
//            }
//            if(steps.contains("stepOB11")){
//                System.out.println("Shipment: Not implemented yet");
//                Thread.sleep(800);
//            }
//        }
        workbook.close();
        fis.close();
    }
    private Map<String,List<Row>> groupRows(Sheet sheet){
        Map<String,List<Row>> map=new LinkedHashMap<>();
        if(sheet==null)return map;
        Iterator<Row> it=sheet.iterator();
        if(it.hasNext())it.next();
        while(it.hasNext()){
            Row row=it.next();
            if(row==null)continue;
            Cell c=row.getCell(0);
            if(c==null)continue;
            String tc=c.toString().trim();
            if(!tc.matches("TST_\\d+"))continue;
            map.computeIfAbsent(tc,x->new ArrayList<>()).add(row);
        }
        return map;
    }
    private String getCellValue(Row row, int index) {
        if (row == null || index < 0) return "";
        Cell cell = row.getCell(index);
        if (cell == null) return "";
        cell.setCellType(CellType.STRING); // optional: force string
        return cell.getStringCellValue().trim();
    }



}