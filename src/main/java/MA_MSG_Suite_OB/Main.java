//package MA_MSG_Suite_OB;
//
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//import java.io.FileInputStream;
//import java.util.*;
//
//
//public class Main {
//    public static String Hostname = "ujdss.sce.manh.com";//Load

//    public static String AuthHost = "ujdss-auth.sce.manh.com";//Load
//    //public static String AuthUsername= "ujdsstage1";//Load
//    // public static String AuthPassword= "Earth-Moon-Sun1";//Load
//    public static String Organization = "HEERLEN51";//Load
//String filePath = "C:\\Users\\2210420\\IdeaProjects\\Testcases\\OOdata.xlsx";
//
//    public static void main(String[] args) throws Exception {
//        String filePath = "C:\\Users\\2210420\\IdeaProjects\\Testcases\\OOdata.xlsx";
//        FileInputStream fis = new FileInputStream(filePath);
//        Workbook workbook = new XSSFWorkbook(fis);
//
//        Sheet sheet1 = workbook.getSheet("Data Creation");
//        Sheet sheet2 = workbook.getSheet("Order Creation");
//        Sheet sheet3 = workbook.getSheet("OPS_Tab");
//        Sheet sheet4 = workbook.getSheet("Tasks");
//
//        Map<String, List<Row>> class1Data = groupRowsByTestcase(sheet1);
//        Map<String, List<Row>> class2Data = groupRowsByTestcase(sheet2);
//        Map<String, List<Row>> class3Data = groupRowsByTestcase(sheet3);
//        Map<String, List<Row>> class4Data = groupRowsByTestcase(sheet4);
//
//        Set<String> allTestcases = new LinkedHashSet<>();
//        allTestcases.addAll(class1Data.keySet());
//        allTestcases.addAll(class2Data.keySet());
//        allTestcases.addAll(class3Data.keySet());
//        allTestcases.addAll(class4Data.keySet());
//
//        for (String testcase : allTestcases) {
//            System.out.println("Executing Testcase: " + testcase);
//
//            // Execute Data Creation rows
//            List<Row> dataRows = class1Data.get(testcase);
//            if (dataRows != null) {
//                for (Row row : dataRows) {
//                    MainA_FinalInventory.execute(row);
//                    //Class1();
//                    System.out.println("✅ Data Creation row executed for " + testcase);
//                }
//            }
//
//             //Execute Order Creation rows
//            List<Row> orderRows = class2Data.get(testcase);
//            if (orderRows != null) {
//                for (Row row : orderRows) {
//                    MainB_OrderCreation.main(testcase);
//                    System.out.println("✅----------------------------------------- Order Creation row executed for " + testcase);
//                }
//            }
//
//            List<Row> WaveRows = class3Data.get(testcase);
//            if (WaveRows != null) {
//                for (Row row : WaveRows) {
//                    MainC1_WaveToReleaseTask.main(filePath, testcase);
//                    System.out.println("✅------------------------------------------- WAVE Creation row executed for " + testcase);
//                }
//            }
//
//            List<Row> ManualTaskRows = class3Data.get(testcase);
//            if (ManualTaskRows != null) {
//                for (Row row : ManualTaskRows) {
//                    MainD_MakePickCart.main(filePath, testcase);
//                    System.out.println("✅------------------------------------------ Task Execution row executed for " + testcase);
//                }
//            }
//
//            List<Row> ManualPackTaskRows = class3Data.get(testcase);
//            if (ManualPackTaskRows != null) {
//                for (Row row : ManualPackTaskRows) {
//                    MainE_Packing.main(filePath, testcase);
//                    System.out.println("✅-------------------------------------- Pack Task Execution row executed for " + testcase);
//                }
//            }
//
//
//
//
//            System.out.println("🔚 Finished Testcase: " + testcase + "\n");
//        }
//
//        workbook.close();
//        fis.close();
//    }
//
//
//    private static Map<String, List<Row>> groupRowsByTestcase(Sheet sheet) {
//        Map<String, List<Row>> grouped = new LinkedHashMap<>();
//        Iterator<Row> iterator = sheet.iterator();
//        iterator.next(); // Skip header
//
//        while (iterator.hasNext()) {
//            Row row = iterator.next();
//            if (row == null) continue;
//
//            Cell testcaseCell = row.getCell(0);
//            if (testcaseCell == null || testcaseCell.getCellType() == CellType.BLANK) continue;
//
//            String testcase = testcaseCell.toString().trim();
//            if (!testcase.matches("TST_\\d+")) continue; // Only allow TST_001, TST_002, etc.
//
//            grouped.computeIfAbsent(testcase, k -> new ArrayList<>()).add(row);
//        }
//        return grouped;
//    }
//}
//
//
//
//
//
//
//
//
