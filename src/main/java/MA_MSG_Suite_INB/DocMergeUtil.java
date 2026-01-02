package MA_MSG_Suite_INB;

import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.*;

public class DocMergeUtil {

    // 🔹 Merge all docs that start with a given testcase ID into one final doc
    public static void mergeDocsForTestcase(String basePath, String testcaseId) throws Exception {

        File dir = new File(basePath);
        File[] files = dir.listFiles((d, name) ->
                name.startsWith(testcaseId) && name.endsWith(".docx"));

        if (files == null || files.length == 0) {
            System.out.println("⚠ No docs found for " + testcaseId);
            return;
        }

        Arrays.sort(files, Comparator.comparingLong(File::lastModified));

        System.out.println("📄 Found " + files.length + " docs for " + testcaseId + ":");
        for (File f : files) {
            System.out.println("   - " + f.getName());
        }

        try (XWPFDocument merged = new XWPFDocument(new FileInputStream(files[0]))) {
            System.out.println("➡ Starting merge with: " + files[0].getName());

            for (int i = 1; i < files.length; i++) {
                System.out.println("➡ Appending: " + files[i].getName());
                try (XWPFDocument doc = new XWPFDocument(new FileInputStream(files[i]))) {
                    for (XWPFParagraph p : doc.getParagraphs()) {
                        XWPFParagraph newPara = merged.createParagraph();
                        newPara.createRun().setText(p.getText());
                    }
                }
            }

            String outputPath = basePath + File.separator + testcaseId + ".docx";
            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                merged.write(out);
            }
            System.out.println("✅ Merged docs for " + testcaseId + " → " + outputPath);
        }
    }

    // 🔹 Merge for all testcase IDs present in the folder (pattern TST_xxx)
    public static void mergeAllTestcases(String basePath) throws Exception {

        File dir = new File(basePath);
        File[] files = dir.listFiles((d, name) -> name.matches("TST_\\d+.*\\.docx"));
        if (files == null || files.length == 0) {
            System.out.println("⚠ No testcase docs found.");
            return;
        }

        Map<String, List<File>> grouped = new HashMap<>();
        for (File f : files) {
            String name = f.getName();
            String[] parts = name.split("_");
            if (parts.length >= 2) {
                String prefix = parts[0] + "_" + parts[1]; // e.g. TST_200
                grouped.computeIfAbsent(prefix, k -> new ArrayList<>()).add(f);
            }
        }

        for (String tc : grouped.keySet()) {
            System.out.println("\n===== Merging group for " + tc + " =====");
            mergeDocsForTestcase(basePath, tc);
        }
    }

    // 🔹 Merge docs grouped by prefix (works for TSG_ or TST_)
    public static void mergeByPrefix(String basePath, String prefix) throws Exception {

        File dir = new File(basePath);
        File[] files = dir.listFiles((d, name) -> name.startsWith(prefix) && name.endsWith(".docx"));

        if (files == null || files.length == 0) {
            System.out.println("⚠ No docs found with prefix " + prefix + " in " + basePath);
            return;
        }

        Map<String, List<File>> grouped = new HashMap<>();
        for (File f : files) {
            String name = f.getName();
            String[] parts = name.split("_");
            if (parts.length >= 2) {
                String groupKey = parts[0] + "_" + parts[1]; // e.g. TSG_10 or TST_200
                grouped.computeIfAbsent(groupKey, k -> new ArrayList<>()).add(f);
            }
        }

        for (Map.Entry<String, List<File>> entry : grouped.entrySet()) {

            String group = entry.getKey();
            List<File> groupFiles = entry.getValue();
            groupFiles.sort(Comparator.comparingLong(File::lastModified));

            System.out.println("\n===== Merging group for " + group + " =====");
            System.out.println("📄 Found " + groupFiles.size() + " docs for " + group + ":");
            for (File f : groupFiles) System.out.println("   - " + f.getName());

            try (XWPFDocument merged = new XWPFDocument(new FileInputStream(groupFiles.get(0)))) {
                System.out.println("➡ Starting merge with: " + groupFiles.get(0).getName());

                for (int i = 1; i < groupFiles.size(); i++) {
                    System.out.println("➡ Appending: " + groupFiles.get(i).getName());
                    try (XWPFDocument doc = new XWPFDocument(new FileInputStream(groupFiles.get(i)))) {
                        for (XWPFParagraph p : doc.getParagraphs()) {
                            XWPFParagraph newPara = merged.createParagraph();
                            newPara.createRun().setText(p.getText());
                        }
                    }
                }

                String outputPath = basePath + File.separator + group + ".docx";
                try (FileOutputStream out = new FileOutputStream(outputPath)) {
                    merged.write(out);
                }
                System.out.println("✅ Merged docs for " + group + " → " + outputPath);
            }
        }
    }
}
