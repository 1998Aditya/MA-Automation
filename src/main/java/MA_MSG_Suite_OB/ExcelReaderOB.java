package MA_MSG_Suite_OB;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;



/**
 * Central place to keep all Excel file locations
 * for the Outbound (OB) automation suite.
 *
 * You can reference these constants from any class, e.g.
 *  ExcelReaderOB.DATA_EXCEL_PATH
 *  ExcelReaderOB.LOGIN_EXCEL_PATH
 */
public final class ExcelReaderOB {

    // ✅ path for data
    public static final String DATA_EXCEL_PATH =
            "C://Users//2210420//IdeaProjects//msg-runner//OOdata.xlsx";

    // 👉 Login.xlsx only for Login sheet
    public static final String LOGIN_EXCEL_PATH =
            "C://Users//2210420//IdeaProjects//msg-runner//Login.xlsx";

    // prevent instantiation
    private ExcelReaderOB() {
    }
}
