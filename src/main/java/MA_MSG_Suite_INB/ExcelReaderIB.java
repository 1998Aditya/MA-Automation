package MA_MSG_Suite_INB;

/**
 * Central place to keep all Excel file locations
 * for the Inbound (IB) automation suite.
 *
 * You can reference these constants from any class, e.g.
 *  ExcelReaderIB.DATA_EXCEL_PATH
 *  ExcelReaderIB.LOGIN_EXCEL_PATH
 */
public final class ExcelReaderIB {

    // ✅ path for data
    public static final String DATA_EXCEL_PATH =
            "C://Users//911136//IdeaProjects//msg-runner//auto_msg.xlsx";

    // 👉 Login.xlsx only for Login sheet
    public static final String LOGIN_EXCEL_PATH =
            "C://Users//911136//IdeaProjects//msg-runner//Login.xlsx";

    // prevent instantiation
    private ExcelReaderIB() {
    }
}
