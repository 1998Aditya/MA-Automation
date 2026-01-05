
MSG-RUNNER – COMPLETE README
============================

This document consolidates ALL discussions, architecture decisions, utilities, and automation features
implemented across the MA_MSG_Suite_INB automation framework.

-------------------------------------------------------------------------------
PROJECT: msg-runner
-------------------------------------------------------------------------------

What you got:
-------------
- A Spring Boot web app that serves a simple HTML UI at / (http://localhost:8080)
- A controller that triggers MA_MSG_Suite_INB.MSG_MAIN (a runnable stub by default)
- Your original MSG_MAIN file included as MSG_MAIN_original.java inside the MA_MSG_Suite_INB package
  (replace the stub MSG_MAIN.java with your original if you want to run your real flow).
- Dockerfile and Maven pom.xml included.

How to run locally (no Docker):
-------------------------------
1. Open terminal, go to project root (where pom.xml is).
2. mvn clean package
3. java -jar target/msg-runner-1.0.0.jar
4. Open http://localhost:8080

How to run with Docker:
----------------------
1. docker build -t msg-runner .
2. docker run -p 8080:8080 msg-runner
3. Open http://localhost:8080

Notes:
------
- The project includes a safe, compiling stub for MA_MSG_Suite_INB.MSG_MAIN (so the project builds out-of-the-box).
- Your original MSG_MAIN (the file you shared) is present as MSG_MAIN_original.java in the same package.
- To run real flows, replace MSG_MAIN stub with original and ensure Selenium & ChromeDriver configs exist.
- For Docker/headless execution, ChromeOptions must be set to headless and Chrome installed.

-------------------------------------------------------------------------------
ARCHITECTURE OVERVIEW
-------------------------------------------------------------------------------

Package:
--------
MA_MSG_Suite_INB

Core Concepts:
--------------
- All automation is DATA-DRIVEN via Excel
- Each Excel sheet contains a mandatory column:
    Testcase (Example: TST_1, TST_2, TST_3)
- Execution happens TESTCASE-FIRST:
    Testcase 1 → All steps
    Testcase 2 → All steps
    Testcase 3 → All steps

This guarantees full end-to-end flow per testcase.

-------------------------------------------------------------------------------
CENTRAL EXCEL CONFIGURATION
-------------------------------------------------------------------------------

Class: ExcelReaderIB
--------------------
Purpose:
- Centralized Excel path management

Fields:
- DATA_EXCEL_PATH   → auto_msg.xlsx (all data sheets)
- LOGIN_EXCEL_PATH  → Login.xlsx (ONLY login sheet)

All classes reference ExcelReaderIB and never hardcode paths.

-------------------------------------------------------------------------------
TESTCASE SEQUENCING MECHANISM
-------------------------------------------------------------------------------

Utility logic:
--------------
- groupRows(Sheet sheet)
- Groups rows by Testcase column (TST_1, TST_2…)
- Uses LinkedHashMap to preserve sequence

Execution order:
----------------
FOR each Testcase:
    FOR each Step:
        Execute step using ONLY rows of that testcase

This ensures strict testcase isolation and sequencing.

-------------------------------------------------------------------------------
INBOUND CONTROLLER FLOW
-------------------------------------------------------------------------------

Method:
-------
executeInboundSteps(jobId, steps, env)

Enhanced Behavior:
------------------
- Reads Testcase sequence first
- For each Testcase:
    step1 → Login
    step2 → MSG_LPN_ASN_Creation
    step3 → MSG_Item_ASN_Creation
    step4 → Reports
    step5 → Condition codes
    step6 → Induction
    step7 → Manual_Item_rcv
    step8 → Manual_LPN_rcv
    step9 → Manual_pallet_putaway

Driver lifecycle:
-----------------
- Single WebDriver per testcase
- Clean quit after testcase completes

-------------------------------------------------------------------------------
MANUAL STEP CLASSES
-------------------------------------------------------------------------------

Manual_Item_rcv
---------------
- Uses DATA_EXCEL_PATH
- Sheet: item_rcv
- Fully testcase-aware
- ASN grouping logic preserved
- Item barcode fetched via API (no UI navigation)

Manual_LPN_rcv
--------------
- Uses DATA_EXCEL_PATH
- Sheet: LPN_rcv
- ASN grouping + dock release logic

Manual_pallet_putaway
---------------------
- Uses DATA_EXCEL_PATH
- Sheet: pallet_putaway
- DropZone & LocationBarcode fetched via API
- UI navigation removed completely

-------------------------------------------------------------------------------
API-BASED SERVICES (UI REMOVED)
-------------------------------------------------------------------------------

ItemBarcodeService
------------------
Purpose:
- Fetch PrimaryBarCode or MAUJDSDefaultEANBarcode

API:
POST /item-master/api/item-master/item/search

Usage:
String barcode = ItemBarcodeService.getItemBarcode(itemId);

Priority:
1. PrimaryBarCode
2. Extended.MAUJDSDefaultEANBarcode

-------------------------------------------------------------------------------

LocationBarcodeService
----------------------
Purpose:
- Fetch LocationBarcode

APIs:
POST /dcinventory/api/dcinventory/location/quickSearch

Queries:
- By LocationId
- By TaskMovementZoneId

Usage:
String locationBarcode =
    LocationBarcodeService.getLocationBarcodeByTaskMovementZone("MHE_TASKMOVE_UNAVAIL");

Replaces:
- Menu navigation
- Pick Drop Locations UI
- Filter & capture via Selenium

-------------------------------------------------------------------------------
TESTCASE REPORTING & SCREENSHOTS
-------------------------------------------------------------------------------

Class: TestcaseReporter
-----------------------
Purpose:
- Capture screenshots
- Generate Word document per testcase

Behavior:
---------
- One DOCX per Testcase
- Screenshots auto-attached
- Step name + timestamp included

Usage:
------
TestcaseReporter.startTestcase("TST_1");
TestcaseReporter.capture(driver, "After ASN Creation");
TestcaseReporter.endTestcase();

Inserted automatically into:
- All Manual_* classes
- Controller boundaries

-------------------------------------------------------------------------------
KEY BENEFITS
-------------------------------------------------------------------------------

✔ Fully data-driven
✔ API-first (UI removed wherever possible)
✔ Faster execution
✔ Stable automation
✔ Clear testcase traceability
✔ Clean separation of concerns
✔ Production-ready framework

-------------------------------------------------------------------------------
END OF DOCUMENT
-------------------------------------------------------------------------------
