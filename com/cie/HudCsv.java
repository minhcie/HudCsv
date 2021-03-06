package com.cie;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.log4j.Logger;

public class HudCsv {
    private static final Logger log = Logger.getLogger(HudCsv.class.getName());

    public static void main(String[] args) {
        if (args.length == 0 || args.length < 2) {
            usage();
        }

        if (args[0].equalsIgnoreCase("import")) {
            importData(args[1]);
        }
        else if (args[0].equalsIgnoreCase("view")) {
            viewData(args[1]);
        }
        else {
            usage();
        }
    }

    static void usage() {
        System.err.println("usage: java -jar HudCsv.jar import/view <excel-file.xlsx>");
        System.err.println("");
        System.exit(-1);
    }

    static void importData(String csvName) {
        Connection conn = null;
        try {
            log.info("Reading excel file " + csvName + " ...");
            File xcel = new File(csvName);
            if (!xcel.exists()) {
                log.error("File not found");
                return;
            }

            conn = DbUtils.getDBConnection();
            if (conn == null) {
                return;
            }

            // Get the workbook object for xlsx file.
            XSSFWorkbook wBook = new XSSFWorkbook(new FileInputStream(xcel));
            XSSFSheet sheet = null;
            String sheetName = null;

            // Iterate thru all sheets.
            int numberOfSheets = wBook.getNumberOfSheets();
            for (int i = 0; i < numberOfSheets; i++) {
                sheet = wBook.getSheetAt(i);
                sheetName = sheet.getSheetName();

                // Import data.
                log.info("Reading " + sheetName + " ...");
                sheetName = sheetName.toLowerCase();
                switch (sheetName) {
                    case "export":
                        ExportCsv.importData(conn, sheet);
                        break;
                    case "organization":
                        OrganizationCsv.importData(conn, sheet);
                        break;
                    case "project":
                        ProjectCsv.importData(conn, sheet);
                        break;
                    case "client":
                        ClientCsv.importData(conn, sheet);
                        break;
                    case "enrollment":
                        EnrollmentCsv.importData(conn, sheet);
                        break;
                    case "exit":
                        ExitCsv.importData(conn, sheet);
                        break;
                    case "incomebenefits":
                        IncomeBenefitsCsv.importData(conn, sheet);
                        break;
                    case "healthanddv":
                        HealthDvCsv.importData(conn, sheet);
                        break;
                    case "disabilities":
                        DisabilitiesCsv.importData(conn, sheet);
                        break;
                    case "employmenteducation":
                        EmploymentEducationCsv.importData(conn, sheet);
                        break;
                    case "casemanager":
                        CaseManagerCsv.importData(conn, sheet);
                        break;
                    case "exitext":
                        ExitExtCsv.importData(conn, sheet);
                        break;
                    default:
                        break;
                }
            }
        }
        catch (IOException ioe) {
            log.error(ioe.getMessage());
        }
        catch (Exception e) {
            log.error(e.getMessage());
        }
        finally {
            DbUtils.closeConnection(conn);
        }
    }

    static void viewData(String csvName) {
        Connection conn = null;
        try {
            conn = DbUtils.getDBConnection();
            if (conn == null) {
                return;
            }

            log.info("Viewing " + csvName + " data ...");
            String data = csvName.toLowerCase();
            switch (data) {
                case "client":
                    List<ClientCsv> clients = ClientCsv.load(conn);
                    for (int i = 0; i < clients.size(); i++) {
                        ClientCsv c = clients.get(i);
                        c.display();
                    }
                    break;
                default:
                    break;
            }
        }
        catch (Exception e) {
            log.error(e.getMessage());
        }
        finally {
            DbUtils.closeConnection(conn);
        }
    }
}
