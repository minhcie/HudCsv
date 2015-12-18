package com.cie;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import org.apache.log4j.Logger;

public class EnrollmentCsv {
    static final Logger log = Logger.getLogger(EnrollmentCsv.class.getName());

    public String exportId;
    public String projectEntryId;
    public String personalId;
    public String projectId;
    public Date entryDate;
    public int priorResidence = 99;
    public String otherPriorResidence;
    public int disablingCondition = 99;
    public int continuouslyHomelessOneYear = 99;
    public int timesHomelessPast3Years = 99;
    public int monthsHomelessPast3Years = 99;
    public int monthsHomelessThisTime = 99;
    public int housingStatus = 99;
    public String lastPermanentZip;
    public Date created;
    public Date updated;
    public String userId;

    public static void importData(Connection conn, XSSFSheet sheet) throws Exception {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");

        Iterator<Row> rowIterator = sheet.iterator();
        Row row;
        Cell cell;
        while (rowIterator.hasNext()) {
            row = rowIterator.next();

            // Ignore column header row.
            if (row.getRowNum() == 0) {
                continue;
            }

            // Row data.
            EnrollmentCsv enroll = new EnrollmentCsv();

            // Iterate to all cells (including empty cell).
            int minColIndex = row.getFirstCellNum();
            int maxColIndex = row.getLastCellNum();
            for (int colIndex = minColIndex; colIndex < maxColIndex; colIndex++) {
                cell = row.getCell(colIndex);
                if (cell == null) {
                    log.info("row " + row.getRowNum() + " col " + colIndex + " is empty");
                    continue;
                }

                String cellValue = ExcelUtils.getCellValue(cell);
                switch (colIndex) {
                    case 0:
                        enroll.projectEntryId = cellValue;
                        break;
                    case 1:
                        enroll.personalId = cellValue;
                        break;
                    case 2:
                        enroll.projectId = cellValue;
                        break;
                    case 3:
                        if (cellValue.length() > 0) {
                            enroll.entryDate = sdf.parse(cellValue);
                        }
                        break;
                    case 6:
                        if (cellValue.length() > 0) {
                            enroll.priorResidence = Integer.parseInt(cellValue);
                        }
                        break;
                    case 7:
                        enroll.otherPriorResidence = cellValue;
                        break;
                    case 9:
                        if (cellValue.length() > 0) {
                            enroll.disablingCondition = Integer.parseInt(cellValue);
                        }
                        break;
                    case 12:
                        if (cellValue.length() > 0) {
                            enroll.timesHomelessPast3Years = Integer.parseInt(cellValue);
                        }
                        break;
                    case 13:
                        if (cellValue.length() > 0) {
                            enroll.monthsHomelessPast3Years = Integer.parseInt(cellValue);
                        }
                        break;
                    case 14:
                        if (cellValue.length() > 0) {
                            enroll.housingStatus = Integer.parseInt(cellValue);
                        }
                        break;
                    case 26:
                        enroll.lastPermanentZip = cellValue;
                        break;
                    case 76:
                        if (cellValue.length() > 0) {
                            enroll.created = sdf2.parse(cellValue);
                        }
                        break;
                    case 77:
                        if (cellValue.length() > 0) {
                            enroll.updated = sdf.parse(cellValue);
                        }
                        break;
                    case 78:
                        enroll.userId = cellValue;
                        break;
                    case 80:
                        enroll.exportId = cellValue;
                        break;
                    default:
                        break;
                }
            }

            // Insert row data.
            enroll.debug();
            enroll.insert(conn);
        }
    }

    public void insert(Connection conn) {
        try {
            StringBuffer sb = new StringBuffer();
            sb.append("INSERT INTO enrollment_csv (exportId, projectEntryId, ");
            sb.append("personalId, projectId, entryDate, priorResidence, ");
            sb.append("otherPriorResidence, disablingCondition, continuouslyHomelessOneYear, ");
            sb.append("timesHomelessPast3Years, monthsHomelessPast3Years, ");
            sb.append("monthsHomelessThisTime, housingStatus, lastPermanentZip, ");
            sb.append("created, updated, userId)");
            sb.append("VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");

            PreparedStatement ps = conn.prepareStatement(sb.toString());
            ps.setString(1, SqlString.encode(this.exportId));
            ps.setString(2, SqlString.encode(this.projectEntryId));
            ps.setString(3, SqlString.encode(this.personalId));
            ps.setString(4, SqlString.encode(this.projectId));
            java.sql.Date sqlDate = new java.sql.Date(this.entryDate.getTime());
            ps.setDate(5, sqlDate);
            ps.setInt(6, this.priorResidence);
            ps.setString(7, SqlString.encode(this.otherPriorResidence));
            ps.setInt(8, this.disablingCondition);
            ps.setInt(9, this.continuouslyHomelessOneYear);
            ps.setInt(10, this.timesHomelessPast3Years);
            ps.setInt(11, this.monthsHomelessPast3Years);
            ps.setInt(12, this.monthsHomelessThisTime);
            ps.setInt(13, this.housingStatus);
            ps.setString(14, SqlString.encode(this.lastPermanentZip));

            Timestamp ts = new Timestamp(this.created.getTime());
            ps.setTimestamp(15, ts);
            ts = new Timestamp(this.updated.getTime());
            ps.setTimestamp(16, ts);
            ps.setString(17, SqlString.encode(this.userId));

            int out = ps.executeUpdate();
            if (out == 0) {
                log.error("Failed to insert enrollment_csv record!");
            }
        }
        catch (SQLException sqle) {
            log.error("SQLException EnrollmentCsv.insert(): " + sqle);
        }
        catch (Exception e) {
            log.error("Exception EnrollmentCsv.insert(): " + e);
        }
    }

    public void debug() {
        log.info("Enrollment ID: " + this.projectEntryId);
        log.info("Client ID: " + this.personalId);
        log.info("Project ID: " + this.projectId);
        this.displayPriorResidence();
        log.info("Other Prior Resident: " + this.otherPriorResidence);
        this.displayDisablingCondition();
        //log.info("Continuously Homeless One Year: " + this.continuouslyHomelessOneYear);
        log.info("Times Homeless Past 3 Years: " + this.timesHomelessPast3Years);
        log.info("Months Homeless Past 3 Years: " + this.monthsHomelessPast3Years);
        //log.info("Months Homeless This Time: " + this.monthsHomelessThisTime);
        this.displayHousingStatus();
        log.info("Last Permanent Zip: " + this.lastPermanentZip);
        log.info("Entry Date: " + this.entryDate.toString());
        log.info("Date Created: " + this.created.toString());
        log.info("Date Updated: " + this.updated.toString());
        log.info("User ID: " + this.userId);
        log.info("\n");
    }

    private void displayPriorResidence() {
        switch (this.priorResidence) {
            case 1:
                log.info("Prior Resident: Emergency shelter, including hotel or motel paid for with emergency shelter voucher");
                break;
            case 2:
                log.info("Prior Resident: Transitional housing for homeless persons");
                break;
            case 3:
                log.info("Prior Resident: Permanent housing for formerly homeless persons");
                break;
            case 4:
                log.info("Prior Resident: Psychiatric hospital or other psychiatric facility");
                break;
            case 5:
                log.info("Prior Resident: Substance abuse treatment facility or detox center");
                break;
            case 6:
                log.info("Prior Resident: Hospital or other residential non-psychiatric medical facility");
                break;
            case 7:
                log.info("Prior Resident: Jail, prison or juvenile detention facility");
                break;
            case 8:
                log.info("Prior Resident: Client doesn’t know");
                break;
            case 9:
                log.info("Prior Resident: Client refused");
                break;
            case 12:
                log.info("Prior Resident: Staying or living in a family member’s room, apartment or house");
                break;
            case 13:
                log.info("Prior Resident: Staying or living in a friend’s room, apartment or house");
                break;
            case 14:
                log.info("Prior Resident: Hotel or motel paid for without emergency shelter voucher");
                break;
            case 15:
                log.info("Prior Resident: Foster care home or foster care group home");
                break;
            case 16:
                log.info("Prior Resident: Place not meant for habitation");
                break;
            case 17:
                log.info("Prior Resident: Other");
                break;
            case 18:
                log.info("Prior Resident: Safe Haven");
                break;
            case 19:
                log.info("Prior Resident: Rental by client, with VASH subsidy");
                break;
            case 20:
                log.info("Prior Resident: Rental by client, with other ongoing housing subsidy");
                break;
            case 21:
                log.info("Prior Resident: Owned by client, with ongoing housing subsidy");
                break;
            case 22:
                log.info("Prior Resident: Rental by client, no ongoing housing subsidy");
                break;
            case 23:
                log.info("Prior Resident: Owned by client, no ongoing housing subsidy");
                break;
            case 24:
                log.info("Prior Resident: Long-term care facility or nursing home");
                break;
            case 25:
                log.info("Prior Resident: Rental by client, with GPD TIP subsidy");
                break;
            case 26:
                log.info("Prior Resident: Residential project or halfway house with no homeless criteria");
                break;
            default:
                log.info("Data not collected");
                break;
        }
    }

    private void displayDisablingCondition() {
        switch (this.disablingCondition) {
            case 0:
                log.info("Disabling Condition: No");
                break;
            case 1:
                log.info("Disabling Condition: Yes");
                break;
            case 8:
                log.info("Disabling Condition: Client doesn't know");
                break;
            case 9:
                log.info("Disabling Condition: Client refused");
                break;
            default:
                log.info("Disabling Condition: Data not collected");
                break;
        }
    }

    private void displayHousingStatus() {
        switch (this.housingStatus) {
            case 1:
                log.info("Housing Status: Category 1 - Homeless");
                break;
            case 2:
                log.info("Housing Status: Category 2 - At imminent risk of losing housing");
                break;
            case 3:
                log.info("Housing Status: At-risk of homelessness - prevention programs only");
                break;
            case 4:
                log.info("Housing Status: Stably housed");
                break;
            case 5:
                log.info("Housing Status: Category 3 - Homeless only under other federal statutes");
                break;
            case 6:
                log.info("Housing Status: Category 4 - Fleeing domestic violence");
                break;
            case 8:
                log.info("Housing Status: Client doesn't know");
                break;
            case 9:
                log.info("Housing Status: Client refused");
                break;
            default:
                log.info("Housing Status: Data not collected");
                break;
        }
    }
}
