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

public class ExitExtCsv {
    static final Logger log = Logger.getLogger(ExitExtCsv.class.getName());

    public String exportId;
    public String exitExtId;
    public String personalId;
    public String exitId;
    public int reasonForLeaving = 99;
    public int exitResidencePrior = 99;
    public int exitDisablingCondition = 99;
    public int exitTimesHomelessPast3Years = 99;
    public int exitMonthsHomelessPast3Years = 99;
    public int exitHousingStatus = 99;
    public String exitLastPermanentZip;
    public Date created;
    public Date updated;
    public String userId;

    public static void importData(Connection conn, XSSFSheet sheet) throws Exception {
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
            ExitExtCsv ext = new ExitExtCsv();

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
                        ext.exitExtId = cellValue;
                        break;
                    case 1:
                        ext.personalId = cellValue;
                        break;
                    case 2:
                        ext.exitId = cellValue;
                        break;
                    case 3:
                        if (cellValue.length() > 0) {
                            ext.reasonForLeaving = Integer.parseInt(cellValue);
                        }
                        break;
                    case 4:
                        if (cellValue.length() > 0) {
                            ext.exitResidencePrior = Integer.parseInt(cellValue);
                        }
                        break;
                    case 6:
                        if (cellValue.length() > 0) {
                            ext.exitDisablingCondition = Integer.parseInt(cellValue);
                        }
                        break;
                    case 9:
                        if (cellValue.length() > 0) {
                            ext.exitTimesHomelessPast3Years = Integer.parseInt(cellValue);
                        }
                        break;
                    case 10:
                        if (cellValue.length() > 0) {
                            ext.exitMonthsHomelessPast3Years = Integer.parseInt(cellValue);
                        }
                        break;
                    case 11:
                        if (cellValue.length() > 0) {
                            ext.exitHousingStatus = Integer.parseInt(cellValue);
                        }
                        break;
                    case 12:
                        ext.exitLastPermanentZip = cellValue;
                        break;
                    case 13:
                        if (cellValue.length() > 0) {
                            ext.created = sdf2.parse(cellValue);
                        }
                        break;
                    case 14:
                        if (cellValue.length() > 0) {
                            ext.updated = sdf2.parse(cellValue);
                        }
                        break;
                    case 15:
                        ext.userId = cellValue;
                        break;
                    case 17:
                        ext.exportId = cellValue;
                        break;
                    default:
                        break;
                }
            }

            // Insert row data.
            ext.debug();
            ext.insert(conn);
        }
    }

    public void insert(Connection conn) {
        try {
            StringBuffer sb = new StringBuffer();
            sb.append("INSERT INTO exit_ext_csv (exportId, exitExtId, personalId, ");
            sb.append("exitId, reasonForLeaving, exitResidencePrior, exitDisablingCondition, ");
            sb.append("exitTimesHomelessPast3Years, exitMonthsHomelessPast3Years, ");
            sb.append("exitHousingStatus, exitLastPermanentZip, created, updated, userId) ");
            sb.append("VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");

            PreparedStatement ps = conn.prepareStatement(sb.toString());
            ps.setString(1, SqlString.encode(this.exportId));
            ps.setString(2, SqlString.encode(this.exitExtId));
            ps.setString(3, SqlString.encode(this.personalId));
            ps.setString(4, SqlString.encode(this.exitId));
            ps.setInt(5, this.reasonForLeaving);
            ps.setInt(6, this.exitResidencePrior);
            ps.setInt(7, this.exitDisablingCondition);
            ps.setInt(8, this.exitTimesHomelessPast3Years);
            ps.setInt(9, this.exitMonthsHomelessPast3Years);
            ps.setInt(10, this.exitHousingStatus);
            ps.setString(11, SqlString.encode(this.exitLastPermanentZip));

            Timestamp ts = new Timestamp(this.created.getTime());
            ps.setTimestamp(12, ts);
            ts = new Timestamp(this.updated.getTime());
            ps.setTimestamp(13, ts);
            ps.setString(14, SqlString.encode(this.userId));

            int out = ps.executeUpdate();
            if (out == 0) {
                log.error("Failed to insert exit_ext_csv record!");
            }
        }
        catch (SQLException sqle) {
            log.error("SQLException ExitExtCsv.insert(): " + sqle);
        }
        catch (Exception e) {
            log.error("Exception ExitExtCsv.insert(): " + e);
        }
    }

    public void debug() {
        log.info("Enrollment Ext ID: " + this.exitExtId);
        log.info("Client ID: " + this.personalId);
        log.info("Exit ID: " + this.exitId);
        this.displayReasonForLeaving();
        this.displayDisablingCondition();
        log.info("Times Homeless Past 3 Years: " + this.exitTimesHomelessPast3Years);
        log.info("Months Homeless Past 3 Years: " + this.exitMonthsHomelessPast3Years);
        this.displayHousingStatus();
        log.info("Last Permanent Zip: " + this.exitLastPermanentZip);
        log.info("Date Created: " + this.created.toString());
        log.info("Date Updated: " + this.updated.toString());
        log.info("User ID: " + this.userId);
        log.info("\n");
    }

    private void displayReasonForLeaving() {
        switch (this.reasonForLeaving) {
            case 1:
                log.info("Reason For Leaving: Left for housing opp. before completing program");
                break;
            case 2:
                log.info("Reason For Leaving: Completed program");
                break;
            case 3:
                log.info("Reason For Leaving: Non-Payment of rent/occupancy charge");
                break;
            case 4:
                log.info("Reason For Leaving: Non-compliance with program");
                break;
            case 5:
                log.info("Reason For Leaving: Criminal Activity");
                break;
            case 6:
                log.info("Reason For Leaving: Reached Maximum Time Allowed for Project");
                break;
            case 7:
                log.info("Reason For Leaving: Needs could not be met");
                break;
            case 8:
                log.info("Reason For Leaving: Disagreement with rules/person");
                break;
            case 9:
                log.info("Reason For Leaving: Death");
                break;
            case 10:
                log.info("Reason For Leaving: Unknown/Disappeared");
                break;
            default:
                log.info("Reason For Leaving: Other");
                break;
        }
    }

    private void displayPriorResidence() {
        switch (this.exitResidencePrior) {
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
        switch (this.exitDisablingCondition) {
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
        switch (this.exitHousingStatus) {
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
