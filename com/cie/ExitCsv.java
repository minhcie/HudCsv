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

public class ExitCsv {
    static final Logger log = Logger.getLogger(ExitCsv.class.getName());

    public String exportId;
    public String exitId;
    public String projectEntryId;
    public String personalId;
    public Date exitDate;
    public int destination = 99;
    public String otherDestination;
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
            ExitCsv exit = new ExitCsv();

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
                        exit.exitId = cellValue;
                        break;
                    case 1:
                        exit.projectEntryId = cellValue;
                        break;
                    case 2:
                        exit.personalId = cellValue;
                        break;
                    case 3:
                        if (cellValue.length() > 0) {
                            exit.exitDate = sdf.parse(cellValue);
                        }
                        break;
                    case 4:
                        if (cellValue.length() > 0) {
                            exit.destination = Integer.parseInt(cellValue);
                        }
                        break;
                    case 5:
                        exit.otherDestination = cellValue;
                        break;
                    case 23:
                        if (cellValue.length() > 0) {
                            exit.created = sdf2.parse(cellValue);
                        }
                        break;
                    case 24:
                        if (cellValue.length() > 0) {
                            exit.updated = sdf2.parse(cellValue);
                        }
                        break;
                    case 25:
                        exit.userId = cellValue;
                        break;
                    case 27:
                        exit.exportId = cellValue;
                        break;
                    default:
                        break;
                }
            }

            // Insert row data.
            exit.debug();
            exit.insert(conn);
        }
    }

    public void insert(Connection conn) {
        try {
            StringBuffer sb = new StringBuffer();
            sb.append("INSERT INTO exit_csv (exportId, exitId, projectEntryId, ");
            sb.append("personalId, exitDate, destination, otherDestination, ");
            sb.append("created, updated, userId) ");
            sb.append("VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");

            PreparedStatement ps = conn.prepareStatement(sb.toString());
            ps.setString(1, SqlString.encode(this.exportId));
            ps.setString(2, SqlString.encode(this.exitId));
            ps.setString(3, SqlString.encode(this.projectEntryId));
            ps.setString(4, SqlString.encode(this.personalId));
            java.sql.Date sqlDate = new java.sql.Date(this.exitDate.getTime());
            ps.setDate(5, sqlDate);
            ps.setInt(6, this.destination);
            ps.setString(7, SqlString.encode(this.otherDestination));

            Timestamp ts = new Timestamp(this.created.getTime());
            ps.setTimestamp(8, ts);
            ts = new Timestamp(this.updated.getTime());
            ps.setTimestamp(9, ts);
            ps.setString(10, SqlString.encode(this.userId));

            int out = ps.executeUpdate();
            if (out == 0) {
                log.error("Failed to insert exit_csv record!");
            }
        }
        catch (SQLException sqle) {
            log.error("SQLException ExitCsv.insert(): " + sqle);
        }
        catch (Exception e) {
            log.error("Exception ExitCsv.insert(): " + e);
        }
    }

    public void debug() {
        log.info("Exit ID: " + this.exitId);
        log.info("Enrollment ID: " + this.projectEntryId);
        log.info("Client ID: " + this.personalId);
        log.info("Exit Date: " + this.exitDate.toString());
        this.displayDestination();
        log.info("Other Destination: " + this.otherDestination);
        log.info("Date Created: " + this.created.toString());
        log.info("Date Updated: " + this.updated.toString());
        log.info("User ID: " + this.userId);
        log.info("\n");
    }

    private void displayDestination() {
        switch (this.destination) {
            case 1:
                log.info("Destination: Emergency shelter, including hotel or motel paid for with emergency shelter voucher");
                break;
            case 2:
                log.info("Destination: Transitional housing for homeless persons (including homeless youth)");
                break;
            case 3:
                log.info("Destination: Permanent housing for formerly homeless persons (such as: CoC project; or HUD legacy programs; or HOPWA PH)");
                break;
            case 4:
                log.info("Destination: Psychiatric hospital or other psychiatric facility");
                break;
            case 5:
                log.info("Destination: Substance abuse treatment facility or detox center");
                break;
            case 6:
                log.info("Destination: Hospital or other residential non-psychiatric medical facility");
                break;
            case 7:
                log.info("Destination: Jail, prison or juvenile detention facility");
                break;
            case 8:
                log.info("Destination: Client doesn't know");
                break;
            case 9:
                log.info("Destination: Client refused");
                break;
            case 10:
                log.info("Destination: Rental by client, no ongoing housing subsidy");
                break;
            case 11:
                log.info("Destination: Owned by client, no ongoing housing subsidy");
                break;
            case 12:
                log.info("Destination: Staying or living with family, temporary tenure (e.g., room, apartment or house)");
                break;
            case 13:
                log.info("Destination: Staying or living with friends, temporary tenure (.e.g., room apartment or house)");
                break;
            case 14:
                log.info("Destination: Hotel or motel paid for without emergency shelter voucher");
                break;
            case 15:
                log.info("Destination: Foster care home or foster care group home");
                break;
            case 16:
                log.info("Destination: Place not meant for habitation (e.g., a vehicle, an abandoned building, bus/train/subway station/airport or anywhere outside)");
                break;
            case 17:
                log.info("Destination: Other");
                break;
            case 18:
                log.info("Destination: Safe Haven");
                break;
            case 19:
                log.info("Destination: Rental by client, with VASH housing subsidy");
                break;
            case 20:
                log.info("Destination: Rental by client, with other ongoing housing subsidy");
                break;
            case 21:
                log.info("Destination: Owned by client, with ongoing housing subsidy");
                break;
            case 22:
                log.info("Destination: Staying or living with family, permanent tenure");
                break;
            case 23:
                log.info("Destination: Staying or living with friends, permanent tenure");
                break;
            case 24:
                log.info("Destination: Deceased");
                break;
            case 25:
                log.info("Destination: Long-term care facility or nursing home");
                break;
            case 26:
                log.info("Destination: Moved from one HOPWA funded project to HOPWA PH");
                break;
            case 27:
                log.info("Destination: Moved from one HOPWA funded project to HOPWA TH");
                break;
            case 28:
                log.info("Destination: Rental by client, with GPD TIP housing subsidy");
                break;
            case 29:
                log.info("Destination: Residential project or halfway house with no homeless criteria");
                break;
            case 30:
                log.info("Destination: No exit interview completed");
                break;
            default:
                log.info("Destination: Data not collected");
                break;
        }
    }
}
