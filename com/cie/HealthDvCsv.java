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

public class HealthDvCsv {
    static final Logger log = Logger.getLogger(HealthDvCsv.class.getName());

    public String exportId;
    public String healthDvId;
    public String projectEntryId;
    public String personalId;
    public Date infoDate;
    public int dvVictim = 99;
    public int whenOccurred = 99;
    public int pregnancyStatus = 99;
    public Date dueDate;
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
            HealthDvCsv health = new HealthDvCsv();


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
                        health.healthDvId = cellValue;
                        break;
                    case 1:
                        health.projectEntryId = cellValue;
                        break;
                    case 2:
                        health.personalId = cellValue;
                        break;
                    case 3:
                        if (cellValue.length() > 0) {
                            health.infoDate = sdf.parse(cellValue);
                        }
                        break;
                    case 4:
                        if (cellValue.length() > 0) {
                            health.dvVictim = Integer.parseInt(cellValue);
                        }
                        break;
                    case 5:
                        if (cellValue.length() > 0) {
                            health.whenOccurred = Integer.parseInt(cellValue);
                        }
                        break;
                    case 10:
                        if (cellValue.length() > 0) {
                            health.pregnancyStatus = Integer.parseInt(cellValue);
                        }
                        break;
                    case 11:
                        if (cellValue.length() > 0) {
                            health.dueDate = sdf.parse(cellValue);
                        }
                        break;
                    case 13:
                        if (cellValue.length() > 0) {
                            health.created = sdf.parse(cellValue);
                        }
                        break;
                    case 14:
                        if (cellValue.length() > 0) {
                            health.updated = sdf.parse(cellValue);
                        }
                        break;
                    case 15:
                        health.userId = cellValue;
                        break;
                    case 17:
                        health.exportId = cellValue;
                        break;
                    default:
                        break;
                }
            }

            // Insert row data.
            health.debug();
            health.insert(conn);
        }
    }

    public void insert(Connection conn) {
        try {
            StringBuffer sb = new StringBuffer();
            sb.append("INSERT INTO health_dv_csv (exportId, healthDvId, ");
            sb.append("projectEntryId, personalId, infoDate, dvVictim, ");
            sb.append("whenOccurred, pregnancyStatus, dueDate, created, ");
            sb.append("updated, userId) ");
            sb.append("VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");

            PreparedStatement ps = conn.prepareStatement(sb.toString());
            ps.setString(1, SqlString.encode(this.exportId));
            ps.setString(2, SqlString.encode(this.healthDvId));
            ps.setString(3, SqlString.encode(this.projectEntryId));
            ps.setString(4, SqlString.encode(this.personalId));
            java.sql.Date sqlDate = new java.sql.Date(this.infoDate.getTime());
            ps.setDate(5, sqlDate);
            ps.setInt(6, this.dvVictim);
            ps.setInt(7, this.whenOccurred);
            ps.setInt(8, this.pregnancyStatus);
            if (this.dueDate != null) {
                sqlDate = new java.sql.Date(this.dueDate.getTime());
                ps.setDate(9, sqlDate);
            }
            else {
                ps.setNull(9, java.sql.Types.NULL);
            }

            Timestamp ts = new Timestamp(this.created.getTime());
            ps.setTimestamp(10, ts);
            ts = new Timestamp(this.updated.getTime());
            ps.setTimestamp(11, ts);
            ps.setString(12, SqlString.encode(this.userId));

            int out = ps.executeUpdate();
            if (out == 0) {
                log.error("Failed to insert health_dv_csv record!");
            }
        }
        catch (SQLException sqle) {
            log.error("SQLException HealthDvCsv.insert(): " + sqle);
        }
        catch (Exception e) {
            log.error("Exception HealthDvCsv.insert(): " + e);
        }
    }

    public void debug() {
        log.info("Health and DV ID: " + this.healthDvId);
        log.info("Enrollment ID: " + this.projectEntryId);
        log.info("Client ID: " + this.personalId);
        log.info("Information Date: " + this.infoDate.toString());
        this.displayDVVictim();
        this.displayWhenOccurred();
        log.info("Pregnancy Status: " + this.pregnancyStatus);
        if (this.dueDate != null) {
            log.info("Due Date: " + this.dueDate.toString());
        }
        log.info("Date Created: " + this.created.toString());
        log.info("Date Updated: " + this.updated.toString());
        log.info("User ID: " + this.userId);
        log.info("\n");
    }

    private void displayDVVictim() {
        switch (this.dvVictim) {
            case 0:
                log.info("Domestic Violence Victim: No");
                break;
            case 1:
                log.info("Domestic Violence Victim: Yes");
                break;
            case 8:
                log.info("Domestic Violence Victim: Client doesn't know");
                break;
            case 9:
                log.info("Domestic Violence Victim: Client refused");
                break;
            default:
                log.info("Domestic Violence Victim: Data not collected");
                break;
        }
    }

    private void displayWhenOccurred() {
        switch (this.dvVictim) {
            case 1:
                log.info("When DV Occurred: Within the past three months");
                break;
            case 2:
                log.info("When DV Occurred: Three to six months ago (excluding six months exactly)");
                break;
            case 3:
                log.info("When DV Occurred: Six months to one year ago (excluding one year exactly)");
                break;
            case 4:
                log.info("When DV Occurred: One year or more");
                break;
            case 8:
                log.info("When DV Occurred: Client doesn't know");
                break;
            case 9:
                log.info("When DV Occurred: Client refused");
                break;
            default:
                log.info("When DV Occurred: Data not collected");
                break;
        }
    }
}
