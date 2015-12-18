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

public class DisabilitiesCsv {
    static final Logger log = Logger.getLogger(DisabilitiesCsv.class.getName());

    public String exportId;
    public String disabilitiesId;
    public String projectEntryId;
    public String personalId;
    public Date infoDate;
    public int disabilityType = 99;
    public int disabilityResponse = 99;
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
            DisabilitiesCsv dis = new DisabilitiesCsv();

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
                        dis.disabilitiesId = cellValue;
                        break;
                    case 1:
                        dis.projectEntryId = cellValue;
                        break;
                    case 2:
                        dis.personalId = cellValue;
                        break;
                    case 3:
                        if (cellValue.length() > 0) {
                            dis.infoDate = sdf.parse(cellValue);
                        }
                        break;
                    case 4:
                        if (cellValue.length() > 0) {
                            dis.disabilityType = Integer.parseInt(cellValue);
                        }
                        break;
                    case 5:
                        if (cellValue.length() > 0) {
                            dis.disabilityResponse = Integer.parseInt(cellValue);
                        }
                        break;
                    case 18:
                        if (cellValue.length() > 0) {
                            dis.created = sdf.parse(cellValue);
                        }
                        break;
                    case 19:
                        if (cellValue.length() > 0) {
                            dis.updated = sdf.parse(cellValue);
                        }
                        break;
                    case 20:
                        dis.userId = cellValue;
                        break;
                    case 22:
                        dis.exportId = cellValue;
                        break;
                    default:
                        break;
                }
            }

            // Insert row data.
            dis.debug();
            dis.insert(conn);
        }
    }

    public void insert(Connection conn) {
        try {
            StringBuffer sb = new StringBuffer();
            sb.append("INSERT INTO disabilities_csv (exportId, disabilitiesId, ");
            sb.append("projectEntryId, personalId, infoDate, disabilityType, ");
            sb.append("disabilityResponse, created, updated, userId) ");
            sb.append("VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");

            PreparedStatement ps = conn.prepareStatement(sb.toString());
            ps.setString(1, SqlString.encode(this.exportId));
            ps.setString(2, SqlString.encode(this.disabilitiesId));
            ps.setString(3, SqlString.encode(this.projectEntryId));
            ps.setString(4, SqlString.encode(this.personalId));
            java.sql.Date sqlDate = new java.sql.Date(this.infoDate.getTime());
            ps.setDate(5, sqlDate);
            ps.setInt(6, this.disabilityType);
            ps.setInt(7, this.disabilityResponse);

            Timestamp ts = new Timestamp(this.created.getTime());
            ps.setTimestamp(8, ts);
            ts = new Timestamp(this.updated.getTime());
            ps.setTimestamp(9, ts);
            ps.setString(10, SqlString.encode(this.userId));

            int out = ps.executeUpdate();
            if (out == 0) {
                log.error("Failed to insert disabilities_csv record!");
            }
        }
        catch (SQLException sqle) {
            log.error("SQLException DisabilitiesCsv.insert(): " + sqle);
        }
        catch (Exception e) {
            log.error("Exception DisabilitiesCsv.insert(): " + e);
        }
    }

    public void debug() {
        log.info("Disablities ID: " + this.disabilitiesId);
        log.info("Enrollment ID: " + this.projectEntryId);
        log.info("Client ID: " + this.personalId);
        log.info("Information Date: " + this.infoDate.toString());
        this.displayDisablityType();
        if (this.disabilityType == 10) {
            log.info("Substance Abuse: " + this.disabilityResponse);
        }
        log.info("Date Created: " + this.created.toString());
        log.info("Date Updated: " + this.updated.toString());
        log.info("User ID: " + this.userId);
        log.info("\n");
    }

    private void displayDisablityType() {
        switch (this.disabilityType) {
            case 5:
                log.info("Disablity Type: Physical disability");
                break;
            case 6:
                log.info("Disablity Type: Developmental disability");
                break;
            case 7:
                log.info("Disablity Type: Chronic health condition");
                break;
            case 8:
                log.info("Disablity Type: HIV/AIDS");
                break;
            case 9:
                log.info("Disablity Type: Mental health problem");
                break;
            case 10:
                log.info("Disablity Type: Substance abuse");
                break;
            default:
                log.info("Disablity Type: Data not collected");
                break;
        }
    }
}
