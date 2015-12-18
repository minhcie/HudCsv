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

public class CaseManagerCsv {
    static final Logger log = Logger.getLogger(CaseManagerCsv.class.getName());

    public String exportId;
    public String caseManagerId;
    public String projectId;
    public String personalId;
    public String name;
    public String email;
    public String phone;
    public Date startDate;
    public Date endDate;
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
            CaseManagerCsv mgr = new CaseManagerCsv();

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
                        mgr.caseManagerId = cellValue;
                        break;
                    case 1:
                        mgr.projectId = cellValue;
                        break;
                    case 2:
                        mgr.personalId = cellValue;
                        break;
                    case 3:
                        mgr.name = cellValue;
                        break;
                    case 4:
                        mgr.email = cellValue;
                        break;
                    case 5:
                        mgr.phone = cellValue;
                        break;
                    case 6:
                        if (cellValue.length() > 0) {
                            mgr.startDate = sdf.parse(cellValue);
                        }
                        break;
                    case 7:
                        if (cellValue.length() > 0) {
                            mgr.endDate = sdf.parse(cellValue);
                        }
                        break;
                    case 8:
                        if (cellValue.length() > 0) {
                            mgr.created = sdf2.parse(cellValue);
                        }
                        break;
                    case 9:
                        if (cellValue.length() > 0) {
                            mgr.updated = sdf2.parse(cellValue);
                        }
                        break;
                    case 10:
                        mgr.userId = cellValue;
                        break;
                    case 12:
                        mgr.exportId = cellValue;
                        break;
                    default:
                        break;
                }
            }

            // Insert row data.
            mgr.debug();
            mgr.insert(conn);
        }
    }

    public void insert(Connection conn) {
        try {
            StringBuffer sb = new StringBuffer();
            sb.append("INSERT INTO case_manager_csv (exportId, caseManagerId, ");
            sb.append("projectId, personalId, name, email, phone, startDate, ");
            sb.append("endDate, created, updated, userId) ");
            sb.append("VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");

            PreparedStatement ps = conn.prepareStatement(sb.toString());
            ps.setString(1, SqlString.encode(this.exportId));
            ps.setString(2, SqlString.encode(this.caseManagerId));
            ps.setString(3, SqlString.encode(this.projectId));
            ps.setString(4, SqlString.encode(this.personalId));
            ps.setString(5, SqlString.encode(this.name));
            ps.setString(6, SqlString.encode(this.email));
            ps.setString(7, SqlString.encode(this.phone));

            java.sql.Date sqlDate = new java.sql.Date(this.startDate.getTime());
            ps.setDate(8, sqlDate);
            if (this.endDate != null) {
                sqlDate = new java.sql.Date(this.endDate.getTime());
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
                log.error("Failed to insert case_manager_csv record!");
            }
        }
        catch (SQLException sqle) {
            log.error("SQLException CaseManagerCsv.insert(): " + sqle);
        }
        catch (Exception e) {
            log.error("Exception CaseManagerCsv.insert(): " + e);
        }
    }

    public void debug() {
        log.info("Case Manager ID: " + this.caseManagerId);
        log.info("Enrollment ID: " + this.projectId);
        log.info("Client ID: " + this.personalId);
        log.info("Case Manager Name: " + this.name);
        log.info("Case Manager Email: " + this.email);
        log.info("Case Manager Phone: " + this.phone);
        if (this.startDate != null) {
            log.info("Start Date: " + this.startDate.toString());
        }
        if (this.endDate != null) {
            log.info("End Date: " + this.endDate.toString());
        }
        log.info("Date Created: " + this.created.toString());
        log.info("Date Updated: " + this.updated.toString());
        log.info("User ID: " + this.userId);
        log.info("\n");
    }
}
