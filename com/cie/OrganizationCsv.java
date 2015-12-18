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

public class OrganizationCsv {
    static final Logger log = Logger.getLogger(OrganizationCsv.class.getName());

    public String orgId;
    public String name;
    public Date created;
    public Date updated;
    public String userId;
    public String exportId;

    public static void importData(Connection conn, XSSFSheet sheet) throws Exception {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");

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
            OrganizationCsv org = new OrganizationCsv();

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
                        org.orgId = cellValue;
                        break;
                    case 1:
                        org.name = cellValue;
                        break;
                    case 3:
                        if (cellValue.length() > 0) {
                            org.created = sdf.parse(cellValue);
                        }
                        break;
                    case 4:
                        if (cellValue.length() > 0) {
                            org.updated = sdf.parse(cellValue);
                        }
                        break;
                    case 5:
                        org.userId = cellValue;
                        break;
                    case 7:
                        org.exportId = cellValue;
                        break;
                    default:
                        break;
                }
            }

            // Insert row data.
            org.debug();
            org.insert(conn);
        }
    }

    public void insert(Connection conn) {
        try {
            StringBuffer sb = new StringBuffer();
            sb.append("INSERT INTO organization_csv (exportId, orgId, name, ");
            sb.append("created, updated, userId) ");
            sb.append("VALUES (?, ?, ?, ?, ?, ?)");

            PreparedStatement ps = conn.prepareStatement(sb.toString());
            ps.setString(1, SqlString.encode(this.exportId));
            ps.setString(2, SqlString.encode(this.orgId));
            ps.setString(3, SqlString.encode(this.name));
            Timestamp ts = new Timestamp(this.created.getTime());
            ps.setTimestamp(4, ts);
            ts = new Timestamp(this.updated.getTime());
            ps.setTimestamp(5, ts);
            ps.setString(6, SqlString.encode(this.userId));

            int out = ps.executeUpdate();
            if (out == 0) {
                log.error("Failed to insert organization_csv record!");
            }
        }
        catch (SQLException sqle) {
            log.error("SQLException OrganizationCsv.insert(): " + sqle);
        }
        catch (Exception e) {
            log.error("Exception OrganizationCsv.insert(): " + e);
        }
    }

    public void debug() {
        log.info("Organization ID: " + this.orgId);
        log.info("Organization Name: " + this.name);
        log.info("Date Created: " + this.created.toString());
        log.info("Date Updated: " + this.updated.toString());
        log.info("User ID: " + this.userId);
        log.info("\n");
    }
}
