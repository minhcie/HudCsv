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

public class ProjectCsv {
    static final Logger log = Logger.getLogger(ProjectCsv.class.getName());

    public String exportId;
    public String projectId;
    public String orgId;
    public String name;
    public Date created;
    public Date updated;
    public String userId;

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
            ProjectCsv project = new ProjectCsv();

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
                        project.projectId = cellValue;
                        break;
                    case 1:
                        project.orgId = cellValue;
                        break;
                    case 2:
                        project.name = cellValue;
                        break;
                    case 10:
                        if (cellValue.length() > 0) {
                            project.created = sdf.parse(cellValue);
                        }
                        break;
                    case 11:
                        if (cellValue.length() > 0) {
                            project.updated = sdf.parse(cellValue);
                        }
                        break;
                    case 12:
                        project.userId = cellValue;
                        break;
                    case 14:
                        project.exportId = cellValue;
                        break;
                    default:
                        break;
                }
            }

            // Insert row data.
            project.debug();
            project.insert(conn);
        }
    }

    public void insert(Connection conn) {
        try {
            StringBuffer sb = new StringBuffer();
            sb.append("INSERT INTO project_csv (exportId, projectId, orgId, ");
            sb.append("name, created, updated, userId) ");
            sb.append("VALUES (?, ?, ?, ?, ?, ?, ?)");

            PreparedStatement ps = conn.prepareStatement(sb.toString());
            ps.setString(1, SqlString.encode(this.exportId));
            ps.setString(2, SqlString.encode(this.projectId));
            ps.setString(3, SqlString.encode(this.orgId));
            ps.setString(4, SqlString.encode(this.name));
            Timestamp ts = new Timestamp(this.created.getTime());
            ps.setTimestamp(5, ts);
            ts = new Timestamp(this.updated.getTime());
            ps.setTimestamp(6, ts);
            ps.setString(7, SqlString.encode(this.userId));

            int out = ps.executeUpdate();
            if (out == 0) {
                log.error("Failed to insert project_csv record!");
            }
        }
        catch (SQLException sqle) {
            log.error("SQLException ProjectCsv.insert(): " + sqle);
        }
        catch (Exception e) {
            log.error("Exception ProjectCsv.insert(): " + e);
        }
    }

    public void debug() {
        log.info("Project ID: " + this.projectId);
        log.info("Organization ID: " + this.orgId);
        log.info("Project Name: " + this.name);
        log.info("Date Created: " + this.created.toString());
        log.info("Date Updated: " + this.updated.toString());
        log.info("User ID: " + this.userId);
        log.info("\n");
    }
}
