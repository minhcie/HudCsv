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

public class ExportCsv {
    static final Logger log = Logger.getLogger(ExportCsv.class.getName());

    public String exportId;
    public Date exportDate;          // Export date.
    public Date exportStartDate;     // User entered start date. 
    public Date exportEndDate;       // User entered end date.
    public String softwareName;
    public String softwareVersion;
    public int exportPeriodType = 4; // Other: selected records.
    public int exportDirective = 1;  // Delta refresh.

    public static void importData(Connection conn, XSSFSheet sheet) throws Exception {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
        SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd");

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
            ExportCsv exp = new ExportCsv();

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
                        exp.exportId = cellValue;
                        break;
                    case 8:
                        if (cellValue.length() > 0) {
                            exp.exportDate = sdf.parse(cellValue);
                        }
                        break;
                    case 9:
                        if (cellValue.length() > 0) {
                            exp.exportStartDate = sdf2.parse(cellValue);
                        }
                        break;
                    case 10:
                        if (cellValue.length() > 0) {
                            exp.exportEndDate = sdf2.parse(cellValue);
                        }
                        break;
                    case 11:
                        exp.softwareName = cellValue;
                        break;
                    case 12:
                        exp.softwareVersion = cellValue;
                        break;
                    default:
                        break;
                }
            }

            // Insert row data.
            exp.debug();
            exp.insert(conn);
        }
    }

    public void insert(Connection conn) {
        try {
            StringBuffer sb = new StringBuffer();
            sb.append("INSERT INTO export_csv (exportId, exportDate, ");
            sb.append("exportStartDate, exportEndDate, softwareName, ");
            sb.append("softwareVersion) ");
            sb.append("VALUES (?, ?, ?, ?, ?, ?)");

            PreparedStatement ps = conn.prepareStatement(sb.toString());
            ps.setString(1, SqlString.encode(this.exportId));
            Timestamp ts = new Timestamp(this.exportDate.getTime());
            ps.setTimestamp(2, ts);
            java.sql.Date sqlDate = new java.sql.Date(this.exportStartDate.getTime());
            ps.setDate(3, sqlDate);
            sqlDate = new java.sql.Date(this.exportEndDate.getTime());
            ps.setDate(4, sqlDate);
            ps.setString(5, SqlString.encode(this.softwareName));
            ps.setString(6, SqlString.encode(this.softwareVersion));

            int out = ps.executeUpdate();
            if (out == 0) {
                log.error("Failed to insert export_csv record!");
            }
        }
        catch (SQLException sqle) {
            log.error("SQLException ExportCsv.insert(): " + sqle);
        }
        catch (Exception e) {
            log.error("Exception ExportCsv.insert(): " + e);
        }
    }

    public void debug() {
        log.info("Export ID: " + this.exportId);
        log.info("Export Date: " + this.exportDate.toString());
        log.info("Start Date: " + this.exportStartDate.toString());
        log.info("End Date: " + this.exportEndDate.toString());
        log.info("Software Name: " + this.softwareName);
        log.info("\n");
    }
}
