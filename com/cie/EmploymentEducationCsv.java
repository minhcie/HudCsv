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

public class EmploymentEducationCsv {
    static final Logger log = Logger.getLogger(EmploymentEducationCsv.class.getName());

    public String exportId;
    public String employmentEducationId;
    public String projectEntryId;
    public String personalId;
    public Date infoDate;
    public int lastGradeCompleted = 99;
    public int employed = 99;
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
            EmploymentEducationCsv empEdu = new EmploymentEducationCsv();

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
                        empEdu.employmentEducationId = cellValue;
                        break;
                    case 1:
                        empEdu.projectEntryId = cellValue;
                        break;
                    case 2:
                        empEdu.personalId = cellValue;
                        break;
                    case 3:
                        if (cellValue.length() > 0) {
                            empEdu.infoDate = sdf.parse(cellValue);
                        }
                        break;
                    case 4:
                        if (cellValue.length() > 0) {
                            empEdu.lastGradeCompleted = Integer.parseInt(cellValue);
                        }
                        break;
                    case 6:
                        if (cellValue.length() > 0) {
                            empEdu.employed = Integer.parseInt(cellValue);
                        }
                        break;
                    case 10:
                        if (cellValue.length() > 0) {
                            empEdu.created = sdf.parse(cellValue);
                        }
                        break;
                    case 11:
                        if (cellValue.length() > 0) {
                            empEdu.updated = sdf.parse(cellValue);
                        }
                        break;
                    case 12:
                        empEdu.userId = cellValue;
                        break;
                    case 14:
                        empEdu.exportId = cellValue;
                        break;
                    default:
                        break;
                }
            }

            // Insert row data.
            empEdu.debug();
            empEdu.insert(conn);
        }
    }

    public void insert(Connection conn) {
        try {
            StringBuffer sb = new StringBuffer();
            sb.append("INSERT INTO employment_education_csv (exportId, employmentEducationId, ");
            sb.append("projectEntryId, personalId, infoDate, lastGradeCompleted, ");
            sb.append("employed, created, updated, userId) ");
            sb.append("VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");

            PreparedStatement ps = conn.prepareStatement(sb.toString());
            ps.setString(1, SqlString.encode(this.exportId));
            ps.setString(2, SqlString.encode(this.employmentEducationId));
            ps.setString(3, SqlString.encode(this.projectEntryId));
            ps.setString(4, SqlString.encode(this.personalId));
            java.sql.Date sqlDate = new java.sql.Date(this.infoDate.getTime());
            ps.setDate(5, sqlDate);
            ps.setInt(6, this.lastGradeCompleted);
            ps.setInt(7, this.employed);

            Timestamp ts = new Timestamp(this.created.getTime());
            ps.setTimestamp(8, ts);
            ts = new Timestamp(this.updated.getTime());
            ps.setTimestamp(9, ts);
            ps.setString(10, SqlString.encode(this.userId));

            int out = ps.executeUpdate();
            if (out == 0) {
                log.error("Failed to insert employment_education_csv record!");
            }
        }
        catch (SQLException sqle) {
            log.error("SQLException EmploymentEducationCsv.insert(): " + sqle);
        }
        catch (Exception e) {
            log.error("Exception EmploymentEducationCsv.insert(): " + e);
        }
    }

    public void debug() {
        log.info("Employment Education ID: " + this.employmentEducationId);
        log.info("Enrollment ID: " + this.projectEntryId);
        log.info("Client ID: " + this.personalId);
        log.info("Information Date: " + this.infoDate.toString());
        this.displayEducation();
        log.info("Employed: " + this.employed);
        log.info("Date Created: " + this.created.toString());
        log.info("Date Updated: " + this.updated.toString());
        log.info("User ID: " + this.userId);
        log.info("\n");
    }

    private void displayEducation() {
        switch (this.lastGradeCompleted) {
            case 1:
                log.info("Last Grade Completed: Less than grade 5");
                break;
            case 2:
                log.info("Last Grade Completed: Grades 5-6");
                break;
            case 3:
                log.info("Last Grade Completed: Grades 7-8");
                break;
            case 4:
                log.info("Last Grade Completed: Grades 9-11");
                break;
            case 5:
                log.info("Last Grade Completed: Grades 12");
                break;
            case 6:
                log.info("Last Grade Completed: School program does not have grade levels");
                break;
            case 7:
                log.info("Last Grade Completed: GED");
                break;
            case 10:
                log.info("Last Grade Completed: Some college");
                break;
            case 8:
                log.info("Last Grade Completed: Client doesn't know");
                break;
            case 9:
                log.info("Last Grade Completed: Client refused");
                break;
            default:
                log.info("Last Grade Completed: Data not collected");
                break;
        }
    }
}
