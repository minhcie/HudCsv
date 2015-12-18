package com.cie;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import org.apache.log4j.Logger;

public class ClientCsv {
    static final Logger log = Logger.getLogger(ClientCsv.class.getName());

    public String personalId;
    public String firstName;
    public String middleName;
    public String lastName;
    public int nameDataQuality = 99;
    public String ssn;
    public int ssnDataQuality = 99;
    public Date dob;
    public int dobDataQuality = 99;
    public int americanIndian = 0;
    public int asian = 0;
    public int black = 0;
    public int nativeHawaiian = 0;
    public int white = 0;
    public int raceNone = 99;
    public int ethnicity = 99;
    public int gender = 99;
    public String otherGender;
    public int veteranStatus = 99;
    public int militaryBranch = 99;
    public int dischargeStatus = 99;
    public Date created;
    public Date updated;
    public String userId;
    public String exportId;

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
            ClientCsv client = new ClientCsv();

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
                        client.personalId = cellValue;
                        break;
                    case 1:
                        client.firstName = cellValue;
                        break;
                    case 2:
                        client.middleName = cellValue;
                        break;
                    case 3:
                        client.lastName = cellValue;
                        break;
                    case 5:
                        if (cellValue.length() > 0) {
                            client.nameDataQuality = Integer.parseInt(cellValue);
                        }
                        break;
                    case 6:
                        client.ssn = cellValue;
                        break;
                    case 7:
                        if (cellValue.length() > 0) {
                            client.ssnDataQuality = Integer.parseInt(cellValue);
                        }
                        break;
                    case 8:
                        if (cellValue.length() > 0) {
                            client.dob = sdf.parse(cellValue);
                        }
                        break;
                    case 9:
                        if (cellValue.length() > 0) {
                            client.dobDataQuality = Integer.parseInt(cellValue);
                        }
                        break;
                    case 10:
                        if (cellValue.length() > 0) {
                            client.americanIndian = Integer.parseInt(cellValue);
                        }
                        break;
                    case 11:
                        if (cellValue.length() > 0) {
                            client.asian = Integer.parseInt(cellValue);
                        }
                        break;
                    case 12:
                        if (cellValue.length() > 0) {
                            client.black = Integer.parseInt(cellValue);
                        }
                        break;
                    case 13:
                        if (cellValue.length() > 0) {
                            client.nativeHawaiian = Integer.parseInt(cellValue);
                        }
                        break;
                    case 14:
                        if (cellValue.length() > 0) {
                            client.white = Integer.parseInt(cellValue);
                        }
                        break;
                    case 15:
                        if (cellValue.length() > 0) {
                            client.raceNone = Integer.parseInt(cellValue);
                        }
                        break;
                    case 16:
                        if (cellValue.length() > 0) {
                            client.ethnicity = Integer.parseInt(cellValue);
                        }
                        break;
                    case 17:
                        if (cellValue.length() > 0) {
                            client.gender = Integer.parseInt(cellValue);
                        }
                        break;
                    case 18:
                        client.otherGender = cellValue;
                        break;
                    case 19:
                        if (cellValue.length() > 0) {
                            client.veteranStatus = Integer.parseInt(cellValue);
                        }
                        break;
                    case 30:
                        if (cellValue.length() > 0) {
                            client.militaryBranch = Integer.parseInt(cellValue);
                        }
                        break;
                    case 31:
                        if (cellValue.length() > 0) {
                            client.dischargeStatus = Integer.parseInt(cellValue);
                        }
                        break;
                    case 32:
                        if (cellValue.length() > 0) {
                            client.created = sdf.parse(cellValue);
                        }
                        break;
                    case 33:
                        if (cellValue.length() > 0) {
                            client.updated = sdf.parse(cellValue);
                        }
                        break;
                    case 34:
                        client.userId = cellValue;
                        break;
                    case 36:
                        client.exportId = cellValue;
                        break;
                    default:
                        break;
                }
            }

            // Insert row data.
            client.display();
            client.insert(conn);
        }
    }

    public void insert(Connection conn) {
        try {
            StringBuffer sb = new StringBuffer();
            sb.append("INSERT INTO client_csv (exportId, personalId, firstName, ");
            sb.append("middleName, lastName, nameDataQuality, ssn, ssnDataQuality, ");
            sb.append("dob, dobDataQuality, americanIndian, asian, black, nativeHawaiian, ");
            sb.append("white, raceNone, ethnicity, gender, otherGender, veteranStatus, ");
            sb.append("militaryBranch, dischargeStatus, created, updated, userId) ");
            sb.append("VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");

            PreparedStatement ps = conn.prepareStatement(sb.toString());
            ps.setString(1, SqlString.encode(this.exportId));
            ps.setString(2, SqlString.encode(this.personalId));
            ps.setString(3, SqlString.encode(this.firstName));
            ps.setString(4, SqlString.encode(this.middleName));
            ps.setString(5, SqlString.encode(this.lastName));
            ps.setInt(6, this.nameDataQuality);
            ps.setString(7, SqlString.encode(this.ssn));
            ps.setInt(8, this.ssnDataQuality);
            if (this.dob != null) {
                java.sql.Date sqlDate = new java.sql.Date(this.dob.getTime());
                ps.setDate(9, sqlDate);
            }
            else {
                ps.setNull(9, java.sql.Types.NULL);
            }
            ps.setInt(10, this.dobDataQuality);
            ps.setInt(11, this.americanIndian);
            ps.setInt(12, this.asian);
            ps.setInt(13, this.black);
            ps.setInt(14, this.nativeHawaiian);
            ps.setInt(15, this.white);
            ps.setInt(16, this.raceNone);
            ps.setInt(17, this.ethnicity);
            ps.setInt(18, this.gender);
            ps.setString(19, SqlString.encode(this.otherGender));
            ps.setInt(20, this.veteranStatus);
            ps.setInt(21, this.militaryBranch);
            ps.setInt(22, this.dischargeStatus);

            Timestamp ts = new Timestamp(this.created.getTime());
            ps.setTimestamp(23, ts);
            ts = new Timestamp(this.updated.getTime());
            ps.setTimestamp(24, ts);
            ps.setString(25, SqlString.encode(this.userId));

            int out = ps.executeUpdate();
            if (out == 0) {
                log.error("Failed to insert client_csv record!");
            }
        }
        catch (SQLException sqle) {
            log.error("SQLException ClientCsv.insert(): " + sqle);
        }
        catch (Exception e) {
            log.error("Exception ClientCsv.insert(): " + e);
        }
    }

    public static List<ClientCsv> load(Connection conn) {
        List<ClientCsv> results = new ArrayList<ClientCsv>();
        try {
            StringBuffer sb = new StringBuffer();
            sb.append("SELECT exportId, personalId, firstName, middleName, ");
            sb.append("lastName, nameDataQuality, ssn, ssnDataQuality, dob, ");
            sb.append("dobDataQuality, americanIndian, asian, black, nativeHawaiian, ");
            sb.append("white, raceNone, ethnicity, gender, otherGender, veteranStatus, ");
            sb.append("militaryBranch, dischargeStatus, created, updated, userId ");
            sb.append("FROM client_csv ORDER BY id");

            Statement statement = conn.createStatement();
            ResultSet rs = statement.executeQuery(sb.toString());
            while (rs.next()) {
                ClientCsv c = new ClientCsv();
                c.exportId = rs.getString("exportId");
                c.personalId = rs.getString("personalId");
                c.firstName = rs.getString("firstName");
                c.middleName = rs.getString("middleName");
                c.lastName = rs.getString("lastName");
                c.nameDataQuality = rs.getInt("nameDataQuality");
                c.ssn = rs.getString("ssn");
                c.ssnDataQuality = rs.getInt("ssnDataQuality");
                c.dob = rs.getDate("dob");
                c.dobDataQuality = rs.getInt("dobDataQuality");
                c.americanIndian = rs.getInt("americanIndian");
                c.asian = rs.getInt("asian");
                c.black = rs.getInt("black");
                c.nativeHawaiian = rs.getInt("nativeHawaiian");
                c.white = rs.getInt("white");
                c.raceNone = rs.getInt("raceNone");
                c.ethnicity = rs.getInt("ethnicity");
                c.gender = rs.getInt("gender");
                c.otherGender = rs.getString("otherGender");
                c.militaryBranch = rs.getInt("militaryBranch");
                c.dischargeStatus = rs.getInt("dischargeStatus");
                c.created = rs.getDate("created");
                c.updated = rs.getDate("updated");
                c.userId = rs.getString("userId");
                results.add(c);
            }

            rs.close();
            statement.close();
        }
        catch (SQLException sqle) {
            log.error("SQLException ClientCsv.load(): " + sqle);
        }
        catch (Exception e) {
            log.error("Exception ClientCsv.load(): " + e);
        }
        return results;
    }

    public void display() {
        log.info("Client ID: " + this.personalId);
        log.info("First Name: " + this.firstName);
        log.info("Middle Name: " + this.middleName);
        log.info("Last Name: " + this.lastName);
        this.displayDataQuality(this.nameDataQuality, "name");
        log.info("SSN: " + this.ssn);
        this.displayDataQuality(this.ssnDataQuality, "SSN");
        if (this.dob != null) {
            log.info("DOB: " + this.dob.toString());
        }
        this.displayDataQuality(this.dobDataQuality, "DOB");
        log.info("American Indian or Alaska Native: " + this.americanIndian);
        log.info("Asian: " + this.asian);
        log.info("Black or African American: " + this.black);
        log.info("Native Hawaiian or Other Pacific Islander: " + this.nativeHawaiian);
        log.info("White: " + this.white);
        log.info("Race None: " + this.raceNone);
        log.info("Ethnicity: " + this.ethnicity);
        this.displayEthnicity();
        this.displayGender();
        log.info("Other Gender: " + this.otherGender);
        this.displayVeteranStatus();
        log.info("Military Branch: " + this.militaryBranch);
        this.displayMilitaryBranch();
        log.info("Discharge Status: " + this.dischargeStatus);
        this.displayDischargeStatus();
        log.info("Date Created: " + this.created.toString());
        log.info("Date Updated: " + this.updated.toString());
        log.info("User ID: " + this.userId);
        log.info("\n");
    }

    private void displayDataQuality(int dataQuality, String text) {
        switch (dataQuality) {
            case 1:
                log.info("Data Quality: Full " + text + " reported");
                break;
            case 2:
                log.info("Data Quality: Partial " + text + " reported");
                break;
            case 8:
                log.info("Data Quality: Client doesn't know");
                break;
            case 9:
                log.info("Data Quality: Client refused");
                break;
            default:
                log.info("Data Quality: Data not collected");
                break;
        }
    }

    private void displayEthnicity() {
        switch (this.ethnicity) {
            case 0:
                log.info("Ethnicity: Non-Hispanic/Non-Latino");
                break;
            case 1:
                log.info("Ethnicity: Hispanic/Latino");
                break;
            case 8:
                log.info("Ethnicity: Client doesn't know");
                break;
            case 9:
                log.info("Ethnicity: Client refused");
                break;
            default:
                log.info("Ethnicity: Data not collected");
                break;
        }
    }

    private void displayGender() {
        switch (this.gender) {
            case 0:
                log.info("Gender: Female");
                break;
            case 1:
                log.info("Gender: Male");
                break;
            case 2:
                log.info("Gender: Transgender male to female");
                break;
            case 3:
                log.info("Gender: Transgender female to male");
                break;
            case 4:
                log.info("Gender: Other");
                break;
            case 8:
                log.info("Gender: Client doesn't know");
                break;
            case 9:
                log.info("Gender: Client refused");
                break;
            default:
                log.info("Gender: Data not collected");
                break;
        }
    }

    private void displayVeteranStatus() {
        switch (this.veteranStatus) {
            case 0:
                log.info("Military Branch: No");
                break;
            case 1:
                log.info("Military Branch: Yes");
                break;
            case 8:
                log.info("Veteran Status: Client doesn't know");
                break;
            case 9:
                log.info("Veteran Status: Client refused");
                break;
            default:
                log.info("Veteran Status: Data not collected");
                break;
        }
    }

    private void displayMilitaryBranch() {
        switch (this.militaryBranch) {
            case 1:
                log.info("Military Branch: Army");
                break;
            case 2:
                log.info("Military Branch: Air Force");
                break;
            case 3:
                log.info("Military Branch: Navy");
                break;
            case 4:
                log.info("Military Branch: Marines");
                break;
            case 6:
                log.info("Military Branch: Coast Guard");
                break;
            case 8:
                log.info("Military Branch: Client doesn't know");
                break;
            case 9:
                log.info("Military Branch: Client refused");
                break;
            default:
                log.info("Military Branch: Data not collected");
                break;
        }
    }

    private void displayDischargeStatus() {
        switch (this.militaryBranch) {
            case 1:
                log.info("Discharge Status: Honorable");
                break;
            case 2:
                log.info("Discharge Status: General under honorable conditions");
                break;
            case 4:
                log.info("Discharge Status: Bad conduct");
                break;
            case 5:
                log.info("Discharge Status: Dishonorable");
                break;
            case 6:
                log.info("Discharge Status: Under other than honorable conditions (OTH)");
                break;
            case 7:
                log.info("Discharge Status: Uncharacterized");
                break;
            case 8:
                log.info("Discharge Status: Client doesn't know");
                break;
            case 9:
                log.info("Discharge Status: Client refused");
                break;
            default:
                log.info("Discharge Status: Data not collected");
                break;
        }
    }
}
