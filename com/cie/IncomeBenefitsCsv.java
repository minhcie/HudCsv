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

public class IncomeBenefitsCsv {
    static final Logger log = Logger.getLogger(IncomeBenefitsCsv.class.getName());

    public String exportId;
    public String incomeBenefitsId;
    public String projectEntryId;
    public String personalId;
    public Date infoDate;
    // Income sources.
    public int incomeFromAnySource = 99;
    public double totalMonthlyIncome = 0;
    public int earned = 99;
    public int unemployment = 99;
    public int ssi = 99;
    public int ssdi = 99;
    public int vaDisabilityService = 99;
    public int vaDisabilityNonService = 99;
    public int privateDisability = 99;
    public int workersComp = 99;
    public int tanf = 99;
    public int ga = 99;
    public int ssRetirement = 99;
    public int pension = 99;
    public int childSupport = 99;
    public int alimony = 99;
    public int otherSource = 99;
    public String otherSourceIdentify;
    // Other benefits.
    public int benefitsFromAnySource = 99;
    public int snap = 99;
    public int wic = 99;
    public int tanfChildCare = 99;
    public int tanfTransportation = 99;
    public int otherTanf = 99;
    public int rentalAssistanceOngoing = 99;
    public int rentalAssistanceTemp = 99;
    public int otherBenefits = 99;
    public String otherBenefitsIdentify;
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
            IncomeBenefitsCsv income = new IncomeBenefitsCsv();

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
                        income.incomeBenefitsId = cellValue;
                        break;
                    case 1:
                        income.projectEntryId = cellValue;
                        break;
                    case 2:
                        income.personalId = cellValue;
                        break;
                    case 3:
                        if (cellValue.length() > 0) {
                            income.infoDate = sdf.parse(cellValue);
                        }
                        break;
                    case 4:
                        if (cellValue.length() > 0) {
                            income.incomeFromAnySource = Integer.parseInt(cellValue);
                        }
                        break;
                    case 5:
                        if (cellValue.length() > 0) {
                            income.totalMonthlyIncome = Double.parseDouble(cellValue);
                        }
                        break;
                    case 6:
                        if (cellValue.length() > 0) {
                            income.earned = Integer.parseInt(cellValue);
                        }
                        break;
                    case 8:
                        if (cellValue.length() > 0) {
                            income.unemployment = Integer.parseInt(cellValue);
                        }
                        break;
                    case 10:
                        if (cellValue.length() > 0) {
                            income.ssi = Integer.parseInt(cellValue);
                        }
                        break;
                    case 12:
                        if (cellValue.length() > 0) {
                            income.ssdi = Integer.parseInt(cellValue);
                        }
                        break;
                    case 14:
                        if (cellValue.length() > 0) {
                            income.vaDisabilityService = Integer.parseInt(cellValue);
                        }
                        break;
                    case 16:
                        if (cellValue.length() > 0) {
                            income.vaDisabilityNonService = Integer.parseInt(cellValue);
                        }
                        break;
                    case 18:
                        if (cellValue.length() > 0) {
                            income.privateDisability = Integer.parseInt(cellValue);
                        }
                        break;
                    case 20:
                        if (cellValue.length() > 0) {
                            income.workersComp = Integer.parseInt(cellValue);
                        }
                        break;
                    case 22:
                        if (cellValue.length() > 0) {
                            income.tanf = Integer.parseInt(cellValue);
                        }
                        break;
                    case 24:
                        if (cellValue.length() > 0) {
                            income.ga = Integer.parseInt(cellValue);
                        }
                        break;
                    case 26:
                        if (cellValue.length() > 0) {
                            income.ssRetirement = Integer.parseInt(cellValue);
                        }
                        break;
                    case 28:
                        if (cellValue.length() > 0) {
                            income.pension = Integer.parseInt(cellValue);
                        }
                        break;
                    case 30:
                        if (cellValue.length() > 0) {
                            income.childSupport = Integer.parseInt(cellValue);
                        }
                        break;
                    case 32:
                        if (cellValue.length() > 0) {
                            income.alimony = Integer.parseInt(cellValue);
                        }
                        break;
                    case 34:
                        if (cellValue.length() > 0) {
                            income.otherSource = Integer.parseInt(cellValue);
                        }
                        break;
                    case 36:
                        income.otherSourceIdentify = cellValue;
                        break;
                    case 37:
                        if (cellValue.length() > 0) {
                            income.benefitsFromAnySource = Integer.parseInt(cellValue);
                        }
                        break;
                    case 38:
                        if (cellValue.length() > 0) {
                            income.snap = Integer.parseInt(cellValue);
                        }
                        break;
                    case 39:
                        if (cellValue.length() > 0) {
                            income.wic = Integer.parseInt(cellValue);
                        }
                        break;
                    case 40:
                        if (cellValue.length() > 0) {
                            income.tanfChildCare = Integer.parseInt(cellValue);
                        }
                        break;
                    case 41:
                        if (cellValue.length() > 0) {
                            income.tanfTransportation = Integer.parseInt(cellValue);
                        }
                        break;
                    case 42:
                        if (cellValue.length() > 0) {
                            income.otherTanf = Integer.parseInt(cellValue);
                        }
                        break;
                    case 43:
                        if (cellValue.length() > 0) {
                            income.rentalAssistanceOngoing = Integer.parseInt(cellValue);
                        }
                        break;
                    case 44:
                        if (cellValue.length() > 0) {
                            income.rentalAssistanceTemp = Integer.parseInt(cellValue);
                        }
                        break;
                    case 45:
                        if (cellValue.length() > 0) {
                            income.otherBenefits = Integer.parseInt(cellValue);
                        }
                        break;
                    case 46:
                        income.otherBenefitsIdentify = cellValue;
                        break;
                    case 69:
                        if (cellValue.length() > 0) {
                            income.created = sdf2.parse(cellValue);
                        }
                        break;
                    case 70:
                        if (cellValue.length() > 0) {
                            income.updated = sdf2.parse(cellValue);
                        }
                        break;
                    case 71:
                        income.userId = cellValue;
                        break;
                    case 73:
                        income.exportId = cellValue;
                        break;
                    default:
                        break;
                }
            }

            // Insert row data.
            income.debug();
            income.insert(conn);
        }
    }

    public void insert(Connection conn) {
        try {
            StringBuffer sb = new StringBuffer();
            sb.append("INSERT INTO income_benefits_csv (exportId, incomeBenefitsId, ");
            sb.append("projectEntryId, personalId, infoDate, incomeFromAnySource, ");
            sb.append("totalMonthlyIncome, earned, unemployment, ssi, ssdi, ");
            sb.append("vaDisabilityService, vaDisabilityNonService, privateDisability, ");
            sb.append("workersComp, tanf, ga, ssRetirement, pension, childSupport, ");
            sb.append("alimony, otherSource, otherSourceIdentify, benefitsFromAnySource, ");
            sb.append("snap, wic, tanfChildCare, tanfTransportation, otherTanf, ");
            sb.append("rentalAssistanceOngoing, rentalAssistanceTemp, otherBenefits, ");
            sb.append("otherBenefitsIdentify, created, updated, userId) ");
            sb.append("VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");

            PreparedStatement ps = conn.prepareStatement(sb.toString());
            ps.setString(1, SqlString.encode(this.exportId));
            ps.setString(2, SqlString.encode(this.incomeBenefitsId));
            ps.setString(3, SqlString.encode(this.projectEntryId));
            ps.setString(4, SqlString.encode(this.personalId));
            java.sql.Date sqlDate = new java.sql.Date(this.infoDate.getTime());
            ps.setDate(5, sqlDate);

            ps.setInt(6, this.incomeFromAnySource);
            ps.setDouble(7, this.totalMonthlyIncome);
            ps.setInt(8, this.earned);
            ps.setInt(9, this.unemployment);
            ps.setInt(10, this.ssi);
            ps.setInt(11, this.ssdi);
            ps.setInt(12, this.vaDisabilityService);
            ps.setInt(13, this.vaDisabilityNonService);
            ps.setInt(14, this.privateDisability);
            ps.setInt(15, this.workersComp);
            ps.setInt(16, this.tanf);
            ps.setInt(17, this.ga);
            ps.setInt(18, this.ssRetirement);
            ps.setInt(19, this.pension);
            ps.setInt(20, this.childSupport);
            ps.setInt(21, this.alimony);
            ps.setInt(22, this.otherSource);
            ps.setString(23, SqlString.encode(this.otherSourceIdentify));
            ps.setInt(24, this.benefitsFromAnySource);
            ps.setInt(25, this.snap);
            ps.setInt(26, this.wic);
            ps.setInt(27, this.tanfChildCare);
            ps.setInt(28, this.tanfTransportation);
            ps.setInt(29, this.otherTanf);
            ps.setInt(30, this.rentalAssistanceOngoing);
            ps.setInt(31, this.rentalAssistanceTemp);
            ps.setInt(32, this.otherBenefits);
            ps.setString(33, SqlString.encode(this.otherBenefitsIdentify));

            Timestamp ts = new Timestamp(this.created.getTime());
            ps.setTimestamp(34, ts);
            ts = new Timestamp(this.updated.getTime());
            ps.setTimestamp(35, ts);
            ps.setString(36, SqlString.encode(this.userId));

            int out = ps.executeUpdate();
            if (out == 0) {
                log.error("Failed to insert income_benefits_csv record!");
            }
        }
        catch (SQLException sqle) {
            log.error("SQLException IncomeBenefitsCsv.insert(): " + sqle);
        }
        catch (Exception e) {
            log.error("Exception IncomeBenefitsCsv.insert(): " + e);
        }
    }

    public void debug() {
        log.info("Income Benefits ID: " + this.incomeBenefitsId);
        log.info("Enrollment ID: " + this.projectEntryId);
        log.info("Client ID: " + this.personalId);
        log.info("Information Date: " + this.infoDate.toString());
        this.displayFromAnySource(this.incomeFromAnySource, "Income");
        this.displayIncomeSources();
        this.displayFromAnySource(this.benefitsFromAnySource, "Benefits");
        this.displayOtherBenefits();
        log.info("Date Created: " + this.created.toString());
        log.info("Date Updated: " + this.updated.toString());
        log.info("User ID: " + this.userId);
        log.info("\n");
    }

    private void displayFromAnySource(int source, String text) {
        switch (source) {
            case 0:
                log.info(text + " From Any Source: No");
                break;
            case 1:
                log.info(text + " From Any Source: Yes");
                break;
            case 8:
                log.info(text + " From Any Source: Client doesn't know");
                break;
            case 9:
                log.info(text + " From Any Source: Client refused");
                break;
            default:
                log.info(text + " From Any Source: Data not collected");
                break;
        }
    }

    private void displayIncomeSources() {
        String s = "";
        if (this.earned == 1)
            s += "Earned; ";
        if (this.unemployment == 1)
            s += "Unemployment; ";
        if (this.ssi == 1)
            s += "SSI; ";
        if (this.vaDisabilityService == 1)
            s += "VA Disability Service; ";
        if (this.vaDisabilityNonService == 1)
            s += "VA Disability Non Service; ";
        if (this.privateDisability == 1)
            s += "Private Disability; ";
        if (this.workersComp == 1)
            s += "Workers Comp; ";
        if (this.tanf == 1)
            s += "TANF; ";
        if (this.ga == 1)
            s += "GA; ";
        if (this.ssRetirement == 1)
            s += "Social Security Retirement; ";
        if (this.pension == 1)
            s += "Pension; ";
        if (this.childSupport == 1)
            s += "Child Support; ";
        if (this.alimony == 1)
            s += "Alimony";
        if (this.otherSource == 1)
            s += otherSourceIdentify;
        log.info(s);
    }

    private void displayOtherBenefits() {
        String s = "";
        if (this.snap == 1)
            s += "SNAP; ";
        if (this.wic == 1)
            s += "WIC; ";
        if (this.tanfChildCare == 1)
            s += "TANF Child Care; ";
        if (this. tanfTransportation == 1)
            s += "TANF Transportation; ";
        if (this.otherTanf == 1)
            s += "Other TANF; ";
        if (this.rentalAssistanceOngoing == 1)
            s += "Rental Assistance Ongoing; ";
        if (this.rentalAssistanceTemp == 1)
            s += "Rental Assistance Temp; ";
        if (this.otherBenefits == 1)
            s += otherBenefitsIdentify;
        log.info(s);
    }
}
