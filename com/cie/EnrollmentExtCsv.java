package com.cie;

import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.Serializable;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.log4j.Logger;

public class EnrollmentExtCsv implements Serializable {
    static final long serialVersionUID = 1L;
    static final Logger log = Logger.getLogger(EnrollmentExtCsv.class.getName());

    public String id;
    public String enrollmentId;
    public String clientId;
    public int chronicallyHomeless = 99;
    public Date created;
    public Date updated;
    public String userId;

    @Override
    public String toString() {
        return this.id + " - " + this.enrollmentId + " - " + this.clientId +
               " - " + this.chronicallyHomeless;
    }

    public static Map<String, EnrollmentExtCsv> readData() {
        Map<String, EnrollmentExtCsv> enrollExtMap = new HashMap<String, EnrollmentExtCsv>();
        BufferedReader br = null;
        String line = "";
        String cvsSplitBy = ",";

        try {
            SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");

            log.info("Reading enrollmentext.csv ...\n");
            String csvFileName = "../enrollmentext.csv";
            br = new BufferedReader(new FileReader(csvFileName));
            while ((line = br.readLine()) != null) {
                String[] rec = line.split(cvsSplitBy);
                EnrollmentExtCsv ext = new EnrollmentExtCsv();

                ext.id = rec[0];
                ext.enrollmentId = rec[1];
                ext.clientId = rec[2];

                String temp = rec[3];
                if (temp != null && temp.trim().length() > 0) {
                    ext.chronicallyHomeless = Integer.parseInt(temp);
                }

                temp = rec[4];
                if (temp != null && temp.trim().length() > 0) {
                    ext.created = sdf2.parse(temp);
                }
                temp = rec[5];
                if (temp != null && temp.trim().length() > 0) {
                    ext.updated = sdf2.parse(temp);
                }
                ext.userId = rec[6];

                // Put info in the map.
                enrollExtMap.put(ext.id, ext);
            }
        }
        catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        catch (IOException e) {
            e.printStackTrace();
        }
        catch (Exception e) {
            e.printStackTrace();
        }
        finally {
            if (br != null) {
                try {
                    br.close();
                }
                catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        return enrollExtMap;
    }

    // @debug.
    public void display() {
        log.info("Enrollment Ext ID: " + this.id);
        log.info("Enrollment ID: " + this.enrollmentId);
        log.info("Client ID: " + this.clientId);
        this.displayChronicallyHomeless();
        log.info("Date Created: " + this.created.toString());
        log.info("Date Updated: " + this.updated.toString());
        log.info("User ID: " + this.userId);
        log.info("\n");
    }

    private void displayChronicallyHomeless() {
        switch (this.chronicallyHomeless) {
            case 0:
                log.info("Chronically Homeless: No");
                break;
            case 1:
                log.info("Chronically Homeless: Yes");
                break;
            case 8:
                log.info("Chronically Homeless: Client doesn't know");
                break;
            case 9:
                log.info("Chronically Homeless: Client refused");
                break;
            default:
                log.info("Chronically Homeless: Data not collected");
                break;
        }
    }
}
