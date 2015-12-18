package com.cie;

import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;

public class ExcelUtils {
    public static String getCellValue(Cell cell) {
        String retVal = "";
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                retVal = "" + cell.getBooleanCellValue();
                break;
 
            case Cell.CELL_TYPE_STRING:
                retVal = cell.getStringCellValue();
                break;
 
            case Cell.CELL_TYPE_NUMERIC:
                retVal = isNumberOrDate(cell);
                break;
 
            case Cell.CELL_TYPE_BLANK:
            default:
                retVal = "";
        }
        return retVal.trim();
    }

    private static String isNumberOrDate(Cell cell) {
        String retVal;
        if (HSSFDateUtil.isCellDateFormatted(cell)) {
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
            retVal = sdf.format(cell.getDateCellValue());
        }
        else {
            DataFormatter formatter = new DataFormatter();
            retVal = formatter.formatCellValue(cell);
        }
        return retVal;
    }    
}