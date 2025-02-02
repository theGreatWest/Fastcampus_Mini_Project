package org.example.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class FindOriginDataFormat {
    public static Object start(Cell cell){
        switch (cell.getCellType()) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date dateValue = cell.getDateCellValue();
                    DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                    String formattedDate = dateFormat.format(dateValue);
                    return formattedDate;
                } else {
                    double numericValue = cell.getNumericCellValue();
                    if (numericValue == Math.floor(numericValue)) { // 버림
                        int intValue = (int) numericValue;
                        return intValue;
                    } else {
                        return numericValue;
                    }
                }
            case STRING:
                String stringValue = cell.getStringCellValue();
                return stringValue;
            case BOOLEAN:
                boolean booleanValue = cell.getBooleanCellValue();
                return booleanValue;
            case FORMULA:
                String formulaValue = cell.getCellFormula();
                return formulaValue;
            case BLANK:
                return null;
            default:
                return -1;
        }
    }
}
