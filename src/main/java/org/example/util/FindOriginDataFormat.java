package org.example.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class FindOriginDataFormat {
    public static Object findOriginData(Cell cell) {
        switch (cell.getCellType()) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date dateValue = cell.getDateCellValue();
                    DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                    return dateFormat.format(dateValue);
                } else {
                    double numericValue = cell.getNumericCellValue();
                    if (numericValue == Math.floor(numericValue)) { // 버림
                        return (int) numericValue;
                    } else {
                        return numericValue;
                    }
                }
            case STRING:
                return cell.getStringCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return null;
            default:
                return null;
        }
    }
//
//    public static Class<?> findObjType(Object obj) {
//        if (obj == null) {
//            return null;
//        }
//        return obj.getClass();  // 객체의 클래스 타입을 반환
//    }
//
//    public static <T> T ObjToFindType(Object value, Class<T> targetType) {
//        if(value==null) return null;
//        return targetType.cast(value);
//    }
}
