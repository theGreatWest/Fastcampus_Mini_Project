package org.example;

import org.apache.poi.ss.usermodel.*;
import org.example.util.FindOriginDataFormat;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {
    public static void main(String[] args) {
        try {
            FileInputStream file = new FileInputStream(new File("src/main/resources/excel/example.xlsx"));
            Workbook workbook = WorkbookFactory.create(file);
            // 가상 메모리에 있는 격자모양 저장공간
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                if (row.getPhysicalNumberOfCells() == 0) continue;

                for (Cell cell : row) {
                    Object cellValue = FindOriginDataFormat.findOriginData(cell);
                    // util 패키지로 이동
                    if (cellValue != null) System.out.print(FindOriginDataFormat.findOriginData(cell) + "\t");
                }
                System.out.println();
            }
            System.out.println("엑셀에서 데이터 읽어오기 성공!");
            workbook.close();
            file.close();
        } catch (IOException e) {/* pass */}
    }
}
