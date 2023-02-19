package z_gibberish;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class PivotSimulation {
    public static void main(String[] args) throws IOException {
        // Input data
        String[] columns = {"A", "B", "C"};
        Object[][] data = {
                {"P", "Q", 10},
                {"P", "R", 20},
                {"Q", "S", 30},
                {"Q", "T", 40}
        };

        // Create workbook
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Pivot Simulation");

        // Create header
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
        }

        // Write data to sheet
        for (int i = 0; i < data.length; i++) {
            Row row = sheet.createRow(i + 1);
            for (int j = 0; j < data[i].length; j++) {
                Cell cell = row.createCell(j);
                if (data[i][j] instanceof String) {
                    cell.setCellValue((String) data[i][j]);
                } else if (data[i][j] instanceof Integer) {
                    cell.setCellValue((Integer) data[i][j]);
                }
            }
        }

        // Perform pivot simulation
        Map<String, Integer> pivotData = new HashMap<>();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String p = row.getCell(0).getStringCellValue();
            Integer count = pivotData.getOrDefault(p, 0);
            count += (int) row.getCell(2).getNumericCellValue();
            pivotData.put(p, count);
        }

        // Write pivot data to sheet
        int rowNum = sheet.getLastRowNum() + 1;
        for (Map.Entry<String, Integer> entry : pivotData.entrySet()) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(entry.getKey());
            row.createCell(2).setCellValue(entry.getValue());
        }

        // Save workbook
        FileOutputStream outputStream = new FileOutputStream("pivot_simulation.xlsx");
        workbook.write(outputStream);
        workbook.close();
    }
}
