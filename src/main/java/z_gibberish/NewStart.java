package z_gibberish;

import org.apache.poi.ss.usermodel.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.TreeSet;

public class NewStart {
    public static void main(String[] args) {
//        new NewStart();


    }

    private Workbook workbook;
    private FileOutputStream outputStream;
    public NewStart(){
        try {
            FileInputStream file = new FileInputStream("C:\\sat work\\test\\lab pertindakan new1.xls");
            workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheetAt(0);

//            sheet.shiftColumns(0, sheet.getLastRowNum(), 1);
            // Create new column in A
            Row row = sheet.getRow(0);
            Cell CellAC = row.createCell(28);
            CellAC.setCellValue("NOREG");
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                row = sheet.getRow(i);
                CellAC = row.createCell(28);
                CellAC.setCellFormula("A" + (i + 1) + "&B" + (i + 1) + "&C" + (i + 1) + "&D" + (i + 1) + "&E" + (i + 1));
            }

            // Perform pivot simulation
            Map<String, Integer> pivotData = new HashMap<>();
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row rowP = sheet.getRow(i);
                String p = rowP.getCell(0).getStringCellValue();
                Integer count = pivotData.getOrDefault(p, 0);
                count += (int) rowP.getCell(2).getNumericCellValue();
                pivotData.put(p, count);
            }

            // Write pivot data to sheet
            int rowNum = sheet.getLastRowNum() + 1;
            for (Map.Entry<String, Integer> entry : pivotData.entrySet()) {
                Row rowNew = sheet.createRow(rowNum++);
                rowNew.createCell(0).setCellValue(entry.getKey());
                rowNew.createCell(2).setCellValue(entry.getValue());
            }

            outputStream = new FileOutputStream("lab pertindakan new2.xls");
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
                if (outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
