import org.apache.poi.ss.usermodel.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class MoveItRightByOne {
    public static void main(String[] args) {
        new MoveItRightByOne();
    }

    private Workbook workbook;
    private FileOutputStream outputStream;

    public int ShiftBy;

    public MoveItRightByOne() {
        try {
            FileInputStream file = null;
            try {
                file = new FileInputStream("C:\\sat work\\test\\lab pertindakan new1.xls");
            } catch (FileNotFoundException e) {
                throw new RuntimeException(e);
            }
            workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheetAt(0);
            int lastColumn = sheet.getRow(0).getLastCellNum();
            int lastRow = sheet.getLastRowNum();
            System.out.println("Last Column: " + lastColumn);
            System.out.println("Last Row: "+lastRow);


            for (int baris = lastRow-1; baris>=0;baris--){
                for (int i = lastColumn-1; i >= 0; i--) {
                    Cell cell = sheet.getRow(baris).getCell(i);
                    if (cell != null) {
                        if (cell.getCellType() == CellType.STRING) {
                            Cell newCell = sheet.getRow(baris).createCell(i + 1);
                            newCell.setCellValue(cell.getStringCellValue());

                        } else if (cell.getCellType() == CellType.NUMERIC) {
                            Cell newCell = sheet.getRow(baris).createCell(i + 1);
                            newCell.setCellValue(cell.getNumericCellValue());

                        }
                    }
                }
//                else {
//                    continue;
//                }
            }

//          dari kiri kekanan incremental 1 /n jika cell mengandung "KD_INST"
            for (int column=1; column<=lastColumn; column++) {
                Cell cell = sheet.getRow(0).getCell(column);
//                System.out.println(cell.getStringCellValue());
//                System.out.println(cell.getStringCellValue().equals("KD_INST"));

//              beri nama A1 "NOREG
                sheet.getRow(0).getCell(0).setCellValue("NOREG");

//              jika cell mengandung "KD_INST" concat jadi noreg
                if (cell.getStringCellValue().equals("KD_INST")) {
                    for (int i = 1; i <= lastRow; i++) {
                        Cell noReg = sheet.getRow(i).createCell(0);
                        noReg.setCellValue(sheet.getRow(i).getCell(column).getStringCellValue() +
                                sheet.getRow(i).getCell(column + 1).getStringCellValue() +
                                sheet.getRow(i).getCell(column + 2).getStringCellValue() +
                                sheet.getRow(i).getCell(column + 3).getStringCellValue() +
                                sheet.getRow(i).getCell(column + 4).getStringCellValue());
//                noReg.setCellFormula("B" + (i+1) + "&C" + (i+1 ) + "&D" + (i+1 ) + "&E" + (i+1 ) + "&F" + (i+1 ));
                    }
                }
            }
            outputStream = new FileOutputStream("lab pertindakan new.xls");
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
