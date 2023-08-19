package z_gibberish;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;

public class excelChecker {
    public static void main(String[] args) {
        FileOutputStream outputStream = null;
        Workbook workbook = null;

        String fileSource = "C:\\sat work\\2023\\08 agustus\\00. RS Tangerang\\list pasien eca\\";
//        String fileOutput = "C:\\sat work\\test\\";
        String fileOutput = fileSource;
        String namaFile = "all";
        Integer[] sheetArray={
          0,1
        };
//        int sheet = 1;
        int cellKolom = 0;
        String karakterDicari = "_name";

        File xlsxFileDicari = new File (fileSource  + namaFile   +".xlsx");
        File xlsFileDicari = new File (fileSource   + namaFile   +".xls");
        File fileDicari;

        if (xlsxFileDicari.exists ()) {
            fileDicari = xlsxFileDicari;
        } else if (xlsFileDicari.exists ()) {
            fileDicari = xlsFileDicari;
        } else {
            System.out.println ("File not found: " + fileSource + namaFile);
            return;
        }

        try {
//            LocalDateTime start = LocalDateTime.now();
            FileInputStream inputStream  = new FileInputStream(fileDicari);
            workbook = WorkbookFactory.create(inputStream);

            for (int sheet : sheetArray) {
                int lastRow = workbook.getSheetAt (sheet).getLastRowNum ();
                int lastCell = workbook.getSheetAt (sheet).getRow (0).getLastCellNum ();
                for (int i = 0; i <= lastRow; i++) {
                    Row activeRow = workbook.getSheetAt (sheet).getRow (i);
                    Cell activeCell = activeRow.getCell (cellKolom);
                    if (activeCell != null) {
                        CellType cellType = activeCell.getCellType ();
                        if (cellType == CellType.STRING && activeCell.getStringCellValue ().toLowerCase ().contains (karakterDicari.toLowerCase ())) {
                            activeRow.createCell (lastCell).setCellValue (karakterDicari.toLowerCase ());
                        } else if (cellType == CellType.NUMERIC && activeCell.equals (karakterDicari)) {
                            activeRow.createCell (lastCell).setCellValue (karakterDicari.toLowerCase ());
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        try {
            outputStream = new FileOutputStream(fileOutput+namaFile+".xlsx");
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