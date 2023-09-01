package z_gibberish;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class excelChecker {
    public static void main(String[] args) {
        FileOutputStream outputStream = null;
        Workbook workbook = null;

        //letak file
        String fileSource = "C:\\sat work\\2023\\08 agustus\\00. RS Tangerang\\bu kanthi\\";
        //file hasil ditaruh mana?
        String fileOutput = fileSource;
        //nama file
        String namaFile = "LABPK";
        //sheet berapa saja
        Integer[] sheetArray = {
                0
        };
        //data dicari dikolom berapa?
        int cellKolom = 1;
        //karakter apa yang dicari?
        String karakterDicari = "prodia";

        //cek ekstensi file
        File xlsxFileDicari = new File (fileSource + namaFile + ".xlsx");
        File xlsFileDicari = new File (fileSource + namaFile + ".xls");
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
            FileInputStream inputStream = new FileInputStream (fileDicari);
            workbook = WorkbookFactory.create (inputStream);

            for (int sheet : sheetArray) {
//                int lastRow = workbook.getSheetAt (sheet).getLastRowNum ();
//                int lastCell = workbook.getSheetAt (sheet).getRow (0).getLastCellNum ();
//                System.out.println (lastRow + " " + lastCell);
//                for (int i = 0; i <= lastRow; i++) {
//                    Row activeRow = workbook.getSheetAt (sheet).getRow (i);
//                    Cell activeCell = activeRow.getCell (cellKolom);
//                    if (activeCell != null) {
//                        CellType cellType = activeCell.getCellType ();
//                        if (cellType == CellType.STRING && activeCell.getStringCellValue ().toLowerCase ().contains (karakterDicari.toLowerCase ())) {
//                            activeRow.createCell (lastCell).setCellValue (karakterDicari.toLowerCase ());
//                        } else if (cellType == CellType.NUMERIC && activeCell.equals (karakterDicari)) {
//                            activeRow.createCell (lastCell).setCellValue (karakterDicari.toLowerCase ());
//                        }
//                    }
//                }
                Sheet currentSheet = workbook.getSheetAt(sheet);
                int lastRow = currentSheet.getLastRowNum();
                int lastCell = currentSheet.getRow(0).getLastCellNum();
                System.out.println(lastRow + " " + lastCell);

                for (int i = 0; i <= lastRow; i++) {
                    Row activeRow = currentSheet.getRow(i);
                    Cell activeCell = activeRow.getCell(cellKolom);

                    if (activeCell != null) {
                        CellType cellType = activeCell.getCellType();
                        String cellValue = activeCell.getStringCellValue();

                        // Check if the cell value contains the search character(s) ignoring case
                        if (cellType == CellType.STRING && cellValue.toLowerCase().contains(karakterDicari.toLowerCase())) {
                            activeRow.createCell(lastCell).setCellValue(karakterDicari.toLowerCase());
                        } else if (cellType == CellType.NUMERIC && cellValue.equals(karakterDicari)) {
                            activeRow.createCell(lastCell).setCellValue(karakterDicari.toLowerCase());
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace ();
        }
        try {
            outputStream = new FileOutputStream (fileOutput + namaFile + ".xlsx");
            workbook.write (outputStream);
        } catch (IOException e) {
            e.printStackTrace ();
        } finally {
            try {
                if (workbook != null) {
                    workbook.close ();
                }
                if (outputStream != null) {
                    outputStream.close ();
                }
            } catch (IOException e) {
                e.printStackTrace ();
            }
        }

    }
}