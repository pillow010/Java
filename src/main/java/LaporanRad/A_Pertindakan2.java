package main.java.LaporanRad;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class A_Pertindakan2 {
    public static void main(String[] args) {
        new A_Pertindakan2();
    }
    private Workbook BookPertindakanNew;

    private FileOutputStream outputStream;

    public A_Pertindakan2(){
        Sheet SheetA = null;
        Sheet sheetB = null;
        File pertindakanNew = new File("C:\\sat work\\test\\rad pertindakan new.xls");
        try {
            POIFSFileSystem poifs = new POIFSFileSystem(pertindakanNew);
            BookPertindakanNew = new HSSFWorkbook(poifs);

            CellStyle centerTextStyle = BookPertindakanNew.createCellStyle();
            centerTextStyle.setAlignment(HorizontalAlignment.CENTER);
            CellStyle AllBorderStyle = BookPertindakanNew.createCellStyle();
            AllBorderStyle.setBorderBottom(BorderStyle.THIN);
            AllBorderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            AllBorderStyle.setBorderLeft(BorderStyle.THIN);
            AllBorderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            AllBorderStyle.setBorderRight(BorderStyle.THIN);
            AllBorderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            AllBorderStyle.setBorderTop(BorderStyle.THIN);
            AllBorderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());

            Sheet pertindakan_New_Raw = BookPertindakanNew.getSheetAt(0);

            Sheet Pertindakan = BookPertindakanNew.createSheet();
            BookPertindakanNew.setSheetName(1, "1 Pertindakan");
            for (int cell = 0; cell < pertindakan_New_Raw.getRow(0).getLastCellNum(); cell++) {
                for (int row = 0; row < pertindakan_New_Raw.getLastRowNum(); row++) {
                    Pertindakan.createRow(row).createCell(cell);
                }
            }

            // Perform pivot simulation
            Map<String, Integer> pivotJumlahTindakan = new HashMap<>();
            for (int i = 1; i <= pertindakan_New_Raw.getLastRowNum(); i++) {
                Row row = pertindakan_New_Raw.getRow(i);
                String Tindakan = row.getCell(15).getStringCellValue();
                Integer count = pivotJumlahTindakan.getOrDefault(Tindakan, 0);
                count++;
                pivotJumlahTindakan.put(Tindakan, count);
            }

//          Sort any value it contains
            List<Map.Entry<String, Integer>> entriesDoctor = new ArrayList<>(pivotJumlahTindakan.entrySet());
            entriesDoctor.sort(Map.Entry.comparingByKey());
            pivotJumlahTindakan = new LinkedHashMap<>();
            for (Map.Entry<String, Integer> entry : entriesDoctor) {
                pivotJumlahTindakan.put(entry.getKey(), entry.getValue());
            }

            int rowNum = 6;
            for (Map.Entry<String, Integer> entry : pivotJumlahTindakan.entrySet()) {
                Row row = Pertindakan.createRow(rowNum++);
                row.createCell(0).setCellValue(rowNum-6);
                row.createCell(1).setCellValue(entry.getKey());
                row.createCell(2).setCellValue(entry.getValue());
                row.getCell(0).setCellStyle(AllBorderStyle);
                row.getCell(1).setCellStyle(AllBorderStyle);
                row.getCell(2).setCellStyle(AllBorderStyle);
            }

            int columnCountA2 = Pertindakan.getRow(0).getLastCellNum();
            for (int columnIndex = 0; columnIndex < columnCountA2; columnIndex++) {
                Pertindakan.autoSizeColumn(columnIndex);
            }














        } catch (Exception e) {
            e.printStackTrace();
        }


        try {
            outputStream = new FileOutputStream("pertindakanNew.xlsx");
            BookPertindakanNew.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (BookPertindakanNew != null) {
                    BookPertindakanNew.close();
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