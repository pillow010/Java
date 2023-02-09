package main.java.LaporanRad;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.*;
import java.util.*;

public class A_Pertindakan {

    public static void main(String[] args) {
    new A_Pertindakan();
    }
    private Workbook BookPertindakanNew;

    private FileOutputStream OutputPertindakanNew;

    public A_Pertindakan(){
        Sheet SheetPertindakanNew = null;
//        Sheet sheetB = null;
        File PertindakanNew = new File("C:\\sat work\\test\\rad pertindakan new.xls");
//        File jasaUnit = new File("C:\\sat work\\test\\a) LAPORAN REKAP PENERIMAAN JASA UNIT PER PASIEN1.xls");
        try {
            POIFSFileSystem poifs = new POIFSFileSystem(PertindakanNew);

            Workbook BookPertindakanNew = new HSSFWorkbook(poifs);


            CellStyle centerTextStyle = BookPertindakanNew.createCellStyle();
            centerTextStyle.setAlignment(HorizontalAlignment.CENTER);

            SheetPertindakanNew = BookPertindakanNew.getSheetAt(0);
            int SheetPertindakanNewLastRow = SheetPertindakanNew.getLastRowNum();
            int SheetPertindakanNewLastCell = SheetPertindakanNew.getRow(0).getLastCellNum();
            System.out.println(SheetPertindakanNewLastRow);
            System.out.println(SheetPertindakanNewLastCell);

            Sheet Pertindakan = BookPertindakanNew.createSheet();
            BookPertindakanNew.setSheetName(1, "1 Pertindakan");
            for (int cell = 0; cell < SheetPertindakanNew.getRow(0).getLastCellNum(); cell++) {
                for (int row = 0; row < SheetPertindakanNew.getLastRowNum(); row++) {
                    Pertindakan.createRow(row).createCell(cell);
                }
            }

            // Perform pivot simulation
            Map<String, Integer> pivotJumlahTindakan = new HashMap<>();
            for (int i = 1; i <= SheetPertindakanNew.getLastRowNum(); i++) {
                Row row = SheetPertindakanNew.getRow(i);
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
            }



















        }catch (Exception e) {
            e.printStackTrace();
        }
        try {
            OutputPertindakanNew = new FileOutputStream("Lab Half Done.xlsx");
            BookPertindakanNew.write(OutputPertindakanNew);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (BookPertindakanNew != null) {
                    BookPertindakanNew.close();
                }
                if (OutputPertindakanNew != null) {
                    OutputPertindakanNew.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}