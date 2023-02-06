package main.java.LaporanJasa;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;

public class A_RekapJasaDokterDanUnit {

    public static void main(String[] args) {
        new A_RekapJasaDokterDanUnit();
    }
    private Workbook workbook;

    private FileOutputStream outputStream;


    public A_RekapJasaDokterDanUnit() {
        Sheet sheetA = null;
        Sheet sheetB = null;
        File jasaDokter = new File("C:\\sat work\\test\\c) LAPORAN PENERIMAAN JASA PELAYANAN PER TINDAKAN1.xls");
        File jasaUnit = new File("C:\\sat work\\test\\a) LAPORAN REKAP PENERIMAAN JASA UNIT PER PASIEN1.xls");
        try {
            POIFSFileSystem poifs = new POIFSFileSystem(jasaDokter);
            POIFSFileSystem poifs2 = new POIFSFileSystem(jasaDokter);
            workbook = new HSSFWorkbook(poifs);
            Workbook workbook2 = new HSSFWorkbook(poifs2);


            sheetA = workbook.getSheetAt(0);
            int lastRow = sheetA.getLastRowNum();
            sheetB = workbook2.getSheetAt(0);
            int lastRowB = sheetB.getLastRowNum();

            Sheet sheetA2 = workbook.createSheet();
            workbook.setSheetName(1, "JASA DOKTER");
            for (int cell = 0; cell <= 6; cell++) {
                for (int row = 0; row <= lastRow-1; row++) {
                    sheetA2.createRow(row).createCell(cell);
                }
            }

            Sheet sheetA3 = workbook.createSheet();
            workbook.setSheetName(2, "JASA UNIT");
            for (int cell = 0; cell <= 6; cell++) {
                for (int row = 0; row <= lastRowB-1; row++) {
                    sheetA3.createRow(row).createCell(cell);
                }
            }


            // Perform pivot simulation
            Map<String, Double> pivotData = new HashMap<>();
//            for (int i = 1; i <= sheetA.getLastRowNum(); i++) {
//                Row row = sheetA.getRow(i);
//                String p = row.getCell(11).getStringCellValue();
//                Integer count = pivotData.getOrDefault(p, 0);
////                System.out.println(row.getCell(46).getNumericCellValue());
//                count += (int) row.getCell(46).getNumericCellValue();
//                pivotData.put(p, count);
//            }

            for (int i = 1; i <= sheetA.getLastRowNum(); i++) {
                Row row = sheetA.getRow(i);
                String doctor = row.getCell(11).getStringCellValue();
                Double count = pivotData.getOrDefault(doctor, 0.0);
                count += row.getCell(46).getNumericCellValue();
                pivotData.put(doctor, count);
            }

            List<Map.Entry<String, Double>> entries = new ArrayList<>(pivotData.entrySet());
            entries.sort(Map.Entry.comparingByKey());
            pivotData = new LinkedHashMap<>();
            for (Map.Entry<String, Double> entry : entries) {
                pivotData.put(entry.getKey(), entry.getValue());
            }


            // Write pivot data to sheetA
//            int rowNum = sheetA.getLastRowNum() + 1;
            int rowNum = 6;
            for (Map.Entry<String, Double> entry : pivotData.entrySet()) {
                Row row = sheetA2.createRow(rowNum++);
                row.createCell(0).setCellValue(rowNum-6);
                row.createCell(1).setCellValue(entry.getKey());
                row.createCell(2).setCellValue(entry.getValue());
            }



            workbook.removeSheetAt(0);

        } catch (Exception e) {
            e.printStackTrace();
        }


        try {
            outputStream = new FileOutputStream("1. REKAP JASA DOKTER DAN UNIT.xlsx");
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
