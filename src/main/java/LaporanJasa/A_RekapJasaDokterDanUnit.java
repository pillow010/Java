package main.java.LaporanJasa;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.*;
import java.util.*;

public class A_RekapJasaDokterDanUnit {

    public static void main(String[] args) {
//        new A_RekapJasaDokterDanUnit();
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
            POIFSFileSystem poifs2 = new POIFSFileSystem(jasaUnit);
            workbook = new HSSFWorkbook(poifs);
            Workbook workbook2 = new HSSFWorkbook(poifs2);

            CellStyle centerTextStyle = workbook.createCellStyle();
            centerTextStyle.setAlignment(HorizontalAlignment.CENTER);


            sheetA = workbook.getSheetAt(0);
            int lastRow = sheetA.getLastRowNum();
            sheetB = workbook2.getSheetAt(0);
            int lastRowB = sheetB.getLastRowNum();

            Sheet sheetA2 = workbook.createSheet();
            workbook.setSheetName(1, "JASA DOKTER");
            for (int cell = 0; cell <= 2; cell++) {
                for (int row = 0; row <= lastRow-1; row++) {
                    sheetA2.createRow(row).createCell(cell);
                }
            }

            Sheet sheetA3 = workbook.createSheet();
            workbook.setSheetName(2, "JASA UNIT");
            for (int cell = 0; cell <= 2; cell++) {
                for (int row = 0; row <= lastRowB-1; row++) {
                    sheetA3.createRow(row).createCell(cell);
                }
            }

            // Perform pivot simulation
            Map<String, Double> pivotDataDoctor = new HashMap<>();
            for (int i = 1; i <= sheetA.getLastRowNum(); i++) {
                Row row = sheetA.getRow(i);
                String doctor = row.getCell(11).getStringCellValue();
                Double count = pivotDataDoctor.getOrDefault(doctor, 0.0);
                count += row.getCell(46).getNumericCellValue();
                pivotDataDoctor.put(doctor, count);
            }

            Map<String, Double> pivotDataUnit = new HashMap<>();
            for (int i = 1; i <= sheetB.getLastRowNum(); i++) {
                Row row = sheetB.getRow(i);
                String doctor = row.getCell(5).getStringCellValue();
                Double count = pivotDataUnit.getOrDefault(doctor, 0.0);
                count += row.getCell(19).getNumericCellValue();
                pivotDataUnit.put(doctor, count);
            }


//          Sort any value it contain
            List<Map.Entry<String, Double>> entriesDoctor = new ArrayList<>(pivotDataDoctor.entrySet());
            entriesDoctor.sort(Map.Entry.comparingByKey());
            pivotDataDoctor = new LinkedHashMap<>();
            for (Map.Entry<String, Double> entry : entriesDoctor) {
                pivotDataDoctor.put(entry.getKey(), entry.getValue());
            }

            List<Map.Entry<String, Double>> entriesUnit = new ArrayList<>(pivotDataUnit.entrySet());
            entriesUnit.sort(Map.Entry.comparingByKey());
            pivotDataUnit = new LinkedHashMap<>();
            for (Map.Entry<String, Double> entry : entriesUnit) {
                pivotDataUnit.put(entry.getKey(), entry.getValue());
            }

            // Write pivot data to sheetA
//            int rowNum = sheetA.getLastRowNum() + 1;
            int rowNum = 6;
            for (Map.Entry<String, Double> entry : pivotDataDoctor.entrySet()) {
                Row row = sheetA2.createRow(rowNum++);
                row.createCell(0).setCellValue(rowNum-6);
                row.createCell(1).setCellValue(entry.getKey());
                row.createCell(2).setCellValue(entry.getValue());
            }

            int rowNumA3 = 6;
            for (Map.Entry<String, Double> entry : pivotDataUnit.entrySet()) {
                Row row = sheetA3.createRow(rowNumA3++);
                row.createCell(0).setCellValue(rowNumA3-6);
                row.createCell(1).setCellValue(entry.getKey());
                row.createCell(2).setCellValue(entry.getValue());
            }

            int columnCountA2 = sheetA2.getRow(0).getLastCellNum();
            for (int columnIndex = 0; columnIndex < columnCountA2; columnIndex++) {
                sheetA2.autoSizeColumn(columnIndex);
            }
            int columnCountA3 = sheetA3.getRow(0).getLastCellNum();
            for (int columnIndex = 0; columnIndex < columnCountA3; columnIndex++) {
                sheetA3.autoSizeColumn(columnIndex);
            }



            sheetA2.getRow(5).createCell(0).setCellValue("NO");
            sheetA2.getRow(5).createCell(1).setCellValue("NAMA DOKTER");
            sheetA2.getRow(5).createCell(2).setCellValue("TOTAL");
            int lastCellA2 = sheetA2.getRow(5).getLastCellNum();
            for (int title=0;title<lastCellA2;title++){
                    sheetA2.getRow(5).getCell(title).setCellStyle(centerTextStyle);
            }

            sheetA3.getRow(5).createCell(0).setCellValue("NO");
            sheetA3.getRow(5).createCell(1).setCellValue("NAMA UNIT");
            sheetA3.getRow(5).createCell(2).setCellValue("TOTAL");
            int lastCellA3 = sheetA3.getRow(5).getLastCellNum();
            for (int title=0;title<lastCellA3;title++){
                sheetA3.getRow(5).getCell(title).setCellStyle(centerTextStyle);
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
