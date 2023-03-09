package LaporanJasa;

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
        File jasaDokter = new File("C:\\sat work\\test\\c) LAPORAN PENERIMAAN JASA PELAYANAN PER TINDAKAN1.xls");
        File jasaUnit = new File("C:\\sat work\\test\\a) LAPORAN REKAP PENERIMAAN JASA UNIT PER PASIEN1.xls");
        System.out.println ("A_RekapJasaDokterDanUnit is starting");
        try {
            POIFSFileSystem poifs = new POIFSFileSystem(jasaDokter);
            POIFSFileSystem poifs2 = new POIFSFileSystem(jasaUnit);
            workbook = new HSSFWorkbook(poifs);
            Workbook workbook2 = new HSSFWorkbook(poifs2);

            // Make Styling
            CellStyle centerTextCellStyle = workbook.createCellStyle ();
            centerTextCellStyle.setAlignment (HorizontalAlignment.CENTER);
            CellStyle AllBorderCellStyle = workbook.createCellStyle ();
            AllBorderCellStyle.setBorderBottom (BorderStyle.THIN);
            AllBorderCellStyle.setBottomBorderColor (IndexedColors.BLACK.getIndex ());
            AllBorderCellStyle.setBorderLeft (BorderStyle.THIN);
            AllBorderCellStyle.setLeftBorderColor (IndexedColors.BLACK.getIndex ());
            AllBorderCellStyle.setBorderRight (BorderStyle.THIN);
            AllBorderCellStyle.setRightBorderColor (IndexedColors.BLACK.getIndex ());
            AllBorderCellStyle.setBorderTop (BorderStyle.THIN);
            AllBorderCellStyle.setTopBorderColor (IndexedColors.BLACK.getIndex ());
            CellStyle BorderCenterCellStyle = workbook.createCellStyle ();
            BorderCenterCellStyle.setAlignment (HorizontalAlignment.CENTER);
            BorderCenterCellStyle.setBorderBottom (BorderStyle.THIN);
            BorderCenterCellStyle.setBottomBorderColor (IndexedColors.BLACK.getIndex ());
            BorderCenterCellStyle.setBorderLeft (BorderStyle.THIN);
            BorderCenterCellStyle.setLeftBorderColor (IndexedColors.BLACK.getIndex ());
            BorderCenterCellStyle.setBorderRight (BorderStyle.THIN);
            BorderCenterCellStyle.setRightBorderColor (IndexedColors.BLACK.getIndex ());
            BorderCenterCellStyle.setBorderTop (BorderStyle.THIN);
            BorderCenterCellStyle.setTopBorderColor (IndexedColors.BLACK.getIndex ());

            CellStyle centerTextStyle = workbook.createCellStyle();
            centerTextStyle.setAlignment(HorizontalAlignment.CENTER);

            Sheet sheetA = workbook.getSheetAt(0);
            Sheet sheetB = workbook2.getSheetAt(0);

            Sheet sheetA2 = workbook.createSheet();
            workbook.setSheetName(1, "JASA DOKTER");


            Sheet sheetA3 = workbook.createSheet();
            workbook.setSheetName(2, "JASA UNIT");


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


//          Sort any value it contains
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




            sheetA2.createRow (5).createCell(0).setCellValue("NO");
            sheetA2.getRow (5).createCell(1).setCellValue("NAMA DOKTER");
            sheetA2.getRow(5).createCell(2).setCellValue("TOTAL");


            sheetA3.createRow (5).createCell(0).setCellValue("NO");
            sheetA3.getRow(5).createCell(1).setCellValue("NAMA UNIT");
            sheetA3.getRow(5).createCell(2).setCellValue("TOTAL");


            //buat header center kemudian border semuanya ps. use'<' because return 2 but there is 0, and 1. no number 2.
            for (int rightCell = 0; rightCell < sheetA2.getRow (5).getLastCellNum (); rightCell++) {
                sheetA2.getRow (5).getCell (rightCell).setCellStyle (BorderCenterCellStyle);
                sheetA2.autoSizeColumn (rightCell);
                for (int downRow = 5; downRow <= sheetA2.getLastRowNum (); downRow++) {
                    sheetA2.getRow (downRow).getCell (rightCell).setCellStyle (AllBorderCellStyle);
                }
            }

            //buat header center kemudian border semuanya ps. use'<' because return 2 but there is 0, and 1. no number 2.
            for (int rightCell = 0; rightCell < sheetA3.getRow (5).getLastCellNum (); rightCell++) {
                sheetA3.getRow (5).getCell (rightCell).setCellStyle (BorderCenterCellStyle);
                sheetA3.autoSizeColumn (rightCell);
                for (int downRow = 5; downRow <= sheetA3.getLastRowNum (); downRow++) {
                    sheetA3.getRow (downRow).getCell (rightCell).setCellStyle (AllBorderCellStyle);
                }
            }


            workbook.removeSheetAt(0);

            System.out.println ("A_RekapJasaDokterDanUnit Done");
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
