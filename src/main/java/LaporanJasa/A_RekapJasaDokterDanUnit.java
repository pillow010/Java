package LaporanJasa;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.time.Duration;
import java.util.*;
import java.time.LocalDateTime;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

public class A_RekapJasaDokterDanUnit {

    public static void main(String[] args) {
//        new A_RekapJasaDokterDanUnit();
    }
    private Workbook workbook;

    private FileOutputStream outputStream;


    public A_RekapJasaDokterDanUnit() {


        //XLSX VER
        File jasaUnit = new File("C:\\sat work\\test\\a) LAPORAN REKAP PENERIMAAN JASA UNIT PER PASIEN.xlsx");  //XLSX
//        File jasaUnit = new File("C:\\sat work\\test\\a) LAPORAN REKAP PENERIMAAN JASA UNIT PER PASIEN.xls");
        File jasaDokter = new File("C:\\sat work\\test\\c) LAPORAN PENERIMAAN JASA PELAYANAN PER TINDAKAN.xlsx");      //XLSX
//        File jasaDokter = new File("C:\\sat work\\test\\c) LAPORAN PENERIMAAN JASA PELAYANAN PER TINDAKAN.xls");
        LocalDateTime start = LocalDateTime.now ();
        System.out.println ("A_RekapJasaDokterDanUnit is starting");
        try {
            FileInputStream inputStream2 = new FileInputStream(jasaUnit);     //XLSX
//            POIFSFileSystem poifs2 = new POIFSFileSystem(jasaUnit);
            FileInputStream inputStream1 = new FileInputStream(jasaDokter); //XLSX
//            POIFSFileSystem poifs = new POIFSFileSystem(jasaDokter);
            Workbook workbook2 = new XSSFWorkbook (inputStream2);                       //XLSX
//            Workbook workbook2 = new HSSFWorkbook(poifs2);
            workbook = new XSSFWorkbook(inputStream1);                      //XLSX
//            workbook = new HSSFWorkbook(poifs);


            // Make Styling
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

            CellStyle totalStyle = workbook.createCellStyle ();
            totalStyle.setAlignment(HorizontalAlignment.RIGHT);
            totalStyle.setBorderBottom (BorderStyle.THIN);
            totalStyle.setBottomBorderColor (IndexedColors.BLACK.getIndex ());
            totalStyle.setBorderLeft (BorderStyle.THIN);
            totalStyle.setLeftBorderColor (IndexedColors.BLACK.getIndex ());
            totalStyle.setBorderRight (BorderStyle.THIN);
            totalStyle.setRightBorderColor (IndexedColors.BLACK.getIndex ());
            totalStyle.setBorderTop (BorderStyle.THIN);
            totalStyle.setTopBorderColor (IndexedColors.BLACK.getIndex ());

            CellStyle centerTextStyle = workbook.createCellStyle();
            centerTextStyle.setAlignment(HorizontalAlignment.CENTER);

            //create workbook
            Sheet sheetWorkbookDokter = workbook.getSheetAt(0); //jasdok
            Sheet sheetWorkbookUnit = workbook2.getSheetAt(0);//jasnit

            //Cara Bayar
            String caraBayarDokter = sheetWorkbookDokter.getRow(1).getCell(17).getStringCellValue();
            String caraBayarUnit = sheetWorkbookUnit.getRow(1).getCell(24).getStringCellValue();

            if (caraBayarDokter.contains("PBI")) {
                caraBayarDokter = "JKN";
            }
            if (caraBayarUnit.contains("PBI")) {
                caraBayarUnit = "JKN";
            }

            //create Sheet
            Sheet sheetJasaDokter = workbook.createSheet();
            workbook.setSheetName(1, "JASA DOKTER "+caraBayarDokter);
            Sheet sheetJasaUnit = workbook.createSheet();
            workbook.setSheetName(2, "JASA UNIT "+caraBayarUnit);


            // Read data from sheetWorkbookDokter and group by doctor name, summing the netto values
            Map<String, Double> pivotDataDoctor = StreamSupport.stream(sheetWorkbookDokter.spliterator(), false)
                    .skip(1) // skip header row
                    .collect(Collectors.groupingBy(
                            row -> row.getCell(11).getStringCellValue(),
                            Collectors.summingDouble(row -> row.getCell(46).getNumericCellValue())
                    ));

            Map<String, Double> pivotDataUnit = StreamSupport.stream(sheetWorkbookUnit.spliterator(), false)
                    .skip(1)
                    .collect(Collectors.groupingBy(
                            row -> row.getCell(5).getStringCellValue(),
                            Collectors.summingDouble(row -> {
                                Cell cell = row.getCell(19);
                                if (cell.getCellType() == CellType.NUMERIC) {
                                    return cell.getNumericCellValue();
                                } else if (cell.getCellType() == CellType.STRING) {
                                    try {
                                        return Double.parseDouble(cell.getStringCellValue());
                                    } catch (NumberFormatException e) {
                                        // Handle the case when the cell value is not a valid number
                                        // You can add appropriate error handling or logging here
                                        return 0.0; // Set a default value for invalid cell values
                                    }
                                } else {
                                    // Handle other cell types if needed
                                    return 0.0; // Set a default value for unsupported cell types
                                }
                            })
                    ));


            // Sort by doctor name and write to sheetJasaDokter
            List<Map.Entry<String, Double>> entriesDoctor = new ArrayList<>(pivotDataDoctor.entrySet());
            List<Map.Entry<String, Double>> entriesUnit = new ArrayList<> (pivotDataUnit.entrySet ());
            entriesDoctor.sort(Map.Entry.comparingByKey());
            entriesUnit.sort (Map.Entry.comparingByKey ());
            int rowNum = 6;
            for (Map.Entry<String, Double> entry : entriesDoctor) {
                Row row = sheetJasaDokter.createRow(rowNum++);
                row.createCell(0).setCellValue(rowNum - 6);
                row.createCell(1).setCellValue(entry.getKey());
                DecimalFormat formatter = new DecimalFormat ("#,##0;-#,##0");
                String formattedValue = formatter.format(entry.getValue());
                row.createCell(2).setCellValue(formattedValue + ",-");
            }

            rowNum = 6;
            for (Map.Entry<String, Double> entry : entriesUnit){
                Row rowA3 = sheetJasaUnit.createRow(rowNum++);
                rowA3.createCell(0).setCellValue(rowNum - 6);
                rowA3.createCell(1).setCellValue(entry.getKey());
                DecimalFormat formatter = new DecimalFormat("#,##0;-#,##0");
                String formattedValue = formatter.format(entry.getValue());
                rowA3.createCell(2).setCellValue(formattedValue + ",-");
            }

            sheetJasaDokter.createRow (0).createCell (0).setCellValue (caraBayarDokter);
            sheetJasaDokter.createRow (5).createCell(0).setCellValue("NO");
            sheetJasaDokter.getRow (5).createCell(1).setCellValue("NAMA DOKTER");
            sheetJasaDokter.getRow(5).createCell(2).setCellValue("TOTAL");

            sheetJasaUnit.createRow (0).createCell (0).setCellValue (caraBayarUnit);
            sheetJasaUnit.createRow (5).createCell(0).setCellValue("NO");
            sheetJasaUnit.getRow(5).createCell(1).setCellValue("NAMA UNIT");
            sheetJasaUnit.getRow(5).createCell(2).setCellValue("TOTAL");


            //buat header center kemudian border semuanya ps. use'<' because return 2 but there is 0, and 1. no number 2.
            for (int rightCell = 0; rightCell < sheetJasaDokter.getRow (5).getLastCellNum (); rightCell++) {
                sheetJasaDokter.getRow (5).getCell (rightCell).setCellStyle (BorderCenterCellStyle);
                sheetJasaDokter.autoSizeColumn (rightCell);
                for (int downRow = 6; downRow <= sheetJasaDokter.getLastRowNum (); downRow++) {
                    if (!(rightCell == 2)) {
                        sheetJasaDokter.getRow (downRow).getCell (rightCell).setCellStyle (AllBorderCellStyle);
                    } else {
                        sheetJasaDokter.getRow (downRow).getCell (2).setCellStyle (totalStyle);
                    }
                }
            }

            //buat header center kemudian border semuanya ps. use'<' because return 2 but there is 0, and 1. no number 2.
            for (int rightCell = 0; rightCell < sheetJasaUnit.getRow (5).getLastCellNum (); rightCell++) {
                sheetJasaUnit.getRow (5).getCell (rightCell).setCellStyle (BorderCenterCellStyle);
                sheetJasaUnit.autoSizeColumn (rightCell);
                for (int downRow = 6; downRow <= sheetJasaUnit.getLastRowNum (); downRow++) {
                    if (!(rightCell == 2)) {
                        sheetJasaUnit.getRow (downRow).getCell (rightCell).setCellStyle (AllBorderCellStyle);
                    } else {
                        sheetJasaUnit.getRow (downRow).getCell (2).setCellStyle (totalStyle);
                    }
                }
            }

            workbook.removeSheetAt(0);

            LocalDateTime end = LocalDateTime.now ();
            Duration duration = Duration.between(start, end);
            long seconds = duration.toMillis ();
            System.out.println ("A_RekapJasaDokterDanUnit Done in " + seconds);
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
