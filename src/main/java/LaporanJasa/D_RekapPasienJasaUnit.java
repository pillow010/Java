package LaporanJasa;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;

import static LaporanJasa.LaporanJasaCommandCenter.fileOutput;
import static LaporanJasa.LaporanJasaCommandCenter.fileSource;

public class D_RekapPasienJasaUnit {

    public static void main(String[] args) {
//        new D_RekapPasienJasaUnit ();
    }

    private Workbook workbook;
    private FileOutputStream outputStream;


    public D_RekapPasienJasaUnit(){
        File xlsxFile = new File(fileSource + "a) LAPORAN REKAP PENERIMAAN JASA UNIT PER PASIEN.xlsx");
        File xlsFile = new File(fileSource + "a) LAPORAN REKAP PENERIMAAN JASA UNIT PER PASIEN.xls");
        File jasaUnit;
        if (xlsxFile.exists()) {
            jasaUnit = xlsxFile;
        } else if (xlsFile.exists()) {
            jasaUnit = xlsFile;
        } else {
            System.out.println("File not found: " + fileSource + "a) LAPORAN REKAP PENERIMAAN JASA UNIT PER PASIEN");
            return;
        }

        System.out.println("D_RekapPasienJasaUnit is starting");
        try {
            LocalDateTime start = LocalDateTime.now();
            FileInputStream inputStream = new FileInputStream(jasaUnit);
            workbook = WorkbookFactory.create(inputStream);

            Sheet sheet = workbook.getSheetAt(0);
            Sheet sheet2 = workbook.createSheet();

            DataFormat dataFormat = workbook.createDataFormat();

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
            totalStyle.setDataFormat (dataFormat.getFormat("_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-"));
            totalStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            totalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            workbook.setSheetName(1, "4. REKAP PASIEN JASA UNIT");
            int lastColumn = sheet.getRow(0).getLastCellNum();
            int lastRow = sheet.getLastRowNum();
            for (int row =1;row<=lastRow;row++){
                sheet2.createRow(row);
            }

            Row row = sheet2.getRow(0);
            if (row == null) {
                row = sheet2.createRow(0);
            }
            row.createCell(0).setCellValue("NORM");
            row.createCell(1).setCellValue("NAMA");
            row.createCell(2).setCellValue("NOREG");
            row.createCell(3).setCellValue("TGL REG");
            row.createCell(4).setCellValue("NAMA UNIT");
            row.createCell(5).setCellValue("KET INST");
            row.createCell(6).setCellValue("KET SUB INST");
            row.createCell(7).setCellValue("KET DTL SUB INST");
            row.createCell(8).setCellValue("JML PENERIMAAN");
            row.createCell(9).setCellValue("JML KOREKSI");
            row.createCell(10).setCellValue("JML PAJAK");
            row.createCell(11).setCellValue("JML PENGAMBILAN");
            row.createCell(12).setCellValue("JML NETTO");

            for (int column = 0; column <= lastColumn - 1; column++) {
//          jika cell mengandung "KD_INST" concat jadi noreg
                Cell cell = sheet.getRow(0).getCell(column);

                String cellValue = cell.getStringCellValue();
                int targetColumn2 = switch (cellValue) {
                    case "NORM" -> 0;
                    case "NAMA" -> 1;
                    case "NO_REG" -> 2;
                    case "TGL_REG" -> 3;
                    case "NAMA_UNIT"-> 4;
                    case "KET_INST"-> 5;
                    case "KET_SUB_INST"-> 6;
                    case "KET_DTL_SUBINST"-> 7;
                    case "JML_PENERIMAAN"-> 8;
                    case "JML_KOREKSI"-> 9;
                    case "JML_PAJAK"-> 10;
                    case "JML_PENGAMBILAN"-> 11;
                    case "JML_NETTO"-> 12;
                    default -> -1;
                };

                if (targetColumn2 != -1) {
                    for (int i = 1; i <= lastRow; i++) {
                        Cell targetCell = sheet2.getRow(i).createCell(targetColumn2);
                        if (sheet.getRow (i).getCell (column)==null){
                            targetCell.setCellValue ("");
                        }else if (sheet.getRow(i).getCell(column).getCellType() == CellType.STRING) {
                            targetCell.setCellValue (sheet.getRow (i).getCell (column).getStringCellValue ());
                        }else if (cellValue.equals("JML_NETTO")) {
//                            targetCell.setCellValue(formatter.format(sheet.getRow(i).getCell(column).getNumericCellValue()));
                            targetCell.setCellValue(sheet.getRow(i).getCell(column).getNumericCellValue());
                        } else {
                            targetCell.setCellValue(sheet.getRow(i).getCell(column).getNumericCellValue());
                        }
                    }
                }

            }
            int columnCountA2 = sheet2.getRow(0).getLastCellNum();
            for (int columnIndex = 0; columnIndex < columnCountA2; columnIndex++) {
                sheet2.autoSizeColumn(columnIndex);
            }
            CellStyle centerTextStyle = workbook.createCellStyle();
            centerTextStyle.setAlignment(HorizontalAlignment.CENTER);
            int lastCellA2 = sheet2.getRow(0).getLastCellNum();
            for (int title=0;title<lastCellA2;title++){
                sheet2.getRow(0).getCell(title).setCellStyle(centerTextStyle);
            }

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

            // Get the last column index
            int lastColumnIndex = sheet2.getRow(5).getLastCellNum();

            // Apply header style and border to the header row and auto-size columns
            for (int rightCell = 0; rightCell < lastColumnIndex; rightCell++) {
                sheet2.getRow(0).getCell(rightCell).setCellStyle(BorderCenterCellStyle);
                sheet2.autoSizeColumn(rightCell);
            }

            // Apply cell styles and borders to the data rows
            for (int rightCell = 0; rightCell < lastColumnIndex; rightCell++) {
                for (int downRow = 1; downRow <= sheet2.getLastRowNum(); downRow++) {
                    sheet2.getRow(downRow).getCell(rightCell).setCellStyle(AllBorderCellStyle);
                }
            }

            // Apply the target column style to column 22 for all data rows
            for (int downRow = 1; downRow <= sheet2.getLastRowNum(); downRow++) {
                sheet2.getRow(downRow).getCell(12).setCellStyle(totalStyle);
            }

            workbook.removeSheetAt(0);
            LocalDateTime end = LocalDateTime.now();
            Duration duration = Duration.between(start, end);
            double seconds = duration.toMillis() / 1000.0;
            System.out.printf("D_RekapPasienJasaUnit Done in %.4f seconds%n", seconds);

        } catch (Exception e) {
            e.printStackTrace();
        }

        try {
            outputStream = new FileOutputStream(fileOutput+"4. REKAP PASIEN JASA UNIT.xlsx");
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
