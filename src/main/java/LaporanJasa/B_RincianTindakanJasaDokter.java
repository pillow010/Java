package LaporanJasa;

import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.time.Duration;
import java.time.LocalDateTime;

import static LaporanJasa.LaporanJasaCommandCenter.*;

public class B_RincianTindakanJasaDokter {
    public static void main(String[] args) {
//        new B_RincianTindakanJasaDokter();
    }

    private Workbook workbook;
    private FileOutputStream outputStream;


    public B_RincianTindakanJasaDokter() {
        File xlsxFile   = new File(fileSource + "c) LAPORAN PENERIMAAN JASA PELAYANAN PER TINDAKAN.xlsx");
        File xlsFile    = new File(fileSource + "c) LAPORAN PENERIMAAN JASA PELAYANAN PER TINDAKAN.xls");
        File jasaDokter;
        if (xlsxFile.exists()) {
            jasaDokter = xlsxFile;
        } else if (xlsFile.exists()) {
            jasaDokter = xlsFile;
        } else {
            System.out.println("File not found: " + fileSource + "c) LAPORAN PENERIMAAN JASA PELAYANAN PER TINDAKAN");
            return;
        }

        System.out.println("B_RincianTindakanJasaDokter is starting");
        try {
            LocalDateTime start = LocalDateTime.now();
            FileInputStream inputStream = new FileInputStream(jasaDokter);
            workbook = WorkbookFactory.create(inputStream);


            Sheet sheet = workbook.getSheetAt(0);
            Sheet sheet2 = workbook.createSheet();

            workbook.setSheetName(1, "2. RINCIAN TINDAKAN JASA DOKTER");
            int lastColumn = sheet.getRow(0).getLastCellNum();
            int lastRow = sheet.getLastRowNum();
            for (int row =1;row<=lastRow;row++){
                sheet2.createRow(row);
            }

            Row row = sheet2.getRow(0);
            if (row == null) {
                row = sheet2.createRow(0);
            }
            row.createCell(0).setCellValue("NAMA PASIEN");
            row.createCell(1).setCellValue("NORM");
            row.createCell(2).setCellValue("NO REG");
            row.createCell(3).setCellValue("KET INST");
            row.createCell(4).setCellValue("KET SUB INST");
            row.createCell(5).setCellValue("KET DTL SUB INST");
            row.createCell(6).setCellValue("NAMA DOKTER");
            row.createCell(7).setCellValue("POSISI");
            row.createCell(8).setCellValue("TGL TINDAKAN");
            row.createCell(9).setCellValue("NM TINDAKAN");
            row.createCell(10).setCellValue("RUANG RAWAT");
            row.createCell(11).setCellValue("PAKET JAMINAN");
            row.createCell(12).setCellValue("JASA PELAYANAN TARIF");
            row.createCell(13).setCellValue("JASA PELAYANAN JAMIN");
            row.createCell(14).setCellValue("JML PENDAPATAN");
            row.createCell(15).setCellValue("JML PENERIMAAN TUNAI");
            row.createCell(16).setCellValue("JML PENERIMAAN PIUTANG");
            row.createCell(17).setCellValue("JML PENERIMAAN JMN");
            row.createCell(18).setCellValue("JML KOREKSI");
            row.createCell(19).setCellValue("JML PAJAK");
            row.createCell(20).setCellValue("JML PENGURANG JASA");
            row.createCell(21).setCellValue("JML PENGAMBILAN");
            row.createCell(22).setCellValue("JML NETTO");

            for (int column = 0; column <= lastColumn - 1; column++) {
                // jika cell mengandung "KD_INST" concat jadi noreg
                Cell cell = sheet.getRow(0).getCell(column);
                if (sheet.getRow(0).getCell(column).getStringCellValue().equals("KD_INST")) {
                    for (int i = 1; i <= lastRow; i++) {
                        Cell noReg = sheet2.getRow(i).createCell(2);
                        noReg.setCellValue(sheet.getRow(i).getCell(column).getStringCellValue() +
                                sheet.getRow(i).getCell(column + 1).getStringCellValue() +
                                sheet.getRow(i).getCell(column + 2).getStringCellValue() +
                                sheet.getRow(i).getCell(column + 3).getStringCellValue() +
                                sheet.getRow(i).getCell(column + 4).getStringCellValue());
                    }
                }

                String cellValue = cell.getStringCellValue();
                int targetColumn2 = switch (cellValue) {
                    case "NAMA_PASIEN" -> 0;
                    case "NORM" -> 1;
                    case "NO_REG" -> 2;
                    case "KET_INST" -> 3;
                    case "KET_SUB_INST" -> 4;
                    case "KET_DTL_SUB_INST" -> 5;
                    case "NAMA_DOKTER" -> 6;
                    case "POSISI" -> 7;
                    case "TGL_TINDAKAN" -> 8;
                    case "NM_TINDAKAN" -> 9;
                    case "RUANG_RAWAT" -> 10;
                    case "PAKET_JAMINAN" -> 11;
                    case "JASA_PELAYANAN_TARIF" -> 12;
                    case "JASA_PELAYANAN_JAMIN" -> 13;
                    case "JML_PENDAPATAN" -> 14;
                    case "JML_PENERIMAAN_TUNAI" -> 15;
                    case "JML_PENERIMAAN_PIUTANG" -> 16;
                    case "JML_PENERIMAAN_JMN" -> 17;
                    case "JML_KOREKSI" -> 18;
                    case "JML_PAJAK" -> 19;
                    case "JML_PENGURANG_JASA" -> 20;
                    case "JML_PENGAMBILAN" -> 21;
                    case "JML_NETTO" -> 22;
                    default -> -1;
                };

                if (targetColumn2 != -1) {
                    for (int i = 1; i <= lastRow; i++) {
                        Cell targetCell = sheet2.getRow(i).createCell(targetColumn2);
                        if (sheet.getRow (i).getCell (column)==null){
                            targetCell.setCellValue ("");
                        }else if (sheet.getRow(i).getCell(column).getCellType() == CellType.STRING)
                        {
                            targetCell.setCellValue(sheet.getRow(i).getCell(column).getStringCellValue());
                        } else if (targetColumn2 == 22) {
//                            targetCell.setCellValue(formatter.format(sheet.getRow(i).getCell(column).getNumericCellValue()));
                            targetCell.setCellValue(sheet.getRow(i).getCell(column).getNumericCellValue());
                        } else {
                            targetCell.setCellValue(sheet.getRow(i).getCell(column).getNumericCellValue());
                        }
                    }
                }
            }

            DataFormat dataFormat = workbook.createDataFormat();

            // Create a cell style for the target column
            CellStyle targetColumnColourNetto = workbook.createCellStyle();
            // Set the background color of the cells in the target column to yellow
            targetColumnColourNetto.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            targetColumnColourNetto.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            targetColumnColourNetto.setAlignment (HorizontalAlignment.RIGHT);
            targetColumnColourNetto.setBorderBottom (BorderStyle.THIN);
            targetColumnColourNetto.setBottomBorderColor (IndexedColors.BLACK.getIndex ());
            targetColumnColourNetto.setBorderLeft (BorderStyle.THIN);
            targetColumnColourNetto.setLeftBorderColor (IndexedColors.BLACK.getIndex ());
            targetColumnColourNetto.setBorderRight (BorderStyle.THIN);
            targetColumnColourNetto.setRightBorderColor (IndexedColors.BLACK.getIndex ());
            targetColumnColourNetto.setBorderTop (BorderStyle.THIN);
            targetColumnColourNetto.setTopBorderColor (IndexedColors.BLACK.getIndex ());
            targetColumnColourNetto.setDataFormat (dataFormat.getFormat("_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-"));

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
                sheet2.getRow(downRow).getCell(22).setCellStyle(targetColumnColourNetto);
            }


            workbook.removeSheetAt(0);

            LocalDateTime end = LocalDateTime.now();
            Duration duration = Duration.between(start, end);
            double seconds = duration.toMillis() / 1000.0;
            System.out.printf("B_RincianTindakanJasaDokter Done in %.4f seconds%n", seconds);
        } catch (Exception e) {
            e.printStackTrace();
        }


        try {
            outputStream = new FileOutputStream(fileOutput+"2. RINCIAN TINDAKAN JASA DOKTER.xlsx");
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
