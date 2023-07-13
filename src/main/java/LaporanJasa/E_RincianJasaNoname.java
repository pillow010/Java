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

public class E_RincianJasaNoname {

    public static void main(String[] args) {
//        new E_RincianJasaNoname ();
    }

    private Workbook workbook;
    private FileOutputStream outputStream;

    public E_RincianJasaNoname() {
        // Determine the file extension based on the availability of XLSX and XLS files
        File xlsxFile = new File (fileSource + "d) LAPORAN PENERIMAAN JASA PELAYANAN TANPA PEMILIK.xlsx");
        File xlsFile = new File (fileSource + "d) LAPORAN PENERIMAAN JASA PELAYANAN TANPA PEMILIK.xls");
        File jasaTanpaPemilik;
        if (xlsxFile.exists ()) {
            jasaTanpaPemilik = xlsxFile;
        } else if (xlsFile.exists ()) {
            jasaTanpaPemilik = xlsFile;
        } else {
            System.out.println ("File not found: " + fileSource + "d) LAPORAN PENERIMAAN JASA PELAYANAN TANPA PEMILIK");
            return;
        }

        System.out.println ("E_RincianJasaNoname is starting");
        try {
            LocalDateTime start = LocalDateTime.now ();
            FileInputStream inputStream = new FileInputStream (jasaTanpaPemilik);
            workbook = WorkbookFactory.create (inputStream);

            Sheet sheet = workbook.getSheetAt (0);
            Sheet sheet2 = workbook.createSheet ();

            DataFormat dataFormat = workbook.createDataFormat ();

            CellStyle totalStyle = workbook.createCellStyle ();
            totalStyle.setAlignment (HorizontalAlignment.RIGHT);
            totalStyle.setBorderBottom (BorderStyle.THIN);
            totalStyle.setBottomBorderColor (IndexedColors.BLACK.getIndex ());
            totalStyle.setBorderLeft (BorderStyle.THIN);
            totalStyle.setLeftBorderColor (IndexedColors.BLACK.getIndex ());
            totalStyle.setBorderRight (BorderStyle.THIN);
            totalStyle.setRightBorderColor (IndexedColors.BLACK.getIndex ());
            totalStyle.setBorderTop (BorderStyle.THIN);
            totalStyle.setTopBorderColor (IndexedColors.BLACK.getIndex ());
            totalStyle.setDataFormat (dataFormat.getFormat ("_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-"));
            totalStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            totalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            workbook.setSheetName (1, "5. RINCIAN JASA NONAME ");
            int lastColumn = sheet.getRow (0).getLastCellNum ();
            int lastRow = sheet.getLastRowNum ();
            for (int row = 1; row <= lastRow; row++) {
                sheet2.createRow (row);
            }

            Row row = sheet2.getRow (0);
            if (row == null) {
                row = sheet2.createRow (0);
            }
            row.createCell (0).setCellValue ("NAMA PASIEN");
            row.createCell (1).setCellValue ("NORM");
            row.createCell (2).setCellValue ("NO REG");
            row.createCell (3).setCellValue ("POSISI NONAME");
            row.createCell (4).setCellValue ("TGL TINDAKAN");
            row.createCell (5).setCellValue ("NM TINDAKAN");
            row.createCell (6).setCellValue ("PAKET JAMINAN");
            row.createCell (7).setCellValue ("RUANG RAWAT");
            row.createCell (8).setCellValue ("JML PENDAPATAN");
            row.createCell (9).setCellValue ("JML PENERIMAAN TUNAI");
            row.createCell (10).setCellValue ("JML PENERIMAAN PIUTANG");
            row.createCell (11).setCellValue ("JML PENERIMAAN JMN");
            row.createCell (12).setCellValue ("JML KOREKSI");
            row.createCell (13).setCellValue ("JML PAJAK");
            row.createCell (14).setCellValue ("JML PENGURANG JASA");
            row.createCell (15).setCellValue ("JML NETTO");
            row.createCell (16).setCellValue ("NAMA DOKTER DPJP");
            row.createCell (17).setCellValue ("KET INST");
            row.createCell (18).setCellValue ("KET SUB INST");
            row.createCell (19).setCellValue ("KET DTL SUB INST");


            for (int column = 0; column <= lastColumn - 1; column++) {
//          jika cell mengandung "KD_INST" concat jadi noreg
                Cell cell = sheet.getRow (0).getCell (column);
                if (sheet.getRow (0).getCell (column).getStringCellValue ().equals ("KD_INST")) {
                    for (int i = 1; i <= lastRow; i++) {
                        if (sheet.getRow (i).getCell (column) == null) {
                            sheet2.getRow (i).createCell (column).setCellValue ("");
                        } else {
                            Cell noReg = sheet2.getRow (i).createCell (2);
                            noReg.setCellValue (sheet.getRow (i).getCell (column).getStringCellValue () +
                                    sheet.getRow (i).getCell (column + 1).getStringCellValue () +
                                    sheet.getRow (i).getCell (column + 2).getStringCellValue () +
                                    sheet.getRow (i).getCell (column + 3).getStringCellValue () +
                                    sheet.getRow (i).getCell (column + 4).getStringCellValue ());
                        }
                    }
                }

                String cellValue = cell.getStringCellValue ();
                int targetColumn2 = switch (cellValue) {
                    case "NAMA_PASIEN" -> 0;
                    case "NORM" -> 1;
                    case "NO_REG" -> 2;
                    case "POSISI_NONAME" -> 3;
                    case "TGL_TINDAKAN" -> 4;
                    case "NM_TINDAKAN" -> 5;
                    case "PAKET_JAMINAN" -> 6;
                    case "RUANG_RAWAT" -> 7;
                    case "JML_PENDAPATAN" -> 8;
                    case "JML_PENERIMAAN_TUNAI" -> 9;
                    case "JML_PENERIMAAN_PIUTANG" -> 10;
                    case "JML_PENERIMAAN_JMN" -> 11;
                    case "JML_KOREKSI" -> 12;
                    case "JML_PAJAK" -> 13;
                    case "JML_PENGURANG_JASA" -> 14;
                    case "JML_NETTO" -> 15;
                    case "NAMA_DOKTER_DPJP" -> 16;
                    case "KET_INST" -> 17;
                    case "KET_SUB_INST" -> 18;
                    case "KET_DTL_SUB_INST" -> 19;
                    default -> -1;
                };

                if (targetColumn2 != -1) {
                    for (int i = 1; i <= lastRow; i++) {
                        Cell targetCell = sheet2.getRow (i).createCell (targetColumn2);
                        if (sheet.getRow (i).getCell (column) == null) {
                            targetCell.setCellValue ("");
                        } else if (sheet.getRow (i).getCell (column).getCellType () == CellType.STRING) {
                            targetCell.setCellValue (sheet.getRow (i).getCell (column).getStringCellValue ());
                        } else if (targetColumn2 == 15) {
//                            targetCell.setCellValue(formatter.format(sheet.getRow(i).getCell(column).getNumericCellValue()));
                            targetCell.setCellValue (sheet.getRow (i).getCell (column).getNumericCellValue ());
                        } else {
                            targetCell.setCellValue (sheet.getRow (i).getCell (column).getNumericCellValue ());
                        }
                    }
                }

            }
            int columnCountA2 = sheet2.getRow (0).getLastCellNum ();
            for (int columnIndex = 0; columnIndex < columnCountA2; columnIndex++) {
                sheet2.autoSizeColumn (columnIndex);
            }
            CellStyle centerTextStyle = workbook.createCellStyle ();
            centerTextStyle.setAlignment (HorizontalAlignment.CENTER);
            int lastCellA2 = sheet2.getRow (0).getLastCellNum ();
            for (int title = 0; title < lastCellA2; title++) {
                sheet2.getRow (0).getCell (title).setCellStyle (centerTextStyle);
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
            int lastColumnIndex = sheet2.getRow (5).getLastCellNum ();

            // Apply header style and border to the header row and auto-size columns
            for (int rightCell = 0; rightCell < lastColumnIndex; rightCell++) {
                sheet2.getRow (0).getCell (rightCell).setCellStyle (BorderCenterCellStyle);
                sheet2.autoSizeColumn (rightCell);
            }

            // Apply cell styles and borders to the data rows
            for (int rightCell = 0; rightCell < lastColumnIndex; rightCell++) {
                for (int downRow = 1; downRow <= sheet2.getLastRowNum (); downRow++) {
                    sheet2.getRow (downRow).getCell (rightCell).setCellStyle (AllBorderCellStyle);
                }
            }

            // Apply the target column style to column 22 for all data rows
            for (int downRow = 1; downRow <= sheet2.getLastRowNum (); downRow++) {
                sheet2.getRow (downRow).getCell (15).setCellStyle (totalStyle);
            }
            workbook.removeSheetAt (0);
            LocalDateTime end = LocalDateTime.now();
            Duration duration = Duration.between(start, end);
            double seconds = duration.toMillis() / 1000.0;
            System.out.printf("E_RincianJasaNoname Done in %.4f seconds%n", seconds);

        } catch (Exception e) {
            e.printStackTrace ();
        }


        try {
            outputStream = new FileOutputStream (fileOutput + "5. RINCIAN JASA NONAME.xlsx");
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
