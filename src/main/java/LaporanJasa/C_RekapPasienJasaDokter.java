package LaporanJasa;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

public class C_RekapPasienJasaDokter {
    public static void main(String[] args) {
//        new C_RekapPasienJasaDokter();
    }

    private Workbook workbook;
    private FileOutputStream outputStream;


    public C_RekapPasienJasaDokter() {
        File inputFS = new File("C:\\sat work\\test\\b) LAPORAN REKAP PENERIMAAN JASA PELAYANAN PER PASIEN1.xls");
        System.out.println ("C_RekapPasienJasaDokter is starting");
        try {
            POIFSFileSystem poifs = new POIFSFileSystem(inputFS);
            workbook = new HSSFWorkbook(poifs);

            Sheet sheet = workbook.getSheetAt(0);
            Sheet sheet2 = workbook.createSheet();

            workbook.setSheetName(1, "3. REKAP PASIEN JASA DOKTER");
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
            row.createCell(3).setCellValue("TGL REG");
            row.createCell(4).setCellValue("KET INST");
            row.createCell(5).setCellValue("KET SUB INST");
            row.createCell(6).setCellValue("KET DTL SUB INST");
            row.createCell(7).setCellValue("NAMA DOKTER RSF");
            row.createCell(8).setCellValue("NAMA BANK");
            row.createCell(9).setCellValue("NO REKENING");
            row.createCell(10).setCellValue("JML PENERIMAAN");
            row.createCell(11).setCellValue("JML KOREKSI");
            row.createCell(12).setCellValue("JML PAJAK");
            row.createCell(13).setCellValue("JML PENGAMBILAN");
            row.createCell(14).setCellValue("JML NETTO");



            for (int column = 0; column <= lastColumn - 1; column++) {
//          jika cell mengandung "KD_INST" concat jadi noreg
                Cell cell = sheet.getRow(0).getCell(column);

                String cellValue = cell.getStringCellValue();
                int targetColumn2 = switch (cellValue) {
                    case "NAMA_PASIEN" -> 0;
                    case "NORM" -> 1;
                    case "NOREG" -> 2;
                    case "TGL_REG" -> 3;
                    case "KET_INST" -> 4;
                    case "KET_SUB_INST" -> 5;
                    case "KET_DTL_SUB_INST" -> 6;
                    case "NM_DOKTER_RSF" -> 7;
                    case "NAMA_BANK" -> 8;
                    case "NO_REKENING" -> 9;
                    case "JML_PENERIMAAN" -> 10;
                    case "JML_KOREKSI" -> 11;
                    case "JML_PAJAK" -> 12;
                    case "JML_PENGAMBILAN" -> 13;
                    case "JML_NETTO" -> 14;
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

            //buat header center kemudian border semuanya ps. use'<' because return 2 but there is 0, and 1. no number 2.
            for (int rightCell = 0; rightCell < sheet2.getRow (5).getLastCellNum (); rightCell++) {
                sheet2.getRow (0).getCell (rightCell).setCellStyle (BorderCenterCellStyle);
                sheet2.autoSizeColumn (rightCell);
                for (int downRow = 1; downRow <= sheet2.getLastRowNum (); downRow++) {
                    sheet2.getRow (downRow).getCell (rightCell).setCellStyle (AllBorderCellStyle);
                }
            }

            workbook.removeSheetAt(0);
            System.out.println ("C_RekapPasienJasaDokter Done");

        } catch (Exception e) {
            e.printStackTrace();
        }


        try {
            outputStream = new FileOutputStream("3. REKAP PASIEN JASA DOKTER.xlsx");
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
