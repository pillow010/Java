package main.java.LaporanJasa;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class D_RekapPasienJasaUnit {

    public static void main(String[] args) {
    }

    private Workbook workbook;
    private FileOutputStream outputStream;


    public D_RekapPasienJasaUnit(){

        Sheet sheet = null;
        File inputFS = new File("C:\\sat work\\test\\a) LAPORAN REKAP PENERIMAAN JASA UNIT PER PASIEN1.xls");
        try {
            POIFSFileSystem poifs = new POIFSFileSystem(inputFS);
            workbook = new HSSFWorkbook(poifs);

            sheet = workbook.getSheetAt(0);
            Sheet sheet2 = workbook.createSheet();

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
                        if (sheet.getRow(i).getCell(column).getCellType() == CellType.STRING)
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
            workbook.removeSheetAt(0);

        } catch (Exception e) {
            e.printStackTrace();
        }


        try {
            outputStream = new FileOutputStream("4. REKAP PASIEN JASA UNIT.xlsx");
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
