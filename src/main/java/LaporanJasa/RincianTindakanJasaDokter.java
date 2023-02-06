package main.java.LaporanJasa;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class RincianTindakanJasaDokter {
    public static void main(String[] args) {
        new RincianTindakanJasaDokter();
    }

    private Workbook workbook;
    private FileOutputStream outputStream;


    public RincianTindakanJasaDokter() {
        Sheet sheet = null;
        File inputFS = new File("C:\\sat work\\test\\c) LAPORAN PENERIMAAN JASA PELAYANAN PER TINDAKAN1.xls");
        try {
            POIFSFileSystem poifs = new POIFSFileSystem(inputFS);
            workbook = new HSSFWorkbook(poifs);

            sheet = workbook.getSheetAt(0);
            Sheet sheet2 = workbook.createSheet();
            workbook.setSheetName(1, "2. RINCIAN TINDAKAN JASA DOKTER");
            int lastColumn = sheet.getRow(0).getLastCellNum();
            int lastRow = sheet.getLastRowNum();
            System.out.println("Last Column: " + lastColumn);
            System.out.println("Last Row: "+lastRow);

            sheet.getRow(0).createCell(60).setCellValue("NAMA PASIEN");
            sheet.getRow(0).createCell(61).setCellValue("NORM");

            for (int column = 0; column <= lastColumn - 1; column++) {
                Cell cell = sheet.getRow(0).getCell(column);
                String cellValue = cell.getStringCellValue();
                int targetColumn = switch (cellValue) {
                    case "NAMA_PASIEN" -> 60;
                    case "NORM" -> 61;
                    case "NO_REG" -> 62;
                    case "KET_INST" -> 63;
                    case "KET_SUB_INST" -> 64;
                    case "KET_DTL_SUB_INST" -> 65;
                    case "NAMA_DOKTER" -> 66;
                    case "POSISI" -> 67;
                    case "TGL_TINDAKAN" -> 68;
                    case "NM_TINDAKAN" -> 69;
                    case "RUANG_RAWAT" -> 70;
                    case "PAKET_JAMINAN" -> 71;
                    case "JASA_PELAYANAN_TARIF" -> 72;
                    case "JASA_PELAYANAN_JAMIN" -> 73;
                    case "JML_PENDAPATAN" -> 74;
                    case "JML_PENERIMAAN_TUNAI" -> 75;
                    case "JML_PENERIMAAN_PIUTANG" -> 76;
                    case "JML_PENERIMAAN_JMN" -> 77;
                    case "JML_KOREKSI" -> 78;
                    case "JML_PAJAK" -> 79;
                    case "JML_PENGURANG_JASA" -> 80;
                    case "JML_PENGAMBILAN" -> 81;
                    case "JML_NETTO" -> 82;
                    default -> -1;
                };


                if (targetColumn != -1) {
                    for (int i = 1; i <= lastRow; i++) {
                        Cell targetCell = sheet.getRow(i).createCell(targetColumn);
                        sheet.getRow(0).createCell(targetColumn).setCellValue(cellValue);
                        if (sheet.getRow(i).getCell(column).getCellType() == CellType.STRING)

                        {
                            targetCell.setCellValue(sheet.getRow(i).getCell(column).getStringCellValue());
                        } else
                            targetCell.setCellValue(sheet.getRow(i).getCell(column).getNumericCellValue());

                    }
                }

//                beri nama A1 "NOREG
                sheet.getRow(0).createCell(62).setCellValue("NOREG");

//              jika cell mengandung "KD_INST" concat jadi noreg
                if (cell.getStringCellValue().equals("KD_INST")) {
                    for (int i = 1; i <= lastRow; i++) {
                        Cell noReg = sheet.getRow(i).createCell(62);
                        noReg.setCellValue(sheet.getRow(i).getCell(column).getStringCellValue() +
                                sheet.getRow(i).getCell(column + 1).getStringCellValue() +
                                sheet.getRow(i).getCell(column + 2).getStringCellValue() +
                                sheet.getRow(i).getCell(column + 3).getStringCellValue() +
                                sheet.getRow(i).getCell(column + 4).getStringCellValue());
                    }
                }



            }
//



//            for (int i = 0; i < 50; i++) {
//                sheet.shiftColumns(i + 1, lastColumn, -1);
//            }
        } catch (Exception e) {
            e.printStackTrace();
        }
//        sheet = workbook.getSheetAt(0);
//        int lastColumn = sheet.getRow(0).getLastCellNum();
//        for (int i = 0; i < 50; i++) {
//            sheet.shiftColumns(i + 1, lastColumn, -1);
//        }


        try {
            outputStream = new FileOutputStream("2. RINCIAN TINDAKAN JASA DOKTER.xlsx");
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
