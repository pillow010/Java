package main.java.LaporanJasa;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

public class B_RincianTindakanJasaDokter {
    public static void main(String[] args) {
        new B_RincianTindakanJasaDokter();
    }

    private Workbook workbook;
    private FileOutputStream outputStream;


    public B_RincianTindakanJasaDokter() {
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
//          jika cell mengandung "KD_INST" concat jadi noreg
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
                        if (sheet.getRow(i).getCell(column).getCellType() == CellType.STRING)
                        {
                            targetCell.setCellValue(sheet.getRow(i).getCell(column).getStringCellValue());
                        } else {
                            targetCell.setCellValue(sheet.getRow(i).getCell(column).getNumericCellValue());

                        }
                    }
                }

            }

            workbook.removeSheetAt(0);

        } catch (Exception e) {
            e.printStackTrace();
        }


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
