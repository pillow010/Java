package main.java.LaporanJasa;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class E_RincianJasaNoname {

    public static void main(String[] args) {
        new E_RincianJasaNoname();
    }
    private Workbook workbook;
    private FileOutputStream outputStream;

    public E_RincianJasaNoname(){

        Sheet sheet = null;
        File inputFS = new File("C:\\sat work\\test\\d) LAPORAN PENERIMAAN JASA PELAYANAN TANPA PEMILIK1.xls");
        try {
            POIFSFileSystem poifs = new POIFSFileSystem(inputFS);
            workbook = new HSSFWorkbook(poifs);

//            XSSFWorkbook workbook = new XSSFWorkbook(inputFS);

            sheet = workbook.getSheetAt(0);
            Sheet sheet2 = workbook.createSheet();

            workbook.setSheetName(1, "5. RINCIAN JASA NONAME");
            int lastColumn = sheet.getRow(0).getLastCellNum();
            int lastRow = sheet.getLastRowNum();
            System.out.println("Last Column: " + lastColumn);
            System.out.println("Last Row: " + lastRow);
            for (int row = 1; row <= lastRow; row++) {
                sheet2.createRow(row);
            }

            Row row = sheet2.getRow(0);
            if (row == null) {
                row = sheet2.createRow(0);
            }
            row.createCell(0).setCellValue("NAMA PASIEN");
            row.createCell(1).setCellValue("NORM");
            row.createCell(2).setCellValue("NO REG");
            row.createCell(3).setCellValue("POSISI NONAME");
            row.createCell(4).setCellValue("TGL TINDAKAN");
            row.createCell(5).setCellValue("NM TINDAKAN");
            row.createCell(6).setCellValue("PAKET JAMINAN");
            row.createCell(7).setCellValue("RUANG RAWAT");
            row.createCell(8).setCellValue("JML PENDAPATAN");
            row.createCell(9).setCellValue("JML PENERIMAAN TUNAI");
            row.createCell(10).setCellValue("JML PENERIMAAN PIUTANG");
            row.createCell(11).setCellValue("JML PENERIMAAN JMN");
            row.createCell(12).setCellValue("JML KOREKSI");
            row.createCell(13).setCellValue("JML PAJAK");
            row.createCell(14).setCellValue("JML PENGURANG JASA");
            row.createCell(15).setCellValue("JML NETTO");
            row.createCell(16).setCellValue("NAMA DOKTER DPJP");
            row.createCell(17).setCellValue("KET INST");
            row.createCell(18).setCellValue("KET SUB INST");
            row.createCell(19).setCellValue("KET DTL SUB INST");




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
            outputStream = new FileOutputStream("5. RINCIAN JASA NONAME.xlsx");
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
