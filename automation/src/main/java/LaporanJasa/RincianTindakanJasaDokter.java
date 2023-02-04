package LaporanJasa;

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
            int lastColumn = sheet.getRow(0).getLastCellNum();
            int lastRow = sheet.getLastRowNum();
            System.out.println("Last Column: " + lastColumn);
            System.out.println("Last Row: "+lastRow);

            sheet.getRow(0).createCell(60).setCellValue("NAMA PASIEN");
            sheet.getRow(0).createCell(61).setCellValue("NORM");

//            ## V1
//            for (int column=0; column<=lastColumn-1; column++) {
//                Cell cell = sheet.getRow(0).getCell(column);
////                System.out.println(cell.getStringCellValue());
//                if (cell.getStringCellValue().equals("NAMA_PASIEN")) {
//                    for (int i = 1; i <= lastRow; i++) {
//                        Cell NamaPasien = sheet.getRow(i).createCell(60);
//                        NamaPasien.setCellValue(sheet.getRow(i).getCell(column).getStringCellValue());
//                    }
//                }
//                if (cell.getStringCellValue().equals("NORM")) {
//                    for (int i = 1; i <= lastRow; i++) {
//                        Cell NamaPasien = sheet.getRow(i).createCell(61);
//                        NamaPasien.setCellValue(sheet.getRow(i).getCell(column).getStringCellValue());
//                    }
//                }
//                NORM

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
                        if (sheet.getRow(i).getCell(column).getCellType() == CellType.STRING)

                        {
                            targetCell.setCellValue(sheet.getRow(i).getCell(column).getStringCellValue());
                        } else
                            targetCell.setCellValue(sheet.getRow(i).getCell(column).getNumericCellValue());

                    }
                }
            }
//                NO REG
//                KET INST
//                KET SUB INST
//                KET DTL SUB INST
//                NAMA DOKTER
//                POSISI
//                TGL TINDAKAN
//                NM TINDAKAN
//                RUANG RAWAT
//                PAKET JAMINAN
//                JASA PELAYANAN TARIF
//                JASA PELAYANAN JAMIN
//                JML PENDAPATAN
//                JML PENERIMAAN TUNAI
//                JML PENERIMAAN PIUTANG
//                JML PENERIMAAN JMN
//                JML KOREKSI
//                JML PAJAK
//                JML PENGURANG JASA
//                JML PENGAMBILAN
//                JML NETTO



//            }



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
