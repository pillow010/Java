package main.java.LaporanRad;

import StylingLaporan.StylerRepo;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.helpers.HSSFRowShifter;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.helpers.RowShifter;
import org.jetbrains.annotations.NotNull;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class A_PertindakanVer2 extends StylerRepo{
    public static void main(String[] args) {
        new main.java.LaporanRad.A_PertindakanVer2 ();

    }
    private Workbook BookPertindakanNew;

    private FileOutputStream outputStream;

    public A_PertindakanVer2(){
        Sheet SheetA = null;
        Sheet sheetB = null;
        File pertindakanNew = new File("C:\\sat work\\test\\rad pertindakan new.xls");
        try {
            POIFSFileSystem poifs = new POIFSFileSystem(pertindakanNew);
            BookPertindakanNew = new HSSFWorkbook(poifs);

//          Make Styling
            CellStyle centerTextCellStyle = BookPertindakanNew.createCellStyle();
            centerTextCellStyle.setAlignment(HorizontalAlignment.CENTER);
            CellStyle AllBorderCellStyle = BookPertindakanNew.createCellStyle();
            AllBorderCellStyle.setBorderBottom(BorderStyle.THIN);
            AllBorderCellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            AllBorderCellStyle.setBorderLeft(BorderStyle.THIN);
            AllBorderCellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            AllBorderCellStyle.setBorderRight(BorderStyle.THIN);
            AllBorderCellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            AllBorderCellStyle.setBorderTop(BorderStyle.THIN);
            AllBorderCellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            CellStyle BorderCenterCellStyle = BookPertindakanNew.createCellStyle();
            BorderCenterCellStyle.setAlignment(HorizontalAlignment.CENTER);
            BorderCenterCellStyle.setBorderBottom(BorderStyle.THIN);
            BorderCenterCellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            BorderCenterCellStyle.setBorderLeft(BorderStyle.THIN);
            BorderCenterCellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            BorderCenterCellStyle.setBorderRight(BorderStyle.THIN);
            BorderCenterCellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            BorderCenterCellStyle.setBorderTop(BorderStyle.THIN);
            BorderCenterCellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());


//          taruh pertindakan new ke sheet 0
            Sheet pertindakan_New_Raw = BookPertindakanNew.getSheetAt(0);
            BookPertindakanNew.setSheetName(0, "pertindakan_New_Raw");

//          tambah sub inst for later use
            pertindakan_New_Raw.getRow(0).createCell(28).setCellValue("SUB INST");
            for (int i = 1; i<= pertindakanNewRawLastRowNum (pertindakan_New_Raw); i++) {
                Row row = pertindakan_New_Raw.getRow(i);
                Cell cell = row.getCell(24);
                if (cell == null){
                    row.createCell(28).setCellValue("RUJUKAN LUAR RS");
                    row.createCell (24).setCellValue ("RUJUKAN LUAR RS");
                } else if (pertindakan_New_Raw.getRow(i).getCell(24).getStringCellValue().equals("HD")) {
                    if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("01")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("HD");
                    }
                } else if (pertindakan_New_Raw.getRow(i).getCell(24).getStringCellValue().equals("RHM")) {
                    if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("01")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("RHM");
                    }
                } else if (pertindakan_New_Raw.getRow(i).getCell(24).getStringCellValue().equals("MCU")) {
                    if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("01")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("MCU");
                    }
                } else if (pertindakan_New_Raw.getRow(i).getCell(24).getStringCellValue().equals("IGD")) {
                    if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("01")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("UMUM");
                    } else if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("02")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("PONEK");
                    }
                } else if (pertindakan_New_Raw.getRow(i).getCell(24).getStringCellValue().equals("IRNA")) {
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("01")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Teratai 1");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("02")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Teratai 2");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("03")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Matahari");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("04")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Tulip");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("05")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Anyelir");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("06")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("ICU");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("07")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("IGD (Mawar)");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("08")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Perinatologi");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("09")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("NICU");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("10")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("VK (Anggrek)");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("11")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("IBS (Sentral)");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("12")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("IBS (IGD)");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("13")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("ISOLASI");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("14")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("TERATAI");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("15")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("ALAMANDA");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("16")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("LILY");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("17")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("CATTLEYA MAGNOLIA");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("18")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("SAKURA");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("19")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("HCU");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("20")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("PICU");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("21")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("ALAMANDA 2");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("22")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("ALAMANDA 3");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("23")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("KEMBANG LILY");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("24")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("LILY 2");
                    }

                }else if (pertindakan_New_Raw.getRow(i).getCell(24).getStringCellValue().equals("IRJ")) {
                    if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("01")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Umum");
                    }
                    if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("02")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Kebidanan dan Kandungan");
                    }
                    if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("03")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Gigi Umum");
                    }
                    if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("04")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Gigi Anak");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("05")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Bedah Umum");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("06")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Bedah Digestif");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("07")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Penyakit Dalam");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("08")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("THT");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("09")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Konservasi Gigi");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("10")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Periodontik");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("11")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Mata");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("12")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Akupuntur");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("13")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Bedah Urologi");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("14")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Bedah Orthopedi");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("15")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Klinik Sahabat");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("16")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Anak");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("17")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Paru");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("18")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("DOTS");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("19")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Anestesi");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("20")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Saraf");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("21")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Psikiatri");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("22")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Kulit dan Kelamin");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("23")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Tumbuh Kembang Anak");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("24")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Geriatri");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("25")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("KIA -KB");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("26")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Gizi");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("27")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Bedah Vaskuler");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("28")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Jantung");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("29")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("Ispa");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("30")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("NEUROLOGI ANAK");
                    }if (pertindakan_New_Raw.getRow(i).getCell(27).getStringCellValue().equals("31")) {
                        pertindakan_New_Raw.getRow(i).createCell(28).setCellValue("BEDAH ONKOLOGI");}

                }

            }

//          buat sheet 1 pertindakan
            Sheet Pertindakan = BookPertindakanNew.createSheet();
            BookPertindakanNew.setSheetName(1, "1 Pertindakan");

//          buat judul dan kasih kotak
            Pertindakan.createRow(5).createCell(0).setCellValue("NO");
            Pertindakan.getRow(5).createCell(1).setCellValue("Nama Tindakan");
            Pertindakan.getRow(5).createCell(2).setCellValue("Jumlah");


            // Perform pivot simulation, and check if it not contains paket
            Map<String, Integer> pivotJumlahTindakan = new HashMap<>();
            for (int i = 1; i <= pertindakanNewRawLastRowNum (pertindakan_New_Raw); i++) {
                Row row = pertindakan_New_Raw.getRow(i);
                String Tindakan = row.getCell(15).getStringCellValue();
                if (!Tindakan.contains("PAKET")) {
                    Integer count = pivotJumlahTindakan.getOrDefault(Tindakan, 0);
                    count++;
                    pivotJumlahTindakan.put(Tindakan, count);
                }
            }

//          Sort any value it contains
            List<Map.Entry<String, Integer>> entriesDoctor = new ArrayList<>(pivotJumlahTindakan.entrySet());
            entriesDoctor.sort(Map.Entry.comparingByKey());
            pivotJumlahTindakan = new LinkedHashMap<>();
            for (Map.Entry<String, Integer> entry : entriesDoctor) {
                pivotJumlahTindakan.put(entry.getKey(), entry.getValue());
            }

//          tulis hasil pivot ke pertindakan, mulai dari row 6
            int startRow = 6;
            int rowNum = startRow;
            for (Map.Entry<String, Integer> entry : pivotJumlahTindakan.entrySet()) {
                Row row = Pertindakan.createRow (rowNum++);
                row.createCell (0).setCellValue (rowNum - 6);
                row.createCell (1).setCellValue (entry.getKey ());
                row.createCell (2).setCellValue (entry.getValue ());
            }


//            System.out.println (startRow);
            for (int rightCell = 0; rightCell<Pertindakan.getRow (rowNum-1).getLastCellNum ();rightCell++){
                Pertindakan.getRow (startRow-1).getCell (rightCell).setCellStyle(BorderCenterCellStyle);
                for (int downRow = startRow; downRow<= pertindakanNewRawLastRowNum (Pertindakan); downRow++){
                    Pertindakan.getRow (downRow).getCell (rightCell).setCellStyle(AllBorderCellStyle);
                }
            }

//          cek per row. sesuaikan width nya
            int columnCountA2 = Pertindakan.getRow (startRow-1).getLastCellNum();
            for (int columnIndex = 0; columnIndex < columnCountA2; columnIndex++) {
                Pertindakan.autoSizeColumn(columnIndex);
            }
//          buat sheet 2 pertindakan
            Sheet Ganjil = BookPertindakanNew.createSheet();
            BookPertindakanNew.setSheetName(2, "Ganjil");

            for (int i=0;i<pertindakanNewRawLastRowNum (pertindakan_New_Raw);i++){
                Ganjil.createRow (i);
            }

            for (int column = 0; column < pertindakan_New_Raw.getRow (0).getLastCellNum (); column++) {
                Cell cell = pertindakan_New_Raw.getRow(0).getCell(column);
                if (cell.getStringCellValue().equals("KD_INST")) {
                    pertindakan_New_Raw.getRow (0).createCell (29).setCellValue ("NOREG");
                    for (int i = 1; i <= pertindakanNewRawLastRowNum (pertindakan_New_Raw); i++) {
                        Cell noReg = pertindakan_New_Raw.getRow(i).createCell (29);
                        noReg.setCellValue(pertindakan_New_Raw.getRow(i).getCell(column).getStringCellValue() +
                                pertindakan_New_Raw.getRow(i).getCell(column + 1).getStringCellValue() +
                                pertindakan_New_Raw.getRow(i).getCell(column + 2).getStringCellValue() +
                                pertindakan_New_Raw.getRow(i).getCell(column + 3).getStringCellValue() +
                                pertindakan_New_Raw.getRow(i).getCell(column + 4).getStringCellValue());
                    }
                }

//              noreg, jenis cara bayar , tanggal, nick inst asal
                String cellValue = pertindakan_New_Raw.getRow (0).getCell (column).getStringCellValue ();
                int targetColumn2 = switch (cellValue) {
                    case "NOREG" -> 0;
                    case "JNS_CR_BYR" -> 1;
                    case "TGL_MASUK" -> 2;
                    case "NICK_INST_ASAL" -> 3;
                    default -> -1;
                };

                if (targetColumn2 != -1) {
                    Ganjil.getRow (0).createCell (targetColumn2).setCellValue (cellValue);
                    for (int i = 1; i < pertindakanNewRawLastRowNum (pertindakan_New_Raw); i++) {
                        String Tindakan = pertindakan_New_Raw.getRow(i).getCell(15).getStringCellValue();
                        if (!Tindakan.contains("PAKET")) {
                            Cell targetCell = Ganjil.getRow (i).createCell (targetColumn2);
                            if (pertindakan_New_Raw.getRow (i).getCell (column).getCellType () == CellType.STRING) {
                                targetCell.setCellValue (pertindakan_New_Raw.getRow (i).getCell (column).getStringCellValue ());
                            } else {
                                targetCell.setCellValue (pertindakan_New_Raw.getRow (i).getCell (column).getNumericCellValue ());

                            }
                        }
                    }
                }
            }

//          buat sheet 3 pertindakan
            Sheet Genap = BookPertindakanNew.createSheet();
            BookPertindakanNew.setSheetName(3, "Genap");

            for (int cell = 0; cell<=pertindakan_New_Raw.getRow (0).getLastCellNum ();cell++) {
                for (int i = 0; i <= pertindakanNewRawLastRowNum (pertindakan_New_Raw)+1; i++) {
                    Genap.createRow (i);
                }
            }
            for (int cell = 0; cell<=pertindakan_New_Raw.getRow (0).getLastCellNum ()-1;cell++) {
                for (int row = 0; row <= pertindakan_New_Raw.getLastRowNum (); row++) {
                    String Tindakan = pertindakan_New_Raw.getRow(row).getCell(15).getStringCellValue();
                    if (!Tindakan.contains("PAKET")) {
                        Cell currentCell = pertindakan_New_Raw.getRow(row).getCell(cell);
                        if (currentCell != null) {
                            if (currentCell.getCellType () == CellType.STRING) {
                                Genap.getRow (row).createCell (cell)
                                        .setCellValue (currentCell.getStringCellValue ());
                            } else {
                                Genap.getRow (row).createCell (cell)
                                        .setCellValue (currentCell.getNumericCellValue ());
                            }
                        }
                    }
                }
            }


            for (int cell = 0; cell<=pertindakan_New_Raw.getRow (0).getLastCellNum ()-1;cell++) {
                for (int row = 0; row <= pertindakan_New_Raw.getLastRowNum (); row++) {
                    String Tindakan = pertindakan_New_Raw.getRow (row).getCell (15).getStringCellValue ();
                    if (Tindakan.contains ("CT Scan")) {
                        Genap.getRow (row).getCell (15).setCellValue ("CT Scan");
//                            CT Scan, USG , RONTGENT, Konsul Dokter Spesialis
                    }
                }
            }

            // V1 shifting row up
//            for (int row = 0; row <= pertindakanNewRawLastRowNum(pertindakan_New_Raw); row++) {
//                boolean isRowBlank = true;
//                for (int cell = 0; cell <= 30; cell++) {
//                    if (Genap.getRow(row).getCell(cell) != null) {
//                        isRowBlank = false;
//                        break;
//                    }
//                }
//                if (isRowBlank) {
//                    Genap.shiftRows(row + 1, row + 1, -1);
//                }
//            }

            // V2 shifting row up
//            System.out.println (pertindakanNewRawLastRowNum(pertindakan_New_Raw));
//            for (int row = 0; row <= pertindakanNewRawLastRowNum(pertindakan_New_Raw); row++) {
//                boolean isRowBlank = true;
//                System.out.println (row);
//                for (int cell = 0; cell <= 30; cell++) {
//                    if (Genap.getRow(row).getCell(cell) != null) {
//                        System.out.println (row + " "+ cell);
//                        isRowBlank = false;
//                        break;
//                    }
//                }
//                int rowShouldBeShifted = row+1;
//                if (isRowBlank) {
//                    Genap.shiftRows(rowShouldBeShifted, rowShouldBeShifted, -1);
//                }
//            }

            // V3 shifting row up
            for (int row = 0; row <= pertindakanNewRawLastRowNum(pertindakan_New_Raw); row++) {
                boolean isRowBlank = true;
                System.out.println ("up " + row);
                System.out.println (isRowBlank);
                for (int cell = 0; cell <= 30; cell++) {
                    if (Genap.getRow(row).getCell(cell) != null) {
                        System.out.println (row + " "+ cell);
                        isRowBlank = false;
                        break;
                    }
                }
                int rowShouldBeShifted = row + 1;
                if (isRowBlank) {
                    Genap.shiftRows(rowShouldBeShifted, rowShouldBeShifted, -1);
                }
            }
                //v4
//            for (int row = 0; row <= pertindakanNewRawLastRowNum(pertindakan_New_Raw); row++) {
//                boolean isRowBlank = true;
//                System.out.println (row);
//                for (int cell = 0; cell <= 30; cell++) {
//                    if (Genap.getRow(row).getCell(cell) != null) {
//                        System.out.println (row + " "+ cell);
//                        isRowBlank = false;
//                        break;
//                    }
//                }
//                int rowShouldBeShifted = row + 1;
//                if (isRowBlank) {
//                    Genap.shiftRows(rowShouldBeShifted, rowShouldBeShifted, -1);
//                }
//            }
            //V5
//            int row = 0;
//            while (row <= pertindakanNewRawLastRowNum(pertindakan_New_Raw)) {
//                boolean isRowBlank = true;
//                for (int cell = 0; cell <= 30; cell++) {
//                    if (Genap.getRow(row).getCell(cell) != null) {
//                        isRowBlank = false;
//                        break;
//                    }
//                }
//                int rowShouldBeShifted = row + 1;
//                if (isRowBlank) {
//                    Genap.shiftRows(rowShouldBeShifted, rowShouldBeShifted, -1);
//                } else {
//                    row++;
//                }
//            }







//          buat sheet 2 pertindakan
            Sheet Tindakan_crByr_Hari = BookPertindakanNew.createSheet();
            BookPertindakanNew.setSheetName(4, "2.Jml tndakan per cr Byr pr hri");


















        } catch (Exception e) {
            e.printStackTrace();
        }


        try {
            outputStream = new FileOutputStream("pertindakanNew.xlsx");
            BookPertindakanNew.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (BookPertindakanNew != null) {
                    BookPertindakanNew.close();
                }
                if (outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static int pertindakanNewRawLastRowNum(Sheet pertindakan_New_Raw) {
        return pertindakan_New_Raw.getLastRowNum ();
    }

    private static void removeDuplicates(@NotNull Sheet sheet) {
        Set<String> uniqueRows = new HashSet<>();
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            Row currentRow = sheet.getRow(i);
            if (currentRow == null) {
                continue;
            }
            StringBuilder sb = new StringBuilder();
            for (int j = 0; j < currentRow.getLastCellNum(); j++) {
                Cell currentCell = currentRow.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                switch (currentCell.getCellType ()) {
                    case STRING -> sb.append (currentCell.getStringCellValue ());
                    case NUMERIC -> sb.append (currentCell.getNumericCellValue ());
                }
            }
            String rowAsString = sb.toString();
            if (uniqueRows.contains(rowAsString)) {
                sheet.removeRow(currentRow);
                i--;
                lastRowNum--;
            } else {
                uniqueRows.add(rowAsString);
            }
        }
    }

}