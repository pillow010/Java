package LaporanRadiologi;

import StylingLaporan.StylerRepo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



import java.io.*;
import java.util.*;

import java.util.stream.IntStream;

public class RadHalfDone extends StylerRepo{
    public static void main(String[] args) {
        new RadHalfDone ();

    }
    private Workbook BookPertindakanNew;

    private FileOutputStream outputStream;

    String fileNamePertindakanNew = "23 02 rad pertindakan new";
    String fileNameMonitoringf1 = "23 02 rad monitorring f1";

    public RadHalfDone(){
        try {
            InputStream pertindakanNew = new FileInputStream ("C:\\sat work\\test\\"+ fileNamePertindakanNew +".xlsx");
            BookPertindakanNew = new XSSFWorkbook (pertindakanNew);


//          Make Styling
            CellStyle centerTextCellStyle = BookPertindakanNew.createCellStyle ();
            centerTextCellStyle.setAlignment (HorizontalAlignment.CENTER);
            CellStyle AllBorderCellStyle = BookPertindakanNew.createCellStyle ();
            AllBorderCellStyle.setBorderBottom (BorderStyle.THIN);
            AllBorderCellStyle.setBottomBorderColor (IndexedColors.BLACK.getIndex ());
            AllBorderCellStyle.setBorderLeft (BorderStyle.THIN);
            AllBorderCellStyle.setLeftBorderColor (IndexedColors.BLACK.getIndex ());
            AllBorderCellStyle.setBorderRight (BorderStyle.THIN);
            AllBorderCellStyle.setRightBorderColor (IndexedColors.BLACK.getIndex ());
            AllBorderCellStyle.setBorderTop (BorderStyle.THIN);
            AllBorderCellStyle.setTopBorderColor (IndexedColors.BLACK.getIndex ());
            CellStyle BorderCenterCellStyle = BookPertindakanNew.createCellStyle ();
            BorderCenterCellStyle.setAlignment (HorizontalAlignment.CENTER);
            BorderCenterCellStyle.setBorderBottom (BorderStyle.THIN);
            BorderCenterCellStyle.setBottomBorderColor (IndexedColors.BLACK.getIndex ());
            BorderCenterCellStyle.setBorderLeft (BorderStyle.THIN);
            BorderCenterCellStyle.setLeftBorderColor (IndexedColors.BLACK.getIndex ());
            BorderCenterCellStyle.setBorderRight (BorderStyle.THIN);
            BorderCenterCellStyle.setRightBorderColor (IndexedColors.BLACK.getIndex ());
            BorderCenterCellStyle.setBorderTop (BorderStyle.THIN);
            BorderCenterCellStyle.setTopBorderColor (IndexedColors.BLACK.getIndex ());


//          taruh pertindakan new ke sheet 0
            Sheet pertindakan_New_Raw = BookPertindakanNew.getSheetAt (0);
            BookPertindakanNew.setSheetName (0, "pertindakan_New_Raw");
            System.out.println ("00. Doing " + BookPertindakanNew.getSheetAt (0).getSheetName ());

//          tambah sub inst for later use
            pertindakan_New_Raw.getRow (0).createCell (28).setCellValue ("SUB INST");
            for (int i = 1; i <= pertindakanNewRawLastRowNum (pertindakan_New_Raw); i++) {
                Row row = pertindakan_New_Raw.getRow (i);
                Cell cell = row.getCell (24);
                if (cell == null) {
                    row.createCell (28).setCellValue ("RUJUKAN LUAR RS");
                    row.createCell (24).setCellValue ("RUJUKAN LUAR RS");
                } else if (pertindakan_New_Raw.getRow (i).getCell (24).getStringCellValue ().equals ("HD")) {
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("01")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("HD");
                    }
                } else if (pertindakan_New_Raw.getRow (i).getCell (24).getStringCellValue ().equals ("RHM")) {
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("01")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("RHM");
                    }
                } else if (pertindakan_New_Raw.getRow (i).getCell (24).getStringCellValue ().equals ("MCU")) {
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("01")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("MCU");
                    }
                } else if (pertindakan_New_Raw.getRow (i).getCell (24).getStringCellValue ().equals ("IGD")) {
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("01")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("UMUM");
                    } else if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("02")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("PONEK");
                    }
                } else if (pertindakan_New_Raw.getRow (i).getCell (24).getStringCellValue ().equals ("IRNA")) {
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

                } else if (pertindakan_New_Raw.getRow (i).getCell (24).getStringCellValue ().equals ("IRJ")) {
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("01")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Umum");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("02")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Kebidanan dan Kandungan");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("03")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Gigi Umum");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("04")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Gigi Anak");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("05")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Bedah Umum");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("06")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Bedah Digestif");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("07")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Penyakit Dalam");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("08")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("THT");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("09")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Konservasi Gigi");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("10")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Periodontik");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("11")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Mata");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("12")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Akupuntur");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("13")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Bedah Urologi");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("14")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Bedah Orthopedi");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("15")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Klinik Sahabat");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("16")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Anak");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("17")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Paru");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("18")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("DOTS");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("19")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Anestesi");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("20")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Saraf");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("21")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Psikiatri");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("22")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Kulit dan Kelamin");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("23")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Tumbuh Kembang Anak");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("24")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Geriatri");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("25")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("KIA -KB");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("26")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Gizi");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("27")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Bedah Vaskuler");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("28")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Jantung");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("29")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("Ispa");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("30")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("NEUROLOGI ANAK");
                    }
                    if (pertindakan_New_Raw.getRow (i).getCell (27).getStringCellValue ().equals ("31")) {
                        pertindakan_New_Raw.getRow (i).createCell (28).setCellValue ("BEDAH ONKOLOGI");
                    }

                }

            }
//          add noreg
            pertindakan_New_Raw.getRow (0).createCell (29).setCellValue ("NOREG");
            pertindakan_New_Raw.getRow (0).createCell (30).setCellValue ("NOREGTINDAKAN");
            for (int column = 0; column < pertindakan_New_Raw.getRow (0).getLastCellNum (); column++) {
                Cell cell = pertindakan_New_Raw.getRow (0).getCell (column);
                if (cell.getStringCellValue ().equals ("KD_INST")) {
                    for (int i = 1; i <= pertindakanNewRawLastRowNum (pertindakan_New_Raw); i++) {
                        Cell noReg = pertindakan_New_Raw.getRow (i).createCell (29);
                        Cell noRegTindakan = pertindakan_New_Raw.getRow (i).createCell (30);
                        noReg.setCellValue (pertindakan_New_Raw.getRow (i).getCell (column).getStringCellValue () +
                                pertindakan_New_Raw.getRow (i).getCell (column + 1).getStringCellValue () +
                                pertindakan_New_Raw.getRow (i).getCell (column + 2).getStringCellValue () +
                                pertindakan_New_Raw.getRow (i).getCell (column + 3).getStringCellValue () +
                                pertindakan_New_Raw.getRow (i).getCell (column + 4).getStringCellValue ());

                        noRegTindakan.setCellValue (pertindakan_New_Raw.getRow (i).getCell (29).getStringCellValue ()
                                + pertindakan_New_Raw.getRow (i).getCell (15).getStringCellValue ()
                        );
                    }
                }
            }
            System.out.println ("00. " + BookPertindakanNew.getSheetAt (0).getSheetName () + " Complete");


//          buat sheet 1 Ganjil
            Sheet Ganjil = BookPertindakanNew.createSheet ();
            BookPertindakanNew.setSheetName (1, "Ganjil");
            System.out.println ("01. " + BookPertindakanNew.getSheetAt (1).getSheetName () + " Start");


            Set<String> uniqueValues = new HashSet<> ();
            for (int row = 1; row <= pertindakan_New_Raw.getLastRowNum (); row++) {
                if (pertindakan_New_Raw.getRow (row) != null) { // check if row is not empty
                    Cell cell = pertindakan_New_Raw.getRow (row).getCell (29);
                    if (cell != null) { // check if cell is not empty
                        String cellValue = cell.getStringCellValue ();
                        if (!cellValue.isBlank ()) { // check if cell value is not blank
                            uniqueValues.add (cellValue);
                        }
                    }
                }
            }


            for (int i = 0; i <= uniqueValues.size (); i++) {
                Ganjil.createRow (i);
            }
            Ganjil.getRow (0).createCell (0).setCellValue ("NOREG");
            Ganjil.getRow (0).createCell (1).setCellValue ("JENIS CARA BAYAR");
            Ganjil.getRow (0).createCell (2).setCellValue ("TANGGAL MASUK");
            Ganjil.getRow (0).createCell (3).setCellValue ("NIC INST ASAL");

            List<String> sortedValues = uniqueValues.stream ().sorted ().toList ();

            IntStream.range (0, sortedValues.size ())
                    .forEach (i -> {
                        String value = sortedValues.get (i);
                        Ganjil.getRow (i + 1).createCell (0).setCellValue (value);
                    });

            for (int row = 1; row <= sortedValues.size (); row++) {
                String cellValue = Ganjil.getRow (row).getCell (0).getStringCellValue ();
                for (int pertRow = 1; pertRow <= pertindakan_New_Raw.getLastRowNum (); pertRow++) {
                    String pertCellValue = pertindakan_New_Raw.getRow (pertRow).getCell (29).getStringCellValue ();
                    if (cellValue.equals (pertCellValue)) {
                        String JnsCrByr = pertindakan_New_Raw.getRow (pertRow).getCell (8).getStringCellValue ();
                        String TglMsk = pertindakan_New_Raw.getRow (pertRow).getCell (9).getStringCellValue ().substring (0, 10);
                        String NicInstAsal = pertindakan_New_Raw.getRow (pertRow).getCell (24).getStringCellValue ();
                        Ganjil.getRow (row).createCell (1).setCellValue (JnsCrByr);
                        Ganjil.getRow (row).createCell (2).setCellValue (TglMsk);
                        Ganjil.getRow (row).createCell (3).setCellValue (NicInstAsal);
                        break;
                    }
                }
            }

//          cek per row. sesuaikan width nya
            for (int columnIndex = 0; columnIndex < Ganjil.getRow (0).getLastCellNum (); columnIndex++) {
                Ganjil.autoSizeColumn (columnIndex);
            }
            System.out.println ("01. " + BookPertindakanNew.getSheetAt (1).getSheetName () + " Complete");


//          buat sheet 2 Genap
            Sheet Genap = BookPertindakanNew.createSheet ();
            BookPertindakanNew.setSheetName (2, "Genap");
            System.out.println ("02. " + BookPertindakanNew.getSheetAt (2).getSheetName () + " Start");

//
            List<String> values = new ArrayList<> ();
            for (int row = 1; row <= pertindakan_New_Raw.getLastRowNum (); row++) {
                String cellValue = pertindakan_New_Raw.getRow (row).getCell (30).getStringCellValue ();
                String Tindakan = pertindakan_New_Raw.getRow (row).getCell (15).getStringCellValue ();
                if (!Tindakan.contains ("PAKET")) {
                    values.add (cellValue);
                }
            }

            Genap.createRow (0);
            for (int cell = 0; cell < pertindakan_New_Raw.getRow (0).getLastCellNum (); cell++) {
                Genap.getRow (0).createCell (cell).setCellValue (
                        pertindakan_New_Raw.getRow (0).getCell (cell).getStringCellValue ()
                );
            }

            List<String> sortedGenapValues = values.stream ().sorted ().toList ();
            IntStream.range (0, sortedGenapValues.size ())
                    .forEach (i -> {
                        String value = sortedGenapValues.get (i);
                        Genap.createRow (i + 1).createCell (30).setCellValue (value);
                    });

            for (int row = 1; row <= sortedGenapValues.size (); row++) {
                String cellValue = Genap.getRow (row).getCell (30).getStringCellValue ();
                for (int pertRow = 1; pertRow <= pertindakan_New_Raw.getLastRowNum (); pertRow++) {
                    String pertCellValue = pertindakan_New_Raw.getRow (pertRow).getCell (30).getStringCellValue ();
                    if (cellValue.equals (pertCellValue)) {
                        for (int cell = pertindakan_New_Raw.getRow (0).getLastCellNum (); cell >= 0; cell--) {
                            if (pertindakan_New_Raw.getRow (pertRow).getCell (cell) != null) {
                                if (pertindakan_New_Raw
                                        .getRow (pertRow)
                                        .getCell (cell).getCellType () == CellType.STRING) {
                                    Genap.getRow (row).createCell (cell).setCellValue (pertindakan_New_Raw
                                            .getRow (pertRow)
                                            .getCell (cell).getStringCellValue ());
                                } else {
                                    Genap.getRow (row).createCell (cell).setCellValue (pertindakan_New_Raw
                                            .getRow (pertRow)
                                            .getCell (cell).getNumericCellValue ());
                                }
                            }
                        }
                    }
                }
            }

            Genap.getRow (0).createCell (31).setCellValue ("TanggalReg");

//          CT Scan, USG , RONTGENT, Konsul Dokter Spesialis
            for (int row = 1; row <= sortedGenapValues.size (); row++) {
                String Tindakan = Genap.getRow (row).getCell (15).getStringCellValue ();
                if (Tindakan.contains ("CT Scan")) {
                    Genap.getRow (row).createCell (15).setCellValue ("CT Scan");
                } else if (Tindakan.contains ("USG")) {
                    Genap.getRow (row).createCell (15).setCellValue ("USG");
                } else if (Tindakan.contains ("Konsul Dokter Spesialis")) {
                    Genap.getRow (row).createCell (15).setCellValue ("Konsul Dokter Spesialis");
                } else {
                    Genap.getRow (row).createCell (15).setCellValue ("RONTGENT");
                }
                Genap.getRow (row).createCell (31).setCellValue (Genap.getRow (row).getCell (9)
                        .getStringCellValue ().substring (0,10));
            }


//          cek per row. sesuaikan width nya
            for (int columnIndex = 0; columnIndex < Genap.getRow (0).getLastCellNum (); columnIndex++) {
                Genap.autoSizeColumn (columnIndex);
            }
            System.out.println ("02. " + BookPertindakanNew.getSheetAt (2).getSheetName () + " Complete");

//          buat sheet 3 pertindakan
            Sheet Pertindakan = BookPertindakanNew.createSheet ();
            BookPertindakanNew.setSheetName (3, "1 Pertindakan");
            System.out.println ("03. " + BookPertindakanNew.getSheetAt (3).getSheetName () + " Start");

//          buat judul dan kasih kotak
            Pertindakan.createRow (5).createCell (0).setCellValue ("NO");
            Pertindakan.getRow (5).createCell (1).setCellValue ("Nama Tindakan");
            Pertindakan.getRow (5).createCell (2).setCellValue ("Jumlah");


            // Perform pivot simulation, and check if it not contains paket
            Map<String, Integer> pivotJumlahTindakan = new HashMap<> ();
            for (int i = 1; i <= pertindakanNewRawLastRowNum (pertindakan_New_Raw); i++) {
                Row row = pertindakan_New_Raw.getRow (i);
                String Tindakan = row.getCell (15).getStringCellValue ();
                if (!Tindakan.contains ("PAKET")) {
                    Integer count = pivotJumlahTindakan.getOrDefault (Tindakan, 0);
                    count++;
                    pivotJumlahTindakan.put (Tindakan, count);
                }
            }

//          Sort any value it contains
            List<Map.Entry<String, Integer>> entriesDoctor = new ArrayList<> (pivotJumlahTindakan.entrySet ());
            entriesDoctor.sort (Map.Entry.comparingByKey ());
            pivotJumlahTindakan = new LinkedHashMap<> ();
            for (Map.Entry<String, Integer> entry : entriesDoctor) {
                pivotJumlahTindakan.put (entry.getKey (), entry.getValue ());
            }

//          tulis hasil pivot ke pertindakan, mulai dari row 6
            int startRow = 6;
            int rowNum = startRow;
            for (Map.Entry<String, Integer> entry : pivotJumlahTindakan.entrySet ()) {
                Row row = Pertindakan.createRow (rowNum++);
                row.createCell (0).setCellValue (rowNum - 6);
                row.createCell (1).setCellValue (entry.getKey ());
                row.createCell (2).setCellValue (entry.getValue ());
            }

//          buat header center kemudian border semuanya
            for (int rightCell = 0; rightCell < Pertindakan.getRow (rowNum - 1).getLastCellNum (); rightCell++) {
                Pertindakan.getRow (startRow - 1).getCell (rightCell).setCellStyle (BorderCenterCellStyle);
                for (int downRow = startRow; downRow <= pertindakanNewRawLastRowNum (Pertindakan); downRow++) {
                    Pertindakan.getRow (downRow).getCell (rightCell).setCellStyle (AllBorderCellStyle);
                }
            }
//          cek per row. sesuaikan width nya
            int columnCountA2 = Pertindakan.getRow (startRow - 1).getLastCellNum ();
            for (int columnIndex = 0; columnIndex < columnCountA2; columnIndex++) {
                Pertindakan.autoSizeColumn (columnIndex);
            }
            System.out.println ("03. " + BookPertindakanNew.getSheetAt (3).getSheetName () + " Complete");


//        buat sheet 4 Jml tndakan per cr Byr pr hri
            BookPertindakanNew.createSheet ();
//            Sheet TndkanCrByrHr = BookPertindakanNew.getSheetAt (4);
            BookPertindakanNew.setSheetName (4, "2.Jml tndakan per cr Byr pr hri");
            System.out.println ("04. Sheet " + BookPertindakanNew.getSheetAt (4).getSheetName () + " Created");


//            List<String> masterCaraBayar = new ArrayList<>();
//            List<String> masterTanggal = new ArrayList<>();
//            List<String> masterTindakan = new ArrayList<>();
//
//            for (Row row = Genap.getRow(1); row != null; row = Genap.getRow(row.getRowNum()+1)) {
//                String caraBayar = row.getCell(8).getStringCellValue();
//                String tglMsk = row.getCell(9).getStringCellValue().substring(0, 10);
//                String tndk = row.getCell(15).getStringCellValue();
//
//                if (!masterCaraBayar.contains(caraBayar)) {
//                    masterCaraBayar.add(caraBayar);
//                }
//                if (!masterTanggal.contains(tglMsk)) {
//                    masterTanggal.add(tglMsk);
//                }
//                if (!masterTindakan.contains(tndk)) {
//                    masterTindakan.add(tndk);
//                }
//            }
//
//            masterTanggal.stream().sorted ();
//            masterCaraBayar.stream().sorted ();
//            masterTindakan.stream().sorted ();
//
//
//            TndkanCrByrHr.createRow (0).createCell(0);
//            TndkanCrByrHr.createRow (1).createCell(0);
//            TndkanCrByrHr.addMergedRegion (new CellRangeAddress (0,1,0,0));
//            TndkanCrByrHr.getRow (0).getCell (0).setCellValue("Tanggal");
//
//            int rowNumx = 2;
//            for (String tglMsk : masterTanggal) {
//                Row row = TndkanCrByrHr.createRow(rowNumx++);
//                row.createCell(0).setCellValue(tglMsk);
//            }
//
//            int cellTndkanCrByrHr = 1;
//            for (int i = 0; i < masterCaraBayar.size(); i++) {
//                for (int j = 0; j < masterTindakan.size(); j++) {
//                    String caraBayar = masterCaraBayar.get(i);
//                    String tndk = masterTindakan.get(j);
//                    int currentCellTndkanCrByrHr = cellTndkanCrByrHr + j + (i * masterTindakan.size());
//                    TndkanCrByrHr.getRow(0).createCell(currentCellTndkanCrByrHr).setCellValue(caraBayar);
//                    TndkanCrByrHr.getRow(1).createCell(currentCellTndkanCrByrHr).setCellValue(tndk);
//                }
//            }
//
//            //crbyr 8   rw 0 cl 1-28
//            //tgl   9   rw 2-31
//            //tndk  15  rw 1 cl 1-28
//
//            for (int row = 1; row <= Genap.getLastRowNum(); row++) {
//                String caraBayar = Genap.getRow(row).getCell(8).getStringCellValue();
//                String tglMsk = Genap.getRow(row).getCell(9).getStringCellValue().substring(0, 10);
//                String tndk = Genap.getRow(row).getCell(15).getStringCellValue();
//
//                if (caraBayar.equals(TndkanCrByrHr.getRow(0).getCell(1).getStringCellValue())
//                        && tglMsk.equals(TndkanCrByrHr.getRow(2).getCell(1).getStringCellValue())
//                        && tndk.equals(TndkanCrByrHr.getRow(1).getCell(1).getStringCellValue())) {
//                    int currentCount = (int) TndkanCrByrHr.getRow(2).getCell(1).getNumericCellValue();
//                    TndkanCrByrHr.getRow(2).createCell(1).setCellValue(currentCount + 1);
//                }
//            }

//        buat sheet 5 Jml tndakan per cr Byr pr hri
            BookPertindakanNew.createSheet ();
//            Sheet PsnCrByrHr = BookPertindakanNew.getSheetAt (5);
            BookPertindakanNew.setSheetName (5, "3.Pasien per cara bayar pr hari");
            System.out.println ("05. Sheet " + BookPertindakanNew.getSheetAt (5).getSheetName () + " Created");


//        buat sheet 6 Jml tndakan per cr Byr pr hri
            BookPertindakanNew.createSheet ();
            Sheet TndCrByr = BookPertindakanNew.getSheetAt (6);
            BookPertindakanNew.setSheetName (6, "4.Tindakan Percara bayar ");
            System.out.println ("06. " + BookPertindakanNew.getSheetAt (6).getSheetName () + " Start");


            Map<String, Integer> tndkCount = new TreeMap<> ();
            for (int row = 1; row <= Genap.getLastRowNum (); row++) {
                String tndk = Genap.getRow (row).getCell (8).getStringCellValue ();
                tndkCount.put (tndk, tndkCount.getOrDefault (tndk, 0) + 1);
            }

            TndCrByr.createRow (0).createCell (0).setCellValue ("Jenis Cara Bayar");
            TndCrByr.getRow (0).createCell (1).setCellValue ("Jumlah");

            int TndCrByrrow = 0;
            int TndCrByrSum = 0;
            for (Map.Entry<String, Integer> entry : tndkCount.entrySet ()) {
                TndCrByrrow++;
                TndCrByr.createRow (TndCrByrrow).createCell (0).setCellValue (entry.getKey ());
                TndCrByr.getRow (TndCrByrrow).createCell (1).setCellValue (entry.getValue ());
                TndCrByrSum += entry.getValue ();
            }

            int TndCrByrLastRow = TndCrByr.getLastRowNum () + 1;
            TndCrByr.createRow (TndCrByrLastRow).createCell (0).setCellValue ("Grand Total");
            TndCrByr.getRow (TndCrByrLastRow).createCell (1).setCellValue (TndCrByrSum);

//          buat header center kemudian border semuanya ps. use'<' because return 2 but there is 0, and 1. no number 2.
            for (int rightCell = 0; rightCell < TndCrByr.getRow (0).getLastCellNum (); rightCell++) {
                TndCrByr.getRow (0).getCell (rightCell).setCellStyle (BorderCenterCellStyle);
                for (int downRow = 1; downRow <= TndCrByr.getLastRowNum (); downRow++) {
                    TndCrByr.getRow (downRow).getCell (rightCell).setCellStyle (AllBorderCellStyle);
                }
            }
//          cek per row. sesuaikan width nya
            int columnCountTndCrByr = TndCrByr.getRow (0).getLastCellNum ();
            for (int columnIndex = 0; columnIndex < columnCountTndCrByr; columnIndex++) {
                TndCrByr.autoSizeColumn (columnIndex);
            }
            System.out.println ("06. " + BookPertindakanNew.getSheetAt (6).getSheetName () + " Completed");


//        buat sheet 7 Jml tndakan per cr Byr pr hri
            BookPertindakanNew.createSheet ();
            Sheet PsnCrByr = BookPertindakanNew.getSheetAt (7);
            BookPertindakanNew.setSheetName (7, "5.Pasien per cara bayar");
            System.out.println ("07. " + BookPertindakanNew.getSheetAt (7).getSheetName () + " Start");


            Map<String, Integer> PsncrByrCount = new TreeMap<> ();
            for (int row = 1; row <= Ganjil.getLastRowNum (); row++) {
                String crByr = Ganjil.getRow (row).getCell (1).getStringCellValue ();
                PsncrByrCount.put (crByr, PsncrByrCount.getOrDefault (crByr, 0) + 1);
            }

            PsnCrByr.createRow (0).createCell (0).setCellValue ("Jenis Cara Bayar");
            PsnCrByr.getRow (0).createCell (1).setCellValue ("Jumlah");

            int PsnCrByrrow = 0;
            int PsnCrByrSum = 0;
            for (Map.Entry<String, Integer> entry : PsncrByrCount.entrySet ()) {
                PsnCrByrrow++;
                PsnCrByr.createRow (PsnCrByrrow).createCell (0).setCellValue (entry.getKey ());
                PsnCrByr.getRow (PsnCrByrrow).createCell (1).setCellValue (entry.getValue ());
                PsnCrByrSum += entry.getValue ();
            }
            int PsnCrByrLastRow = PsnCrByr.getLastRowNum () + 1;
            PsnCrByr.createRow (PsnCrByrLastRow).createCell (0).setCellValue ("Grand Total");
            PsnCrByr.getRow (PsnCrByrLastRow).createCell (1).setCellValue (PsnCrByrSum);

//          buat header center kemudian border semuanya ps. use'<' because return 2 but there is 0, and 1. no number 2.
            for (int rightCell = 0; rightCell < PsnCrByr.getRow (0).getLastCellNum (); rightCell++) {
                PsnCrByr.getRow (0).getCell (rightCell).setCellStyle (BorderCenterCellStyle);
                for (int downRow = 1; downRow <= PsnCrByr.getLastRowNum (); downRow++) {
                    PsnCrByr.getRow (downRow).getCell (rightCell).setCellStyle (AllBorderCellStyle);
                }
            }
//          cek per row. sesuaikan width nya
            int columnCountPsnCrByr = PsnCrByr.getRow (0).getLastCellNum ();
            for (int columnIndex = 0; columnIndex < columnCountPsnCrByr; columnIndex++) {
                PsnCrByr.autoSizeColumn (columnIndex);
            }
            System.out.println ("07. " + BookPertindakanNew.getSheetAt (7).getSheetName () + " Completed");

//        buat sheet 8 Jml tndakan per cr Byr pr hri
            BookPertindakanNew.createSheet ();
            Sheet TndInstAsal = BookPertindakanNew.getSheetAt (8);
            BookPertindakanNew.setSheetName (8, "6.Tindakan Per instalasi asal");
            System.out.println ("08. " + BookPertindakanNew.getSheetAt (8).getSheetName () + " Start");
            //count inst asal
            Map<String, Integer> TndInstAsalCount = new TreeMap<> ();
            for (int row = 1; row <= Genap.getLastRowNum (); row++) {
                String instAsal = Genap.getRow (row).getCell (24).getStringCellValue ();
                TndInstAsalCount.put (instAsal, TndInstAsalCount.getOrDefault (instAsal, 0) + 1);
            }

            TndInstAsal.createRow(0).createCell(0).setCellValue("Instalasi Asal");
            TndInstAsal.getRow(0).createCell(1).setCellValue("Jumlah");

            int TndInstAsalrow = 0;
            int TndInstAsalSum = 0;
            for (Map.Entry<String, Integer> entry : TndInstAsalCount.entrySet()) {
                TndInstAsalrow++;
                TndInstAsal.createRow (TndInstAsalrow).createCell (0).setCellValue (entry.getKey ());
                TndInstAsal.getRow (TndInstAsalrow).createCell (1).setCellValue (entry.getValue ());
                TndInstAsalSum += entry.getValue ();

                int TndInstAsalLastRow = TndInstAsal.getLastRowNum () + 1;
                TndInstAsal.createRow (TndInstAsalLastRow).createCell (0).setCellValue ("Grand Total");
                TndInstAsal.getRow (TndInstAsalLastRow).createCell (1).setCellValue (TndInstAsalSum);
            }

//          buat header center, adjust width kemudian border ps. use'<' because return 2 but there is 0, and 1. no number 2.
                for (int rightCell = 0; rightCell<TndInstAsal.getRow (0).getLastCellNum ();rightCell++){
                    TndInstAsal.getRow (0).getCell (rightCell).setCellStyle(BorderCenterCellStyle);
                    TndInstAsal.autoSizeColumn(rightCell);
                    for (int downRow = 1; downRow<= TndInstAsal.getLastRowNum (); downRow++){
                        TndInstAsal.getRow (downRow).getCell (rightCell).setCellStyle(AllBorderCellStyle);
                    }
                }

                System.out.println ("08. "+ BookPertindakanNew.getSheetAt (8).getSheetName ()+" Completed");

//        buat sheet 9 Jml tndakan per cr Byr pr hri
            BookPertindakanNew.createSheet ();
            Sheet PsnInstAsal = BookPertindakanNew.getSheetAt (9);
            BookPertindakanNew.setSheetName (9, "7.Pasien per instalasi asal");
            System.out.println ("09. " + BookPertindakanNew.getSheetAt (9).getSheetName () + " Start");

            Map<String, Integer> PsnInstAsalCount = new TreeMap<> ();
            for (int row = 1; row <= Ganjil.getLastRowNum (); row++) {
                String psnInstAsal = Ganjil.getRow (row).getCell (3).getStringCellValue ();
                PsnInstAsalCount.put (psnInstAsal, PsnInstAsalCount.getOrDefault (psnInstAsal, 0) + 1);
            }

            PsnInstAsal.createRow (0).createCell (0).setCellValue ("Jenis Cara Bayar");
            PsnInstAsal.getRow (0).createCell (1).setCellValue ("Jumlah");

            int PsnInstAsalrow = 0;
            int PsnInstAsalSum = 0;
            for (Map.Entry<String, Integer> entry : PsnInstAsalCount.entrySet ()) {
                PsnInstAsalrow++;
                PsnInstAsal.createRow (PsnInstAsalrow).createCell (0).setCellValue (entry.getKey ());
                PsnInstAsal.getRow (PsnInstAsalrow).createCell (1).setCellValue (entry.getValue ());
                PsnInstAsalSum += entry.getValue ();
            }
            int PsnInstAsalLastRow = PsnInstAsal.getLastRowNum () + 1;
            PsnInstAsal.createRow (PsnInstAsalLastRow).createCell (0).setCellValue ("Grand Total");
            PsnInstAsal.getRow (PsnInstAsalLastRow).createCell (1).setCellValue (PsnInstAsalSum);

//          buat header center kemudian border semuanya ps. use'<' because return 2 but there is 0, and 1. no number 2.
            for (int rightCell = 0; rightCell < PsnInstAsal.getRow (0).getLastCellNum (); rightCell++) {
                PsnInstAsal.getRow (0).getCell (rightCell).setCellStyle (BorderCenterCellStyle);
                PsnInstAsal.autoSizeColumn (rightCell);
                for (int downRow = 1; downRow <= PsnInstAsal.getLastRowNum (); downRow++) {
                    PsnInstAsal.getRow (downRow).getCell (rightCell).setCellStyle (AllBorderCellStyle);
                }
            }

            System.out.println ("09. "+ BookPertindakanNew.getSheetAt (9).getSheetName ()+" Completed");

//        buat sheet 10 Jml tndakan per cr Byr pr hri
            BookPertindakanNew.createSheet ();
//            Sheet PsnCrByrHr = BookPertindakanNew.getSheetAt (10);
            BookPertindakanNew.setSheetName (10, "8.Jumlah pendapatan");
            System.out.println ("10. Sheet " + BookPertindakanNew.getSheetAt (10).getSheetName () + " Created");

//        buat sheet 11 Jml tndakan per cr Byr pr hri
            InputStream responTime = new FileInputStream("C:\\sat work\\test\\"+fileNameMonitoringf1+".xlsx");
            Workbook bookRespontime = new XSSFWorkbook(responTime);
            BookPertindakanNew.createSheet ("9. Monitoring Hasil Rad");
            Sheet ResponTime = BookPertindakanNew.getSheetAt (11);
            Sheet ResponTimeRaw = bookRespontime.getSheetAt (0);
//            for (int row =0; row<=ResponTimeRaw.getLastRowNum ();row++){
//                ResponTime.createRow (row);
//                for (int cell=0;cell<ResponTimeRaw.getRow (0).getLastCellNum ();cell++){
////                    ResponTime.getRow (row).createCell (cell).setCellValue (ResponTimeRaw.getRow (row).getCell (cell).getStringCellValue ());
//                    Cell responTimeCell = ResponTimeRaw.getRow(row).getCell(cell);
//                    if (responTimeCell.getCellType() == CellType.STRING) {
//                        ResponTime.getRow(row).createCell(cell).setCellValue(responTimeCell.getStringCellValue());
//                    } else if (responTimeCell.getCellType() == CellType.NUMERIC) {
//                        ResponTime.getRow(row).createCell(cell).setCellValue(responTimeCell.getNumericCellValue());
//                    }
//                }
//            }
            for (int row = 0; row <= ResponTimeRaw.getLastRowNum() && ResponTimeRaw.getRow(row) != null; row++) {
                ResponTime.createRow(row);
                for (int cell = 0; cell < ResponTimeRaw.getRow(0).getLastCellNum(); cell++) {
                    Cell responTimeCell = ResponTimeRaw.getRow(row).getCell(cell);
                    if (responTimeCell == null) {
                        ResponTime.getRow(row).createCell(cell).setCellValue("");
                    } else if (responTimeCell.getCellType() == CellType.STRING) {
                        ResponTime.getRow(row).createCell(cell).setCellValue(responTimeCell.getStringCellValue());
                    } else if (responTimeCell.getCellType() == CellType.NUMERIC) {
                        ResponTime.getRow(row).createCell(cell).setCellValue(responTimeCell.getNumericCellValue());
                    }
                }
            }
            System.out.println ("11. "+ BookPertindakanNew.getSheetAt (11).getSheetName ()+" Completed");






















        } catch (Exception e) {
            e.printStackTrace();
        }


        try {
            outputStream = new FileOutputStream(fileNamePertindakanNew +".xlsx");
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

}
