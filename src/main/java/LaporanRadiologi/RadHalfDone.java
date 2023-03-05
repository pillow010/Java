package LaporanRadiologi;

import StylingLaporan.StylerRepo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.*;


public class RadHalfDone extends StylerRepo{
    public static void main(String[] args) {
        new RadHalfDone ();

    }
    private Workbook BookPertindakanNew;

    private FileOutputStream outputStream;

    String fileNamePertindakanNew = "23 02 rad pertindakan new";
    String fileNameMonitoringf1 = "23 02 rad monitorring f1";
    String fileNamePelayananPenunjang = "23 02 rad pelayanan penunjang";
    String fileNameMonitoringf2 = "23 02 rad monitorring f2";


    public RadHalfDone(){
        try {
            InputStream pertindakanNew = new FileInputStream ("C:\\sat work\\test\\" + fileNamePertindakanNew + ".xlsx");
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

            // hashmap sub-inst
            HashMap<String, String> subInstMap = new HashMap<> ();
            subInstMap.put ("HD01", "HD");
            subInstMap.put ("RHM01", "RHM");
            subInstMap.put ("IGD01", "Umum");
            subInstMap.put ("IGD02", "Ponek");
            subInstMap.put ("IRNA01", "Teratai 1");
            subInstMap.put ("IRNA02", "Teratai 2");
            subInstMap.put ("IRNA03", "Matahari");
            subInstMap.put ("IRNA04", "Tulip");
            subInstMap.put ("IRNA05", "Anyelir");
            subInstMap.put ("IRNA06", "ICU");
            subInstMap.put ("IRNA07", "IGD (Mawar)");
            subInstMap.put ("IRNA08", "Perinatologi");
            subInstMap.put ("IRNA09", "NICU");
            subInstMap.put ("IRNA10", "VK (Anggrek)");
            subInstMap.put ("IRNA11", "IBS (Sentral)");
            subInstMap.put ("IRNA12", "IBS (IGD)");
            subInstMap.put ("IRNA13", "ISOLASI");
            subInstMap.put ("IRNA14", "TERATAI");
            subInstMap.put ("IRNA15", "ALAMANDA");
            subInstMap.put ("IRNA16", "LILY");
            subInstMap.put ("IRNA17", "CATTLEYA MAGNOLIA");
            subInstMap.put ("IRNA18", "SAKURA");
            subInstMap.put ("IRNA19", "HCU");
            subInstMap.put ("IRNA20", "PICU");
            subInstMap.put ("IRNA21", "ALAMANDA 2");
            subInstMap.put ("IRNA22", "ALAMANDA 3");
            subInstMap.put ("IRNA23", "KEMBANG LILY");
            subInstMap.put ("IRNA24", "LILY 2");
            subInstMap.put ("MCU01", "MCU");
            subInstMap.put ("IRJ01", "Umum");
            subInstMap.put ("IRJ02", "Kebidanan dan Kandungan");
            subInstMap.put ("IRJ03", "Gigi Umum");
            subInstMap.put ("IRJ04", "Gigi Anak");
            subInstMap.put ("IRJ05", "Bedah Umum");
            subInstMap.put ("IRJ06", "Bedah Digestif");
            subInstMap.put ("IRJ07", "Penyakit Dalam");
            subInstMap.put ("IRJ08", "THT");
            subInstMap.put ("IRJ09", "Konservasi Gigi");
            subInstMap.put ("IRJ10", "Periodontik");
            subInstMap.put ("IRJ11", "Mata");
            subInstMap.put ("IRJ12", "Akupuntur");
            subInstMap.put ("IRJ13", "Bedah Urologi");
            subInstMap.put ("IRJ14", "Bedah Orthopedi");
            subInstMap.put ("IRJ15", "Klinik Sahabat");
            subInstMap.put ("IRJ16", "Anak");
            subInstMap.put ("IRJ17", "Paru");
            subInstMap.put ("IRJ18", "DOTS");
            subInstMap.put ("IRJ19", "Anestesi");
            subInstMap.put ("IRJ20", "Saraf");
            subInstMap.put ("IRJ21", "Psikiatri");
            subInstMap.put ("IRJ22", "Kulit dan Kelamin");
            subInstMap.put ("IRJ23", "Tumbuh Kembang Anak");
            subInstMap.put ("IRJ24", "Geriatri");
            subInstMap.put ("IRJ25", "KIA -KB");
            subInstMap.put ("IRJ26", "Gizi");
            subInstMap.put ("IRJ27", "Bedah Vaskuler");
            subInstMap.put ("IRJ28", "Jantung");
            subInstMap.put ("IRJ29", "Ispa");
            subInstMap.put ("IRJ30", "NEUROLOGI ANAK");
            subInstMap.put ("IRJ31", "BEDAH ONKOLOGI");


            pertindakan_New_Raw.getRow (0).createCell (28).setCellValue ("SUB INST");
            pertindakan_New_Raw.getRow (0).createCell (29).setCellValue ("NOREG");
            pertindakan_New_Raw.getRow (0).createCell (30).setCellValue ("NOREGTINDAKAN");
            for (int _row = 1; _row <= pertindakanNewRawLastRowNum (pertindakan_New_Raw); _row++) {
                Row row = pertindakan_New_Raw.getRow (_row);
                Cell cell = row.getCell (24);
                //ifnull rujukan luar rs
                if (cell == null) {
                    row.createCell (28).setCellValue ("RUJUKAN LUAR RS");
                    row.createCell (24).setCellValue ("RUJUKAN LUAR RS");
                } else {
                    //not null combile inst and sub inst then get the value
                    String value = cell.getStringCellValue () + row.getCell (27).getStringCellValue ();
                    //is there new sub inst?
                    row.createCell (28).setCellValue (subInstMap.getOrDefault (value, "Not Found"));
                }

                pertindakan_New_Raw.getRow (_row).createCell (29).setCellValue (
                        pertindakan_New_Raw.getRow (_row).getCell (0).getStringCellValue () +
                                pertindakan_New_Raw.getRow (_row).getCell (1).getStringCellValue () +
                                pertindakan_New_Raw.getRow (_row).getCell (2).getStringCellValue () +
                                pertindakan_New_Raw.getRow (_row).getCell (3).getStringCellValue () +
                                pertindakan_New_Raw.getRow (_row).getCell (4).getStringCellValue ()
                );
                pertindakan_New_Raw.getRow (_row).createCell (30).setCellValue (
                        pertindakan_New_Raw.getRow (_row).getCell (29).getStringCellValue ()
                                + pertindakan_New_Raw.getRow (_row).getCell (15).getStringCellValue ()
                );
            }

            System.out.println ("00. " + BookPertindakanNew.getSheetAt (0).getSheetName () + " Complete");


//          buat sheet 1 Ganjil
            Sheet Ganjil = BookPertindakanNew.createSheet ();
            BookPertindakanNew.setSheetName (1, "Ganjil");
            System.out.println ("01. " + BookPertindakanNew.getSheetAt (1).getSheetName () + " Start");

            Ganjil.createRow (0).createCell (0).setCellValue ("NOREG");
            Ganjil.getRow (0).createCell (1).setCellValue ("JENIS CARA BAYAR");
            Ganjil.getRow (0).createCell (2).setCellValue ("TANGGAL MASUK");
            Ganjil.getRow (0).createCell (3).setCellValue ("NIC INST ASAL");

            Set<String> uniqueValues = new HashSet<> ();
            int actualRow = 1;
            for (int row = 1; row <= pertindakan_New_Raw.getLastRowNum (); row++) {
                if (pertindakan_New_Raw.getRow (row) != null) { // check if row is not empty
                    Cell cell = pertindakan_New_Raw.getRow (row).getCell (29);
                    String cellValue = cell.getStringCellValue ();
                    if (row > 0 && !cellValue.isBlank () && !uniqueValues.contains (cellValue)) { // check if cell is not empty
                        Ganjil.createRow (actualRow);
                        String Noreg = pertindakan_New_Raw.getRow (row).getCell (29).getStringCellValue ();
                        String JnsCrByr = pertindakan_New_Raw.getRow (row).getCell (8).getStringCellValue ();
                        String TglMsk = pertindakan_New_Raw.getRow (row).getCell (9).getStringCellValue ().substring (0, 10);
                        String NicInstAsal = pertindakan_New_Raw.getRow (row).getCell (24).getStringCellValue ();
                        Ganjil.getRow (actualRow).createCell (0).setCellValue (Noreg);
                        Ganjil.getRow (actualRow).createCell (1).setCellValue (JnsCrByr);
                        Ganjil.getRow (actualRow).createCell (2).setCellValue (TglMsk);
                        Ganjil.getRow (actualRow).createCell (3).setCellValue (NicInstAsal);
                        uniqueValues.add (cellValue);
                        actualRow++;
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

            Genap.createRow (0);
            for (int cell = 0; cell < pertindakan_New_Raw.getRow (0).getLastCellNum (); cell++) { // for each cell
                Genap.getRow (0).createCell (cell)
                        .setCellValue (pertindakan_New_Raw.getRow (0).getCell (cell).getStringCellValue ());
            }
            Genap.getRow (0).createCell (31).setCellValue ("TanggalReg");
            Genap.getRow (0).createCell (32).setCellValue ("KelompokTindakan");

            int actualGenapRow = 1;
            for (int row = 1; row <= pertindakan_New_Raw.getLastRowNum (); row++) { // for each row
                String Tindakan = pertindakan_New_Raw.getRow (row).getCell (15).getStringCellValue (); // string of tindakan
                if (!Tindakan.contains ("PAKET")) {
                    Genap.createRow (actualGenapRow); // create one row per non-"PAKET" row
                    for (int cell = 0; cell < pertindakan_New_Raw.getRow (0).getLastCellNum (); cell++) { // for each cell
                        Cell pertindakanValue = pertindakan_New_Raw.getRow (row).getCell (cell); // type of value
                        if (pertindakanValue == null) {
                            Genap.getRow (actualGenapRow).createCell (cell).setCellValue ("");
                        } else if (pertindakanValue.getCellType () == CellType.STRING) {
                            Genap.getRow (actualGenapRow).createCell (cell)
                                    .setCellValue (pertindakanValue.getStringCellValue ());
                        } else if (pertindakanValue.getCellType () == CellType.NUMERIC) {
                            Genap.getRow (actualGenapRow).createCell (cell)
                                    .setCellValue (pertindakanValue.getNumericCellValue ());
                        }
                    }
                    Genap.getRow (actualGenapRow).createCell (31).setCellValue (Genap.getRow (actualGenapRow).getCell (9)
                            .getStringCellValue ().substring (0, 10));

                    if (Tindakan.contains ("CT Scan")) {
                        Genap.getRow (actualGenapRow).createCell (32).setCellValue ("CT Scan");
                    } else if (Tindakan.contains ("USG")) {
                        Genap.getRow (actualGenapRow).createCell (32).setCellValue ("USG");
                    } else if (Tindakan.contains ("Konsul Dokter Spesialis")) {
                        Genap.getRow (actualGenapRow).createCell (32).setCellValue ("Konsul Dokter Spesialis");
                    } else {
                        Genap.getRow (actualGenapRow).createCell (32).setCellValue ("RONTGENT");
                    }

                    actualGenapRow++; // increment actualGenapRow only once per row in pertindakan_New_Raw
                }
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

            TndInstAsal.createRow (0).createCell (0).setCellValue ("Instalasi Asal");
            TndInstAsal.getRow (0).createCell (1).setCellValue ("Jumlah");

            int TndInstAsalrow = 0;
            int TndInstAsalSum = 0;
            for (Map.Entry<String, Integer> entry : TndInstAsalCount.entrySet ()) {
                TndInstAsalrow++;
                TndInstAsal.createRow (TndInstAsalrow).createCell (0).setCellValue (entry.getKey ());
                TndInstAsal.getRow (TndInstAsalrow).createCell (1).setCellValue (entry.getValue ());
                TndInstAsalSum += entry.getValue ();

                int TndInstAsalLastRow = TndInstAsal.getLastRowNum () + 1;
                TndInstAsal.createRow (TndInstAsalLastRow).createCell (0).setCellValue ("Grand Total");
                TndInstAsal.getRow (TndInstAsalLastRow).createCell (1).setCellValue (TndInstAsalSum);
            }

//          buat header center, adjust width kemudian border ps. use'<' because return 2 but there is 0, and 1. no number 2.
            for (int rightCell = 0; rightCell < TndInstAsal.getRow (0).getLastCellNum (); rightCell++) {
                TndInstAsal.getRow (0).getCell (rightCell).setCellStyle (BorderCenterCellStyle);
                TndInstAsal.autoSizeColumn (rightCell);
                for (int downRow = 1; downRow <= TndInstAsal.getLastRowNum (); downRow++) {
                    TndInstAsal.getRow (downRow).getCell (rightCell).setCellStyle (AllBorderCellStyle);
                }
            }

            System.out.println ("08. " + BookPertindakanNew.getSheetAt (8).getSheetName () + " Completed");

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

            System.out.println ("09. " + BookPertindakanNew.getSheetAt (9).getSheetName () + " Completed");

//        buat sheet 10 Jml tndakan per cr Byr pr hri
            BookPertindakanNew.createSheet ();
//            Sheet PsnCrByrHr = BookPertindakanNew.getSheetAt (10);
            BookPertindakanNew.setSheetName (10, "8.Jumlah pendapatan");
            System.out.println ("10. Sheet " + BookPertindakanNew.getSheetAt (10).getSheetName () + " Created");

//        buat sheet 11 Jml tndakan per cr Byr pr hri
            InputStream monitoringF1 = new FileInputStream ("C:\\sat work\\test\\" + fileNameMonitoringf1 + ".xlsx");
            InputStream monitoringF2 = new FileInputStream ("C:\\sat work\\test\\" + fileNameMonitoringf2 + ".xlsx");
            InputStream _pelayananPenunjang = new FileInputStream ("C:\\sat work\\test\\" + fileNamePelayananPenunjang + ".xlsx");
            Workbook bookRespontimeF1 = new XSSFWorkbook (monitoringF1);
            Workbook bookRespontimeF2 = new XSSFWorkbook (monitoringF2);
            Workbook bookPelayananPenunjang = new XSSFWorkbook (_pelayananPenunjang);
            BookPertindakanNew.createSheet ("9. Monitoring Hasil Rad");
            Sheet monitoringHasilRad = BookPertindakanNew.getSheetAt (11);
            Sheet monitoringHasilRadF1 = bookRespontimeF1.getSheetAt (0);
            Sheet monitoringHasilRadF2 = bookRespontimeF2.getSheetAt (0);
            Sheet pelayananPenunjang = bookPelayananPenunjang.getSheetAt (0);
            System.out.println ("11. " + BookPertindakanNew.getSheetAt (11).getSheetName () + " Start");
            monitoringHasilRad.createRow (0);


            //ambil dari monitoring hasil format satu, letakkan pada sheet 11
            actualGenapRow = 0;
            for (int row = 0; row <= monitoringHasilRadF1.getLastRowNum () && monitoringHasilRadF1.getRow (row) != null; row++) {
                if (!monitoringHasilRadF1.getRow (row).getCell (3).getStringCellValue ().contains ("PAKET")) {
                    Row newMonitoringRow = monitoringHasilRad.createRow (actualGenapRow);
                    for (int cell = 0; cell < monitoringHasilRadF1.getRow (0).getLastCellNum (); cell++) {
                        Cell responTimeCell = monitoringHasilRadF1.getRow (row).getCell (cell);
                        if (responTimeCell == null) {
                            newMonitoringRow.createCell (cell + 1).setCellValue ("");
                        } else if (responTimeCell.getCellType () == CellType.STRING) {
                            newMonitoringRow.createCell (cell + 1).setCellValue (responTimeCell.getStringCellValue ());
                        } else if (responTimeCell.getCellType () == CellType.NUMERIC) {
                            newMonitoringRow.createCell (cell + 1).setCellValue (responTimeCell.getNumericCellValue ());
                        }
                    }
                    String Keterangan = newMonitoringRow.getCell (4).getStringCellValue ();
                    if (Keterangan.equals ("Konsul Dokter Spesialis") || Keterangan.equals ("C-Arm") || Keterangan.equals ("Foto gigi/dental CR") || Keterangan.equals ("Foto gigi/dental") || Keterangan.equals ("Panoramic") || Keterangan.equals ("")) {
                        newMonitoringRow.createCell (10).setCellValue (Keterangan);
                    }
                    if (newMonitoringRow.getCell (10).getStringCellValue ().equals ("")) {
                        newMonitoringRow.createCell (10).setCellValue ("Belum Dibaca");
                    }
                    actualGenapRow++; // increment actualGenapRow by 1 after creating a new row
                }
            }


            // Create a HashMap to store monitoringValue/SID and row its belong
            HashMap<String, List<Integer>> sampleIDHashMap = new HashMap<String, List<Integer>> ();
            for (int monitoringRow = 1; monitoringRow <= monitoringHasilRad.getLastRowNum (); monitoringRow++) {
                String sampleID = monitoringHasilRad.getRow (monitoringRow).getCell (6).getStringCellValue ();
                if (!sampleIDHashMap.containsKey (sampleID)) {
                    sampleIDHashMap.put (sampleID, new ArrayList<Integer> ());
                }
                sampleIDHashMap.get (sampleID).add (monitoringRow);
            }

            monitoringHasilRad.getRow (0).createCell (0).setCellValue ("NOREG");
            // Loop through pelayananPenunjang sheet and update monitoringHasilRad sheet using HashMap
            for (int row = 1; row <= pelayananPenunjang.getLastRowNum (); row++) {
                String pelayananPenunjangValue = pelayananPenunjang.getRow (row).getCell (29).getStringCellValue ();
                List<Integer> monitoringRows = sampleIDHashMap.get (pelayananPenunjangValue);
                if (monitoringRows != null) {
                    for (Integer monitoringRow : monitoringRows) {
                        monitoringHasilRad.getRow (monitoringRow).createCell (0).setCellValue (
                                pelayananPenunjang.getRow (row).getCell (0).getStringCellValue () +
                                        pelayananPenunjang.getRow (row).getCell (1).getStringCellValue () +
                                        pelayananPenunjang.getRow (row).getCell (2).getStringCellValue () +
                                        pelayananPenunjang.getRow (row).getCell (3).getStringCellValue () +
                                        pelayananPenunjang.getRow (row).getCell (4).getStringCellValue ());
                    }
                }
            }

            HashMap<String, List<Integer>> noRegHashMap = new HashMap<String, List<Integer>> ();
            for (int monitoringRow = 1; monitoringRow <= monitoringHasilRad.getLastRowNum (); monitoringRow++) {
                String noreg = monitoringHasilRad.getRow (monitoringRow).getCell (0).getStringCellValue ();
                if (!noRegHashMap.containsKey (noreg)) {
                    noRegHashMap.put (noreg, new ArrayList<> ());
                }
                noRegHashMap.get (noreg).add (monitoringRow);
            }

            monitoringHasilRad.getRow (0).createCell (11).setCellValue ("Nick Inst");
            monitoringHasilRad.getRow (0).createCell (12).setCellValue ("Sub Inst");
            for (int row = 1; row <= pertindakan_New_Raw.getLastRowNum (); row++) {
                String noreg = pertindakan_New_Raw.getRow (row).getCell (29).getStringCellValue ();
                List<Integer> monitoringRows = noRegHashMap.get (noreg);
                if (monitoringRows != null) {
                    for (Integer monitoringRow : monitoringRows) {
                        monitoringHasilRad.getRow (monitoringRow).createCell (11).setCellValue (
                                pertindakan_New_Raw.getRow (row).getCell (24).getStringCellValue ()
                        );
                        monitoringHasilRad.getRow (monitoringRow).createCell (12).setCellValue (
                                pertindakan_New_Raw.getRow (row).getCell (28).getStringCellValue ()
                        );
                    }
                }
            }

            monitoringHasilRad.getRow (0).createCell (13).setCellValue ("PETUGAS MENYERAHKAN");
            monitoringHasilRad.getRow (0).createCell (14).setCellValue ("PENGAMBIL");
//            monitoringHasilRad.getRow (0).createCell (15).setCellValue ("RESPON TIME RS");
//            monitoringHasilRad.getRow (0).createCell (16).setCellValue ("RESPON TIME RAD");
//            monitoringHasilRad.getRow (0).createCell (17).setCellValue ("TOTAL RS");
//            monitoringHasilRad.getRow (0).createCell (18).setCellValue ("TOTAL RAD");
//            monitoringHasilRad.getRow (0).createCell (19).setCellValue ("RATA2 RS");
//            monitoringHasilRad.getRow (0).createCell (20).setCellValue ("RATA2 RAD");
//            monitoringHasilRad.getRow (0).createCell (21).setCellValue ("CITO");

            HashMap<String, List<Integer>> accIDHashMap = new HashMap<> ();
            for (int monitoringRow =1;monitoringRow<=monitoringHasilRad.getLastRowNum ();monitoringRow++){
                String accId = String.valueOf (monitoringHasilRad.getRow (monitoringRow).getCell (7).getNumericCellValue ()).substring (0,5);
                if (!accIDHashMap.containsKey (accId)){
                    accIDHashMap.put (accId,new ArrayList<> ());
                }
                accIDHashMap.get (accId).add (monitoringRow);
            }

            for (int row=1;row<=monitoringHasilRadF2.getLastRowNum ();row++){
                String accId = monitoringHasilRadF2.getRow (row).getCell (7).getStringCellValue ();
                List<Integer> monitoringRows = accIDHashMap.get (accId);
                Cell menyerahkan = monitoringHasilRadF2.getRow (row).getCell (20);
                Cell menerima = monitoringHasilRadF2.getRow (row).getCell (21);
                if (monitoringRows!=null){
                    for (Integer monitoringRow:monitoringRows) {
                        if (menyerahkan != null) {
                            monitoringHasilRad.getRow (monitoringRow).createCell (13).setCellValue (
                                    menerima.getStringCellValue ()
                            );
                            monitoringHasilRad.getRow (monitoringRow).createCell (14).setCellValue (
                                    menyerahkan.getStringCellValue ()
                            );
                        } else {
                            monitoringHasilRad.getRow (monitoringRow).createCell (13).setCellValue ("-");
                            monitoringHasilRad.getRow (monitoringRow).createCell (14).setCellValue ("-");
                        }
                    }
                }
            }





            //buat header center kemudian border semuanya ps. use'<' because return 2 but there is 0, and 1. no number 2.
            for (int rightCell = 0; rightCell < monitoringHasilRad.getRow (0).getLastCellNum (); rightCell++) {
                monitoringHasilRad.getRow (0).getCell (rightCell).setCellStyle (BorderCenterCellStyle);
                monitoringHasilRad.autoSizeColumn (rightCell);
                for (int downRow = 1; downRow <= monitoringHasilRad.getLastRowNum (); downRow++) {
                    System.out.println (rightCell);
                    monitoringHasilRad.getRow (downRow).getCell (rightCell).setCellStyle (AllBorderCellStyle);
                }
            }




            monitoringHasilRad.getRow (0).createCell (0).setCellValue ("NOREG");
            System.out.println ("11. "+ BookPertindakanNew.getSheetAt (11).getSheetName ()+" Completed");






















        } catch (Exception e) {
            e.printStackTrace();
        }


        try {
            outputStream = new FileOutputStream(fileNamePertindakanNew +" half done.xlsx");
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
