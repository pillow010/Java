package LaporanRadiologi;

import StylingLaporan.StylerRepo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;


public class RadHalfDone extends StylerRepo{
    public static void main(String[] args) {
        new RadHalfDone ();

    }
    boolean doneFinal = true;
    private Workbook BookPertindakanNew;
    private FileOutputStream outputStream;
    String localDate = LocalDate.now().minusMonths (1).format (DateTimeFormatter.ofPattern ("yy MM"));
    LocalDateTime now = LocalDateTime.now();
    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd HHmmss");
    String formattedDateTime = now.format(formatter);
    String fileInput                  = "C:\\sat work\\test\\1. input\\";
    String fileOutput                 = "C:\\sat work\\test\\2. output\\";
    String fileNamePelayananPenunjang = localDate + " rad pelayanan penunjang";
    String fileNamePertindakanNew     = localDate + " rad tindakan new";
    String fileNameMonitoringf1       = localDate + " rad monitoring f1";
    String fileNameMonitoringf2       = localDate + " rad monitoring f2";
    String fileNameResponTime         = localDate + " rad respon time";


    public RadHalfDone(){
        try {
            InputStream pertindakanNew = new FileInputStream (fileInput + fileNamePertindakanNew + ".xlsx");
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
            subInstMap.put ("IRNA25","BOUGENVILLE 2");
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
            subInstMap.put ("IRJ32", "CAPD");


            pertindakan_New_Raw.getRow (0).createCell (28).setCellValue ("SUB INST");
            pertindakan_New_Raw.getRow (0).createCell (29).setCellValue ("NOREG");
            pertindakan_New_Raw.getRow (0).createCell (30).setCellValue ("NOREGTINDAKAN");
            for (int _row = 1; _row <= pertindakanNewRawLastRowNum (pertindakan_New_Raw); _row++) {
                Row row = pertindakan_New_Raw.getRow (_row);
                Cell cell = row.getCell (24);
                //if null rujukan luar rs
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


//          buat sheet 1 Ganjil untuk laporan per pasien
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


//          buat sheet 2 Genap untuk laporan per tindakan
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
            Pertindakan.createRow (0).createCell (0).setCellValue ("NO");
            Pertindakan.getRow (0).createCell (1).setCellValue ("Nama Tindakan");
            Pertindakan.getRow (0).createCell (2).setCellValue ("Jumlah");

            // Perform pivot simulation, and check if it not contains paket
            Map<String, Integer> pivotJumlahTindakan = new HashMap<>();
            for (int i = 1; i <= pertindakanNewRawLastRowNum(pertindakan_New_Raw); i++) {
                Row row = pertindakan_New_Raw.getRow(i);
                String Tindakan = row.getCell(15).getStringCellValue();
                if (!Tindakan.contains("PAKET")) {
                    Integer count = pivotJumlahTindakan.getOrDefault(Tindakan, 0);
                    count++;
                    pivotJumlahTindakan.put(Tindakan, count);
                }
            }

            // Sort any value it contains
            List<Map.Entry<String, Integer>> entriesDoctor = new ArrayList<>(pivotJumlahTindakan.entrySet());
            entriesDoctor.sort(Map.Entry.comparingByKey());
            pivotJumlahTindakan = new LinkedHashMap<>();
            for (Map.Entry<String, Integer> entry : entriesDoctor) {
                pivotJumlahTindakan.put(entry.getKey(), entry.getValue());
            }

            // Write pivot result to Pertindakan sheet, starting from row 6
            int rowNum = 1;
            for (Map.Entry<String, Integer> entry : pivotJumlahTindakan.entrySet()) {
                Row row = Pertindakan.createRow(rowNum++);
                row.createCell(0).setCellValue(rowNum - 1);
                row.createCell(1).setCellValue(entry.getKey());
                row.createCell(2).setCellValue(entry.getValue());
            }

            // Calculate grand total
            int grandTotal = 0;
            for (Integer value : pivotJumlahTindakan.values()) {
                grandTotal += value;
            }

            // Write grand total to Pertindakan sheet
            Row rowx = Pertindakan.createRow(rowNum);
            rowx.createCell(0).setCellValue("");
            rowx.createCell(1).setCellValue("Grand Total");
            rowx.createCell(2).setCellValue(grandTotal);


            System.out.println ("03. " + BookPertindakanNew.getSheetAt (3).getSheetName () + " Complete");


//        buat sheet 4 Jml tndakan per cr Byr pr hri
            BookPertindakanNew.createSheet ();
            Sheet TndkanCrByrHr = BookPertindakanNew.getSheetAt (4);
            BookPertindakanNew.setSheetName (4, "2.Jml tndakan per cr Byr pr hri");

            TndkanCrByrHr.createRow (0);
            TndkanCrByrHr.createRow (1);
            TndkanCrByrHr.createRow (2);

            Set<String> crByrxTnd = new TreeSet<>();
            Set<String> tanggalRegist = new TreeSet<>();
            Map<String, Map<String, Integer>> countMap = new HashMap<>(); // new count map

            for (int row = 1; row <= Genap.getLastRowNum(); row++) {
                String cellcrByr = Genap.getRow(row).getCell(8).getStringCellValue();
                String cellTindakan = Genap.getRow(row).getCell(32).getStringCellValue();
                String crByrxTindakan = cellcrByr + "T.T" + cellTindakan; // I use T.T because i got no idea dot coma dash etc are used
                String cellTanggalReg = Genap.getRow(row).getCell(31).getStringCellValue();

                tanggalRegist.add(cellTanggalReg);
                crByrxTnd.add(crByrxTindakan);

                // increment count in countMap
                if (!countMap.containsKey(crByrxTindakan)) {
                    countMap.put(crByrxTindakan, new HashMap<>());
                }
                Map<String, Integer> tanggalRegCountMap = countMap.get(crByrxTindakan);
                if (!tanggalRegCountMap.containsKey(cellTanggalReg)) {
                    tanggalRegCountMap.put(cellTanggalReg, 1);
                } else {
                    tanggalRegCountMap.put(cellTanggalReg, tanggalRegCountMap.get(cellTanggalReg) + 1);
                }
            }

            // Convert crByrxTnd set to a list and sort it
            List<String> sortedCrByrxTnd = new ArrayList<>(crByrxTnd);
            Collections.sort(sortedCrByrxTnd);

            TndkanCrByrHr.createRow(0).createCell(0).setCellValue("Tanggal");
            int rowStart = 1;
            for (String konten : sortedCrByrxTnd) {
                String[] splitValue = konten.split("T.T");
                String crByr = splitValue[0];
                String tnd = splitValue[1];
                TndkanCrByrHr.getRow(0).createCell(rowStart).setCellValue(crByr);
                TndkanCrByrHr.getRow(1).createCell(rowStart).setCellValue(tnd);
                rowStart++;
            }
            rowStart = 2;
            for (String konten : tanggalRegist) {
                int colStart = 1;
                TndkanCrByrHr.createRow(rowStart).createCell(0).setCellValue(konten);
                for (String item : sortedCrByrxTnd) {
                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
                        TndkanCrByrHr.getRow(rowStart).createCell(colStart++).setCellValue(countMap.get(item).get(konten));
                    } else {
                        TndkanCrByrHr.getRow(rowStart).createCell(colStart++).setCellValue(0);
                    }
                }
                rowStart++;
            }



            System.out.println ("04. Sheet " + BookPertindakanNew.getSheetAt (4).getSheetName () + " Created");


//        buat sheet 5 Jml pasien per cr Byr pr hri
            BookPertindakanNew.createSheet ();
            Sheet PsnCrByrHr = BookPertindakanNew.getSheetAt (5);
            BookPertindakanNew.setSheetName (5, "3.Pasien per cara bayar pr hari");

            PsnCrByrHr.createRow (0);
            PsnCrByrHr.createRow (1);

            Set<String> crByr = new TreeSet<>();
            tanggalRegist = new TreeSet<>();
            countMap = new HashMap<>(); // new count map

            for (int row = 1; row <= Ganjil.getLastRowNum(); row++) {
                String cellcrByr = Ganjil.getRow(row).getCell(1).getStringCellValue();
                String cellTanggalReg = Ganjil.getRow(row).getCell(2).getStringCellValue();

                tanggalRegist.add(cellTanggalReg);
                crByr.add(cellcrByr);

                // increment count in countMap
                if (!countMap.containsKey(cellcrByr)) {
                    countMap.put(cellcrByr, new HashMap<>());
                }
                Map<String, Integer> tanggalRegCountMap = countMap.get(cellcrByr);
                if (!tanggalRegCountMap.containsKey(cellTanggalReg)) {
                    tanggalRegCountMap.put(cellTanggalReg, 1);
                } else {
                    tanggalRegCountMap.put(cellTanggalReg, tanggalRegCountMap.get(cellTanggalReg) + 1);
                }
            }


            PsnCrByrHr.createRow(0).createCell(0).setCellValue("Tanggal");
            rowStart = 1;
            for (String konten : crByr) {
                PsnCrByrHr.getRow(0).createCell(rowStart).setCellValue(konten);
                rowStart++;
            }
            rowStart = 1;
            for (String konten : tanggalRegist) {
                int colStart = 1;
                PsnCrByrHr.createRow(rowStart).createCell(0).setCellValue(konten);
                for (String item : crByr) {
                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
                        PsnCrByrHr.getRow(rowStart).createCell(colStart++).setCellValue(countMap.get(item).get(konten));
                    } else {
                        PsnCrByrHr.getRow(rowStart).createCell(colStart++).setCellValue(0);
                    }
                }
                rowStart++;
            }



            System.out.println ("05. Sheet " + BookPertindakanNew.getSheetAt (5).getSheetName () + " Created");


//        buat sheet 6 Jml tndakan per cr Byr
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

            System.out.println ("06. " + BookPertindakanNew.getSheetAt (6).getSheetName () + " Completed");


//        buat sheet 7 Jml pasien per cr Byr
            BookPertindakanNew.createSheet ();
            Sheet PsnCrByr = BookPertindakanNew.getSheetAt (7);
            BookPertindakanNew.setSheetName (7, "5.Pasien per cara bayar");
            System.out.println ("07. " + BookPertindakanNew.getSheetAt (7).getSheetName () + " Start");
            Map<String, Integer> PsncrByrCount = new TreeMap<> ();
            for (int row = 1; row <= Ganjil.getLastRowNum (); row++) {
                String caraByr = Ganjil.getRow (row).getCell (1).getStringCellValue ();
                PsncrByrCount.put (caraByr, PsncrByrCount.getOrDefault (caraByr, 0) + 1);
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

            System.out.println ("07. " + BookPertindakanNew.getSheetAt (7).getSheetName () + " Completed");

//        buat sheet 8 Jml tndakan per inst asal
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

            System.out.println ("08. " + BookPertindakanNew.getSheetAt (8).getSheetName () + " Completed");

//        buat sheet 9 Jml pasien per inst asal
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

            System.out.println ("09. " + BookPertindakanNew.getSheetAt (9).getSheetName () + " Completed");

//        buat sheet 10 Jml pendapatan
            BookPertindakanNew.createSheet ();
            Sheet pendapatan = BookPertindakanNew.getSheetAt (10);
            BookPertindakanNew.setSheetName (10, "8.Jumlah pendapatan");
            System.out.println ("10. " + BookPertindakanNew.getSheetAt (10).getSheetName () + " Start");

            tanggalRegist = new TreeSet<>();
            Set<String> caraByr = new TreeSet<>();
            Map<String, Map<String, Integer>> countMaptagihan = new HashMap<>(); // fixed countMap variable name
            Map<String, Map<String, Integer>> totalMaptagihan = new HashMap<>(); // added totalMap variable
            for (int row = 1; row <= Genap.getLastRowNum(); row++) {
                String cellTanggalReg = Genap.getRow(row).getCell(31).getStringCellValue();
                String cellCaraBayar = Genap.getRow(row).getCell(8).getStringCellValue();
                Integer cellTotalTarif = (int) Genap.getRow(row).getCell(19).getNumericCellValue(); // cast to int

                tanggalRegist.add(cellTanggalReg);
                caraByr.add(cellCaraBayar);

                // increment count in countMap
                if (!countMaptagihan.containsKey(cellTanggalReg)) { // swapped key and value
                    countMaptagihan.put(cellTanggalReg, new HashMap<>()); // swapped key and value
                }
                Map<String, Integer> tagihanCountMap = countMaptagihan.get(cellTanggalReg); // swapped key and value
                if (!tagihanCountMap.containsKey(cellCaraBayar)) { // swapped key and value
                    tagihanCountMap.put(cellCaraBayar, 1); // swapped key and value
                } else {
                    tagihanCountMap.put(cellCaraBayar, tagihanCountMap.get(cellCaraBayar) + 1); // increment count
                }

                // calculate total in totalMap
                if (!totalMaptagihan.containsKey(cellTanggalReg)) { // swapped key and value
                    totalMaptagihan.put(cellTanggalReg, new HashMap<>()); // swapped key and value
                }
                Map<String, Integer> tagihanTotalMap = totalMaptagihan.get(cellTanggalReg); // swapped key and value
                if (!tagihanTotalMap.containsKey(cellCaraBayar)) { // swapped key and value
                    tagihanTotalMap.put(cellCaraBayar, cellTotalTarif);
                } else {
                    tagihanTotalMap.put(cellCaraBayar, tagihanTotalMap.get(cellCaraBayar) + cellTotalTarif); // increment total
                }
            }

            pendapatan.createRow(0).createCell(0).setCellValue("Tanggal");
            rowStart = 1;
            for (String klmn : caraByr) { // swapped loop
                pendapatan.getRow(0).createCell(rowStart++).setCellValue(klmn); // swapped loop
            }
            rowStart = 1;
            for (String tgl : tanggalRegist) { // swapped loop
                int colStart = 1;
                pendapatan.createRow(rowStart).createCell(0).setCellValue(tgl); // swapped loop
                for (String klmn : caraByr) { // swapped loop
                    if (countMaptagihan.containsKey(tgl) && countMaptagihan.get(tgl).containsKey(klmn)) { // swapped key and value
                        pendapatan.getRow(rowStart).createCell(colStart++).setCellValue(totalMaptagihan.get(tgl).get(klmn)); // get total from totalMap, swapped key and value
                    } else {
                        pendapatan.getRow(rowStart).createCell(colStart++).setCellValue(0);
                    }
                }
                rowStart++;
            }

            System.out.println ("10. " + BookPertindakanNew.getSheetAt (10).getSheetName () + " Completed");

//        buat sheet 11 Jml tndakan per cr Byr pr hri
            InputStream monitoringF1 = new FileInputStream (fileInput + fileNameMonitoringf1 + ".xlsx");
            InputStream monitoringF2 = new FileInputStream (fileInput + fileNameMonitoringf2 + ".xlsx");
            InputStream _pelayananPenunjang = new FileInputStream (fileInput + fileNamePelayananPenunjang + ".xlsx");
            InputStream _responTime = new FileInputStream (fileInput + fileNameResponTime + ".xlsx");
            Workbook bookMonitoringHasilRadF1 = new XSSFWorkbook (monitoringF1);
            Workbook bookMonitoringHasilRadF2 = new XSSFWorkbook (monitoringF2);
            Workbook bookPelayananPenunjang = new XSSFWorkbook (_pelayananPenunjang);
            Workbook bookRespontime = new XSSFWorkbook (_responTime);
            BookPertindakanNew.createSheet ("9. Monitoring Hasil Rad");
            Sheet monitoringHasilRad = BookPertindakanNew.getSheetAt (11);
            Sheet monitoringHasilRadF1 = bookMonitoringHasilRadF1.getSheetAt (0);
            Sheet monitoringHasilRadF2 = bookMonitoringHasilRadF2.getSheetAt (0);
            Sheet pelayananPenunjang = bookPelayananPenunjang.getSheetAt (0);
            Sheet responTime = bookRespontime.getSheetAt (0);
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
            HashMap<String, List<Integer>> sampleIDHashMap = new HashMap<> ();
            for (int monitoringRow = 1; monitoringRow <= monitoringHasilRad.getLastRowNum (); monitoringRow++) {
                String sampleID = monitoringHasilRad.getRow (monitoringRow).getCell (6).getStringCellValue ();
                if (!sampleIDHashMap.containsKey (sampleID)) {
                    sampleIDHashMap.put (sampleID, new ArrayList<> ());
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

            HashMap<String, List<Integer>> noRegHashMap = new HashMap<> ();
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
            monitoringHasilRad.getRow (0).createCell (15).setCellValue ("RESPON TIME RS");
            monitoringHasilRad.getRow (0).createCell (16).setCellValue ("RESPON TIME RAD");
            monitoringHasilRad.getRow (0).createCell (17).setCellValue ("TOTAL RS");
            monitoringHasilRad.getRow (0).createCell (18).setCellValue ("TOTAL RAD");
            monitoringHasilRad.getRow (0).createCell (19).setCellValue ("RATA2 RS");
            monitoringHasilRad.getRow (0).createCell (20).setCellValue ("RATA2 RAD");
            monitoringHasilRad.getRow (0).createCell (21).setCellValue ("CITO");

            HashMap<String, List<Integer>> noregHashMap = new HashMap<> ();
            for (int monitoringRow =1;monitoringRow<=monitoringHasilRad.getLastRowNum ();monitoringRow++){
                String noreg = monitoringHasilRad.getRow (monitoringRow).getCell (0).getStringCellValue ();
                if (!noregHashMap.containsKey (noreg)){
                    noregHashMap.put (noreg,new ArrayList<> ());
                }
                noregHashMap.get (noreg).add (monitoringRow);
            }

            for (int row=1;row<=monitoringHasilRadF2.getLastRowNum ();row++){
                String noreg = monitoringHasilRadF2.getRow (row).getCell (0).getStringCellValue ()+
                        monitoringHasilRadF2.getRow (row).getCell (1).getStringCellValue ()+
                        monitoringHasilRadF2.getRow (row).getCell (2).getStringCellValue ()+
                        monitoringHasilRadF2.getRow (row).getCell (3).getStringCellValue ()+
                        monitoringHasilRadF2.getRow (row).getCell (4).getStringCellValue ();
                List<Integer> monitoringRows = noregHashMap.get (noreg);
                Cell menyerahkan = monitoringHasilRadF2.getRow (row).getCell (20);
                Cell menerima = monitoringHasilRadF2.getRow (row).getCell (21);
                if (monitoringRows!=null){
                    for (Integer monitoringRow:monitoringRows) {
                        if (menyerahkan != null) {
                            monitoringHasilRad.getRow (monitoringRow).createCell (13).setCellValue (
                                    menyerahkan.getStringCellValue ()
                            );
                            monitoringHasilRad.getRow (monitoringRow).createCell (14).setCellValue (
                                    menerima.getStringCellValue ()
                            );
                        } else {
                            monitoringHasilRad.getRow (monitoringRow).createCell (13).setCellValue ("-");
                            monitoringHasilRad.getRow (monitoringRow).createCell (14).setCellValue ("-");
                        }
                    }
                }
            }

            // First, create a Map to store the values from the responTime sheet
            Map<String, String[]> responTimeMap = new HashMap<>();
            int responTimeLastRowNum = responTime.getLastRowNum(); //where is last row?
            for (int row = 1; row <= responTimeLastRowNum; row++) {
                //make sure no "PAKET" get mapped
                if (!responTime.getRow (row).getCell (18).getStringCellValue ().contains ("PAKET")) {
                    //make array to contain it, and stuff everything to it
                    String[] values = new String[26];
                    for (int col = 0; col < 26; col++) {
                        Cell cell = responTime.getRow (row).getCell (col);
                        if (cell == null) {
                            values[col] = "-";
                        } else if (cell.getCellType () == CellType.NUMERIC) {
                            values[col] = String.valueOf (cell.getNumericCellValue ());
                        } else if (cell.getCellType () == CellType.STRING) {
                            values[col] = cell.getStringCellValue ();
                        }
                    }
                    String noreg = values[0] + values[1] + values[2] + values[3] + values[4];
                    responTimeMap.put (noreg, values);
                }
            }

//            // Create a map to store the values from the responTime sheet
//            Map<String, String[]> responTimeMap = new HashMap<>();
//
//            // Get the last row number of the responTime sheet
//            int responTimeLastRowNum = responTime.getLastRowNum();
//
//            // Store all rows in an array for faster access
//            Row[] rows = new Row[responTimeLastRowNum + 1];
//            Iterator<Row> rowIterator = responTime.iterator();
//            int rowIndex = 0;
//
//            // Iterate over the rows of the responTime sheet and store them in the rows array
//            while (rowIterator.hasNext()) {
//                Row row = rowIterator.next();
//                rows[rowIndex++] = row;
//            }
//
//            // Iterate over each row (starting from row 1, skipping header row)
//            for (int row = 1; row <= responTimeLastRowNum; row++) {
//                Row currentRow = rows[row];
//                Cell valueCell = currentRow.getCell(18);
//
//                // Check if the value in column 18 does not contain "PAKET"
//                if (!valueCell.getStringCellValue().contains("PAKET")) {
//                    String[] values = new String[26];
//
//                    // Iterate over each column
//                    for (int col = 0; col < 26; col++) {
//                        Cell cell = currentRow.getCell(col);
//
//                        // Get the cell value and store it in the values array
//                        String cellValue = (cell != null) ? getCellValueAsString(cell) : "-";
//                        values[col] = cellValue;
//                    }
//
//                    // Generate the noreg by concatenating the first 5 values
//                    StringBuilder sb = new StringBuilder();
//                    for (int i = 0; i < 5; i++) {
//                        sb.append(values[i]);
//                    }
//                    String noreg = sb.toString();
//
//                    // Add the noreg and values array to the responTimeMap
//                    responTimeMap.put(noreg, values);
//                }
//            }


            // Loop through the monitoringHasilRad sheet and use the Map to retrieve values from the responTime sheet
            int monitoringHasilRadLastRowNum = monitoringHasilRad.getLastRowNum();
            for (int row = 1; row <= monitoringHasilRadLastRowNum; row++) {
                Row getMonitoringHasilRadRow = monitoringHasilRad.getRow(row);
                String noreg = getMonitoringHasilRadRow.getCell(0).getStringCellValue();
                String[] respontimeValues = responTimeMap.get(noreg);
                if (respontimeValues != null && !respontimeValues[18].contains("PAKET") && respontimeValues[11] != null) {
                    getMonitoringHasilRadRow.createCell(15).setCellValue(respontimeValues[12]);
                    getMonitoringHasilRadRow.createCell(16).setCellValue(respontimeValues[13]);
                    getMonitoringHasilRadRow.createCell(17).setCellValue(respontimeValues[14]);
                    getMonitoringHasilRadRow.createCell(18).setCellValue(respontimeValues[15]);
                    getMonitoringHasilRadRow.createCell(19).setCellValue(respontimeValues[16]);
                    getMonitoringHasilRadRow.createCell(20).setCellValue(respontimeValues[17]);
                    if (respontimeValues[25].equals ("1")) {
                        getMonitoringHasilRadRow.createCell (21).setCellValue ("CITO");
                    }else {
                        getMonitoringHasilRadRow.createCell (21).setCellValue ("Tidak CITO");
                    }
                } else {
                    getMonitoringHasilRadRow.createCell(15).setCellValue("-");
                    getMonitoringHasilRadRow.createCell(16).setCellValue("-");
                    getMonitoringHasilRadRow.createCell(17).setCellValue("-");
                    getMonitoringHasilRadRow.createCell(18).setCellValue("-");
                    getMonitoringHasilRadRow.createCell(19).setCellValue("-");
                    getMonitoringHasilRadRow.createCell(20).setCellValue("-");
                    getMonitoringHasilRadRow.createCell(21).setCellValue("-");
                }
            }

            System.out.println ("11. "+ BookPertindakanNew.getSheetAt (11).getSheetName ()+" Completed");


//        buat sheet 12 ResponTime
            BookPertindakanNew.createSheet ();
            Sheet sheetResponTime = BookPertindakanNew.getSheetAt (12);
            BookPertindakanNew.setSheetName (12, "10.Respon Time");
            System.out.println ("12. " + BookPertindakanNew.getSheetAt (12).getSheetName () + " Start");

            //count inst asal
            Map<String, Integer> keteranganCount = new TreeMap<> ();
            for (int row = 1; row <= monitoringHasilRad.getLastRowNum (); row++) {
                String keterangan = monitoringHasilRad.getRow (row).getCell (   10).getStringCellValue ();
                keteranganCount.put (keterangan, keteranganCount.getOrDefault (keterangan, 0) + 1);
            }

            sheetResponTime.createRow (0).createCell (0).setCellValue ("Keterangan");
            sheetResponTime.getRow (0).createCell (1).setCellValue ("Jumlah");

            int keteranganResponTime = 0;
            int keteranganRespontimeSum = 0;
            for (Map.Entry<String, Integer> entry : keteranganCount.entrySet ()) {
                keteranganResponTime++;
                sheetResponTime.createRow (keteranganResponTime).createCell (0).setCellValue (entry.getKey ());
                sheetResponTime.getRow (keteranganResponTime).createCell (1).setCellValue (entry.getValue ());
                keteranganRespontimeSum += entry.getValue ();

                int TndInstAsalLastRow = sheetResponTime.getLastRowNum () + 1;
                sheetResponTime.createRow (TndInstAsalLastRow).createCell (0).setCellValue ("Grand Total");
                sheetResponTime.getRow (TndInstAsalLastRow).createCell (1).setCellValue (keteranganRespontimeSum);
            }

            System.out.println ("12. "+ BookPertindakanNew.getSheetAt (12).getSheetName ()+" Completed");

            // buat sheet 13 Pasien Kebidanan Kandungan
            BookPertindakanNew.createSheet ();
            Sheet kebidanan = BookPertindakanNew.getSheetAt (13);
            BookPertindakanNew.setSheetName (13, "11. Kebidanan Kandungan");
            kebidanan.createRow (0).createCell (0).setCellValue ("NOREG");
            kebidanan.getRow (0).createCell (1).setCellValue ("NO");
            kebidanan.getRow (0).createCell (2).setCellValue ("TGL MASUK");
            kebidanan.getRow (0).createCell (3).setCellValue ("TGL TINDAKAN");
            kebidanan.getRow (0).createCell (4).setCellValue ("RM");
            kebidanan.getRow (0).createCell (5).setCellValue ("NAMA PASIEN");
            kebidanan.getRow (0).createCell (6).setCellValue ("NAMA PELAKSANA");
            kebidanan.getRow (0).createCell (7).setCellValue ("NAMA TINDAKAN");
            kebidanan.getRow (0).createCell (8).setCellValue ("INSTALASI ASAL");
            kebidanan.getRow (0).createCell (9).setCellValue ("KETERANGAN");

            BookPertindakanNew.createSheet ();
            Sheet ponek = BookPertindakanNew.getSheetAt (14);
            BookPertindakanNew.setSheetName (14, "12. Ponek");
            ponek.createRow (0).createCell (0).setCellValue ("NOREG");
            ponek.getRow (0).createCell (1).setCellValue ("NO");
            ponek.getRow (0).createCell (2).setCellValue ("TGL MASUK");
            ponek.getRow (0).createCell (3).setCellValue ("TGL TINDAKAN");
            ponek.getRow (0).createCell (4).setCellValue ("RM");
            ponek.getRow (0).createCell (5).setCellValue ("NAMA PASIEN");
            ponek.getRow (0).createCell (6).setCellValue ("NAMA PELAKSANA");
            ponek.getRow (0).createCell (7).setCellValue ("NAMA TINDAKAN");
            ponek.getRow (0).createCell (8).setCellValue ("INSTALASI ASAL");
            ponek.getRow (0).createCell (9).setCellValue ("KETERANGAN");

            BookPertindakanNew.createSheet ();
            Sheet isolasi = BookPertindakanNew.getSheetAt (15);
            BookPertindakanNew.setSheetName (15, "13. Isolasi");
            isolasi.createRow (0).createCell (0).setCellValue ("NOREG");
            isolasi.getRow (0).createCell (1).setCellValue ("NO");
            isolasi.getRow (0).createCell (2).setCellValue ("TGL MASUK");
            isolasi.getRow (0).createCell (3).setCellValue ("TGL TINDAKAN");
            isolasi.getRow (0).createCell (4).setCellValue ("RM");
            isolasi.getRow (0).createCell (5).setCellValue ("NAMA PASIEN");
            isolasi.getRow (0).createCell (6).setCellValue ("NAMA PELAKSANA");
            isolasi.getRow (0).createCell (7).setCellValue ("NAMA TINDAKAN");
            isolasi.getRow (0).createCell (8).setCellValue ("INSTALASI ASAL");
            isolasi.getRow (0).createCell (9).setCellValue ("KETERANGAN");

            int rowPonek=1;
            int rowIsolasi=1;
            int rowKebidananKandungan = 1;
            for (Row row : Genap) {
                // Cache frequently used cells
                Cell noregCell = row.getCell (29);
                Cell regdateCell = row.getCell (9);
                Cell rmCell = row.getCell (6);
                Cell nameCell = row.getCell (5);
                Cell namaTindakanCell = row.getCell (15);
                Cell nickInstAsalCell = row.getCell (24);
                if (row.getCell (28).getStringCellValue ().equals ("Kebidanan dan Kandungan")) {
                    boolean matchFound = false;  // initialize flag variable
                    for (Row pelayananRow : pelayananPenunjang) {
                        String noregPelayananPenunjang = pelayananRow.getCell (0).getStringCellValue () +
                                pelayananRow.getCell (1).getStringCellValue () +
                                pelayananRow.getCell (2).getStringCellValue () +
                                pelayananRow.getCell (3).getStringCellValue () +
                                pelayananRow.getCell (4).getStringCellValue ();
                        if (Objects.equals (noregCell.getStringCellValue (), noregPelayananPenunjang)) {
                            if (!matchFound) {  // check flag before printing row number
                                matchFound = true;  // set flag to true
                                Row kebidananRow = kebidanan.createRow (rowKebidananKandungan);
                                kebidananRow.createCell (0).setCellValue (noregCell.getStringCellValue ());
                                kebidananRow.createCell (1).setCellValue (rowKebidananKandungan);
                                kebidananRow.createCell (2).setCellValue (regdateCell.getStringCellValue ());
                                kebidananRow.createCell (3).setCellValue (regdateCell.getStringCellValue ());
                                kebidananRow.createCell (4).setCellValue (rmCell.getStringCellValue ());
                                kebidananRow.createCell (5).setCellValue (nameCell.getStringCellValue ());
                                kebidananRow.createCell (6).setCellValue (pelayananRow.getCell (27).getStringCellValue ());
                                kebidananRow.createCell (7).setCellValue (namaTindakanCell.getStringCellValue ());
                                kebidananRow.createCell (8).setCellValue (nickInstAsalCell.getStringCellValue ());
                                for (Row monitoringRow : monitoringHasilRad) {
                                    if (Objects.equals (noregCell.getStringCellValue (), monitoringRow.getCell (0).getStringCellValue ())) {
                                        kebidananRow.createCell (9).setCellValue (monitoringRow.getCell (10).getStringCellValue ());
                                        break;
                                    }
                                }
                                rowKebidananKandungan++;
                            }
                        }
                    }
                } else if (row.getCell (28).getStringCellValue ().equals ("Ponek")) {
                    boolean matchFound = false;  // initialize flag variable
                    for (Row pelayananRow : pelayananPenunjang) {
                        String noregPelayananPenunjang = pelayananRow.getCell (0).getStringCellValue () +
                                pelayananRow.getCell (1).getStringCellValue () +
                                pelayananRow.getCell (2).getStringCellValue () +
                                pelayananRow.getCell (3).getStringCellValue () +
                                pelayananRow.getCell (4).getStringCellValue ();
                        if (Objects.equals (noregCell.getStringCellValue (), noregPelayananPenunjang)) {
                            if (!matchFound) {  // check flag before printing row number
                                matchFound = true;  // set flag to true
                                Row ponekRow = ponek.createRow (rowPonek);
                                ponekRow.createCell (0).setCellValue (noregCell.getStringCellValue ());
                                ponekRow.createCell (1).setCellValue (rowPonek);
                                ponekRow.createCell (2).setCellValue (regdateCell.getStringCellValue ());
                                ponekRow.createCell (3).setCellValue (regdateCell.getStringCellValue ());
                                ponekRow.createCell (4).setCellValue (rmCell.getStringCellValue ());
                                ponekRow.createCell (5).setCellValue (nameCell.getStringCellValue ());
                                ponekRow.createCell (6).setCellValue (pelayananRow.getCell (27).getStringCellValue ());
                                ponekRow.createCell (7).setCellValue (namaTindakanCell.getStringCellValue ());
                                ponekRow.createCell (8).setCellValue (nickInstAsalCell.getStringCellValue ());
                                for (Row monitoringRow : monitoringHasilRad) {
                                    if (Objects.equals (noregCell.getStringCellValue (), monitoringRow.getCell (0).getStringCellValue ())) {
                                        ponekRow.createCell (9).setCellValue (monitoringRow.getCell (10).getStringCellValue ());
                                        break;
                                    }
                                }
                                rowPonek++;
                            }
                        }
                    }

                } else if (row.getCell (28).getStringCellValue ().equals ("ISOLASI")) {
                    boolean matchFound = false;  // initialize flag variable
                    for (Row pelayananRow : pelayananPenunjang) {
                        String noregPelayananPenunjang = pelayananRow.getCell (0).getStringCellValue () +
                                pelayananRow.getCell (1).getStringCellValue () +
                                pelayananRow.getCell (2).getStringCellValue () +
                                pelayananRow.getCell (3).getStringCellValue () +
                                pelayananRow.getCell (4).getStringCellValue ();
                        if (Objects.equals (noregCell.getStringCellValue (), noregPelayananPenunjang)) {
                            if (!matchFound) {  // check flag before printing row number
                                matchFound = true;  // set flag to true
                                Row isolasiRow = isolasi.createRow (rowIsolasi);
                                isolasiRow.createCell (0).setCellValue (noregCell.getStringCellValue ());
                                isolasiRow.createCell (1).setCellValue (rowIsolasi);
                                isolasiRow.createCell (2).setCellValue (regdateCell.getStringCellValue ());
                                isolasiRow.createCell (3).setCellValue (regdateCell.getStringCellValue ());
                                isolasiRow.createCell (4).setCellValue (rmCell.getStringCellValue ());
                                isolasiRow.createCell (5).setCellValue (nameCell.getStringCellValue ());
                                isolasiRow.createCell (6).setCellValue (pelayananRow.getCell (27).getStringCellValue ());
                                isolasiRow.createCell (7).setCellValue (namaTindakanCell.getStringCellValue ());
                                isolasiRow.createCell (8).setCellValue (nickInstAsalCell.getStringCellValue ());
                                for (Row monitoringRow : monitoringHasilRad) {
                                    if (Objects.equals (noregCell.getStringCellValue (), monitoringRow.getCell (0).getStringCellValue ())) {
                                        isolasiRow.createCell (9).setCellValue (monitoringRow.getCell (10).getStringCellValue ());
                                        break;
                                    }
                                }
                                rowIsolasi++;
                            }
                        }
                    }
                }
            }

            // Sheet numbers 3 to 15
            for (int sheetNum = 3; sheetNum <= 15; sheetNum++) {
                Sheet currentSheet = BookPertindakanNew.getSheetAt(sheetNum);
                System.out.println (currentSheet.getSheetName ());
                for (int rightCell = 0; rightCell < currentSheet.getRow(0).getLastCellNum(); rightCell++) {
                    currentSheet.getRow(0).getCell(rightCell).setCellStyle(BorderCenterCellStyle);
                    currentSheet.autoSizeColumn(rightCell);
                    for (int downRow = 1; downRow <= currentSheet.getLastRowNum(); downRow++) {
                        if (currentSheet.getRow (downRow).getCell (rightCell)==null){
                            System.out.println (downRow);
                            currentSheet.getRow (downRow).createCell (rightCell).setCellValue ("");
                        }
                        currentSheet.getRow(downRow).getCell(rightCell).setCellStyle(AllBorderCellStyle);
                    }
                }
            }

            if (doneFinal){
                BookPertindakanNew.removeSheetAt (0);
                BookPertindakanNew.removeSheetAt (0);
                BookPertindakanNew.removeSheetAt (0);
            }




        } catch (Exception e) {
            e.printStackTrace();
        }


        try {
            if (doneFinal){
                outputStream = new FileOutputStream (fileOutput+"Done Rad "+localDate+".xlsx");
            } else {
                outputStream = new FileOutputStream (fileNamePertindakanNew + " half done.xlsx");
            }
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

    // Helper method to retrieve cell value as a String
    private String getCellValueAsString(Cell cell) {
        return switch (cell.getCellType ()) {
            case NUMERIC -> String.valueOf (cell.getNumericCellValue ());
            case STRING -> cell.getStringCellValue ();
            default -> "-";
        };
    }

}
