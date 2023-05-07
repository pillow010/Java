package LaporanLab;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class LabDone {

    public static void main(String[] args) {
        Workbook BookPertindakanNew = null;
        FileOutputStream outputStream = null;
        boolean doneFinal = true;

        String localDate = LocalDate.now ().minusMonths (1).format (DateTimeFormatter.ofPattern ("yy MM"));
        String fileNamePertindakanNew = localDate + " lab tindakan new";
//        String fileNameHasilRinci = localDate + " lab hasil rinci";
        String fileNameRegister = localDate + " lab register";


        try {
            InputStream pertindakanNew = new FileInputStream ("C:\\sat work\\test\\" + fileNamePertindakanNew + ".xlsx");
            BookPertindakanNew = new XSSFWorkbook (pertindakanNew);
            InputStream register = new FileInputStream ("C:\\sat work\\test\\" + fileNameRegister + ".xlsx");
            Workbook bookRegister = new XSSFWorkbook (register);

//          taruh pertindakan new ke sheet 0
            Sheet pertindakan_New_Raw = BookPertindakanNew.getSheetAt (0);
            Sheet noDuplicate = BookPertindakanNew.createSheet ();
            BookPertindakanNew.setSheetName (1, "noDuplicate");
            Sheet Pertindakan = BookPertindakanNew.createSheet ();
            BookPertindakanNew.setSheetName (2, "1. Pertindakan");
//            System.out.println ("03. " + BookPertindakanNew.getSheetAt (2).getSheetName () + " Start");

            Sheet registerSheet = bookRegister.getSheetAt (0);
            BookPertindakanNew.setSheetName (0, "pertindakan_New_Raw");
            System.out.println ("00. Doing " + BookPertindakanNew.getSheetAt (0).getSheetName () + ", "
                    + BookPertindakanNew.getSheetAt (1).getSheetName () + ", "
                    + BookPertindakanNew.getSheetAt (2).getSheetName ());

//          hashmap sub-inst
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
            subInstMap.put ("IRNA25", "BOUGENVILLE 2");
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
            pertindakan_New_Raw.getRow (0).createCell (31).setCellValue ("KELAMIN");

//          buat judul dan kasih kotak untuk sheet pertindakan
            Pertindakan.createRow (0).createCell (0).setCellValue ("NO");
            Pertindakan.getRow (0).createCell (1).setCellValue ("Nama Tindakan");
            Pertindakan.getRow (0).createCell (2).setCellValue ("Jumlah");


//          create a header row for noDuplicate sheet
            Row headerRow = noDuplicate.createRow (0);
            for (int i = 0; i < pertindakan_New_Raw.getRow (0).getLastCellNum (); i++) {
                Cell headerCell = headerRow.createCell (i);
                headerCell.setCellValue (pertindakan_New_Raw.getRow (0).getCell (i).getStringCellValue ());
            }

//          fill cell 28-31
            Set<String> uniqueValues = new HashSet<> ();
            Map<String, Integer> pivotJumlahTindakan = new HashMap<> ();
            int actualRow = 1;
            for (int _row = 1; _row <= pertindakanNewRawLastRowNum (pertindakan_New_Raw); _row++) {
                Row row = pertindakan_New_Raw.getRow (_row);
                Cell nickInst = row.getCell (24);
                Cell kdSubInstAsal = row.getCell (27);
                Cell kdInst = row.getCell (0);
                Cell kdSubInst = row.getCell (1);
                Cell kdDtlSubInst = row.getCell (2);
                Cell kdPengunjung = row.getCell (3);
                Cell kdKunjungan = row.getCell (4);
                String noreg = kdInst.getStringCellValue () + kdSubInst.getStringCellValue ()
                        + kdDtlSubInst.getStringCellValue () + kdPengunjung.getStringCellValue ()
                        + kdKunjungan.getStringCellValue ();

                //if null rujukan luar rs
                if (nickInst == null) {
                    row.createCell (28).setCellValue ("RUJUKAN LUAR RS");
                    row.createCell (24).setCellValue ("RUJUKAN LUAR RS");
                } else {
                    //not null combine inst and sub inst then get the value
                    String value = nickInst.getStringCellValue () + kdSubInstAsal.getStringCellValue ();
                    //is there new sub inst?
                    row.createCell (28).setCellValue (subInstMap.getOrDefault (value, "Not Found"));
                }

                row.createCell (29).setCellValue (noreg);
                row.createCell (30).setCellValue (row.getCell (29).getStringCellValue () + row.getCell (15).getStringCellValue ());


                if (pertindakan_New_Raw.getRow (_row) != null) { // check if row is not empty
                    Cell cell = pertindakan_New_Raw.getRow (_row).getCell (29);
                    String cellValue = cell.getStringCellValue ();
                    if (!cellValue.isBlank () && !uniqueValues.contains (cellValue)) { // check if cell is not empty
                        Row newRow = noDuplicate.createRow (actualRow++);
                        for (int i = 0; i <= row.getLastCellNum (); i++) {
                            Cell oldCell = row.getCell (i);
                            Cell newCell = newRow.createCell (i);
                            if (oldCell != null) {
                                if (oldCell.getCellType () == CellType.STRING) {
                                    newCell.setCellValue (oldCell.getStringCellValue ());
                                } else if (oldCell.getCellType () == CellType.NUMERIC) {
                                    newCell.setCellValue (oldCell.getNumericCellValue ());
                                }
                            } else {
                                newCell.setCellValue ("");
                            }
                        }
                        uniqueValues.add (cellValue);
                    }
                }

                String Tindakan = pertindakan_New_Raw.getRow (_row).getCell (15).getStringCellValue ();
                if (!Tindakan.contains ("PAKET")) {
                    Integer count = pivotJumlahTindakan.getOrDefault (Tindakan, 0);
                    count++;
                    pivotJumlahTindakan.put (Tindakan, count);
                }
            }


//          Create a hash table to store the registerSheet data
            Map<String, String> registerData = new HashMap<> ();
            for (int rowRegister = 1; rowRegister <= registerSheet.getLastRowNum (); rowRegister++) {
                String noregRegister = registerSheet.getRow (rowRegister).getCell (0).getStringCellValue ()
                        + registerSheet.getRow (rowRegister).getCell (1).getStringCellValue ()
                        + registerSheet.getRow (rowRegister).getCell (2).getStringCellValue ()
                        + registerSheet.getRow (rowRegister).getCell (3).getStringCellValue ()
                        + registerSheet.getRow (rowRegister).getCell (4).getStringCellValue ();
                String kelamin = registerSheet.getRow (rowRegister).getCell (10).getStringCellValue ();
                registerData.put (noregRegister, kelamin);
            }

//          Set the values for the new sheet using the hash table
            for (int rowNoDuplicate = 1; rowNoDuplicate <= noDuplicate.getLastRowNum (); rowNoDuplicate++) {
                String noreg = noDuplicate.getRow (rowNoDuplicate).getCell (29).getStringCellValue ();
                String kelamin = registerData.get (noreg);
                if (kelamin != null) {
                    noDuplicate.getRow (rowNoDuplicate).createCell (31).setCellValue (kelamin);
                }
            }

            // Sort any value it contains
            List<Map.Entry<String, Integer>> entriesDoctor = new ArrayList<> (pivotJumlahTindakan.entrySet ());
            entriesDoctor.sort (Map.Entry.comparingByKey ());
            pivotJumlahTindakan = new LinkedHashMap<> ();
            for (Map.Entry<String, Integer> entry : entriesDoctor) {
                pivotJumlahTindakan.put (entry.getKey (), entry.getValue ());
            }

            // Write pivot result to Pertindakan sheet, starting from row 6
            int rowNum = 1;
            for (Map.Entry<String, Integer> entry : pivotJumlahTindakan.entrySet ()) {
                Row row = Pertindakan.createRow (rowNum++);
                row.createCell (0).setCellValue (rowNum - 1);
                row.createCell (1).setCellValue (entry.getKey ());
                row.createCell (2).setCellValue (entry.getValue ());
            }

            // Calculate grand total
            int grandTotal = 0;
            for (Integer value : pivotJumlahTindakan.values ()) {
                grandTotal += value;
            }

            // Write grand total to Pertindakan sheet
            Row rowx = Pertindakan.createRow (rowNum);
            rowx.createCell (0).setCellValue ("");
            rowx.createCell (1).setCellValue ("Grand Total");
            rowx.createCell (2).setCellValue (grandTotal);

            System.out.println ("00. " + BookPertindakanNew.getSheetAt (0).getSheetName () + ", "
                    + BookPertindakanNew.getSheetAt (1).getSheetName () + ", "
                    + BookPertindakanNew.getSheetAt (2).getSheetName () + " Complete");


//          buat sheet 2 Inst per cara bayar
            BookPertindakanNew.createSheet ();
            Sheet instCaraBayar = BookPertindakanNew.getSheetAt (3);
            BookPertindakanNew.setSheetName (3, "2. Inst per Cara Bayar");
            System.out.println ("01. " + BookPertindakanNew.getSheetAt (3).getSheetName () + " Start");
            instCaraBayar.createRow(0);
            instCaraBayar.createRow(1);

//          create place to store value
            Set<String> crByr = new TreeSet<>();
            Set<String> nickInstXTindakan = new TreeSet<>();
            Map<String, Map<String, Integer>> countMap = new HashMap<>(); // new count map

//          mapping value
            for (int row = 1; row <= noDuplicate.getLastRowNum(); row++) {
                String cellcrByr = noDuplicate.getRow(row).getCell(8).getStringCellValue();
                String cellNickInst = noDuplicate.getRow(row).getCell(24).getStringCellValue();

                crByr.add(cellcrByr);
                nickInstXTindakan.add(cellNickInst);

                // increment count in countMap
                if (!countMap.containsKey(cellNickInst)) {
                    countMap.put(cellNickInst, new HashMap<>());
                }
                Map<String, Integer> crBayarCountMap = countMap.get(cellNickInst);
                if (!crBayarCountMap.containsKey(cellcrByr)) {
                    crBayarCountMap.put(cellcrByr, 1);
                } else {
                    crBayarCountMap.put(cellcrByr, crBayarCountMap.get(cellcrByr) + 1);
                }
            }

//          writing cell
            instCaraBayar.createRow(0).createCell(0).setCellValue("Tanggal");
            int rowStart = 1;
            for (String konten : nickInstXTindakan) {
                instCaraBayar.getRow(0).createCell(rowStart).setCellValue(konten);
                rowStart++;
            }

//          filling row
            rowStart = 1;
            int lastCol = nickInstXTindakan.size() + 1;
            for (String konten : crByr) {
                int colStart = 1;
                instCaraBayar.createRow(rowStart).createCell(0).setCellValue(konten);
                int total = 0;
                for (String item : nickInstXTindakan) {
                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
                        int count = countMap.get(item).get(konten);
                        instCaraBayar.getRow(rowStart).createCell(colStart++).setCellValue(count);
                        total += count;
                    } else {
                        instCaraBayar.getRow(rowStart).createCell(colStart++).setCellValue(0);
                    }
                }
                instCaraBayar.getRow(rowStart).createCell(lastCol).setCellValue(total); // add row total
                rowStart++;
            }

//          add grand total to last row
            instCaraBayar.createRow(rowStart);
            int lastCell = instCaraBayar.getRow (0).getLastCellNum ();
            instCaraBayar.getRow (0).createCell (lastCell).setCellValue ("Grand Total");
            instCaraBayar.getRow(rowStart).createCell(0).setCellValue("Grand Total");
            int colStart = 1;
            for (String item : nickInstXTindakan) {
                int total = 0;
                for (String konten : crByr) {
                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
                        total += countMap.get(item).get(konten);
                    }
                }
                instCaraBayar.getRow(rowStart).createCell(colStart++).setCellValue(total);
            }
            instCaraBayar.getRow(rowStart).createCell(lastCol).setCellValue(noDuplicate.getLastRowNum()); // add total number of rows


            System.out.println ("01. " + BookPertindakanNew.getSheetAt (3).getSheetName () + " Completed");

//          buat sheet 3 kelamin per cara bayar
            BookPertindakanNew.createSheet ();
            Sheet klaminPerCrByr = BookPertindakanNew.getSheetAt (4);
            BookPertindakanNew.setSheetName (4, "3. Kelamin per Cara Bayar");
            System.out.println ("02. " + BookPertindakanNew.getSheetAt (4).getSheetName () + " Start");

            Map<String, Integer> klmnCrByrCount = new TreeMap<> ();
            for (int row = 1; row <= noDuplicate.getLastRowNum (); row++) {
                String caraByr = noDuplicate.getRow (row).getCell (8).getStringCellValue ();
                String kelamin = noDuplicate.getRow (row).getCell (31).getStringCellValue ();
                String crByrxTindakan = caraByr + "T.T" + kelamin; // I use T.T because i got no idea dot coma dash etc are used
                klmnCrByrCount.put (crByrxTindakan, klmnCrByrCount.getOrDefault (crByrxTindakan, 0) + 1);
            }

            klaminPerCrByr.createRow (0).createCell (0).setCellValue ("Jenis Cara Bayar");
            klaminPerCrByr.getRow (0).createCell (1).setCellValue ("Kelamin ");
            klaminPerCrByr.getRow (0).createCell (2).setCellValue ("Jumlah");

            int PsnCrByrrow = 0;
            int PsnCrByrSum = 0;
            for (Map.Entry<String, Integer> entry : klmnCrByrCount.entrySet ()) {
                PsnCrByrrow++;
                klaminPerCrByr.createRow (PsnCrByrrow);
                String[] splitValue = entry.getKey ().split("T.T");
                String crByrSplit = splitValue[0];
                String tnd = splitValue[1];
                klaminPerCrByr.getRow(PsnCrByrrow).createCell(0).setCellValue(crByrSplit);
                klaminPerCrByr.getRow(PsnCrByrrow).createCell(1).setCellValue(tnd);
                klaminPerCrByr.getRow(PsnCrByrrow).createCell(2).setCellValue (entry.getValue ());
                PsnCrByrSum += entry.getValue ();
            }
            int PsnCrByrLastRow = klaminPerCrByr.getLastRowNum () + 1;
            klaminPerCrByr.createRow (PsnCrByrLastRow).createCell (0).setCellValue ("Grand Total");
            klaminPerCrByr.getRow (PsnCrByrLastRow).createCell (2).setCellValue (PsnCrByrSum);

            System.out.println ("02. " + BookPertindakanNew.getSheetAt (4).getSheetName () + " Completed");

//          buat sheet 4 per Sub Instalasi
            BookPertindakanNew.createSheet ();
            Sheet perSubInstalasi = BookPertindakanNew.getSheetAt (5);
            BookPertindakanNew.setSheetName (5, "4. Laporan Per sub");
            System.out.println ("03. " + BookPertindakanNew.getSheetAt (5).getSheetName () + " Start");

            Map<String, Integer> subInstalasiCount = new TreeMap<> ();
            for (int row = 1; row <= noDuplicate.getLastRowNum (); row++) {
                String instalasi = noDuplicate.getRow (row).getCell (24).getStringCellValue ();
                String subInstalasi = noDuplicate.getRow (row).getCell (28).getStringCellValue ();
                String instalasiXSubInstalasi = instalasi + "~~~" + subInstalasi; // I use T.T because i got no idea dot coma dash etc are used
                subInstalasiCount.put (instalasiXSubInstalasi, subInstalasiCount.getOrDefault (instalasiXSubInstalasi, 0) + 1);
            }

            perSubInstalasi.createRow (0).createCell (0).setCellValue ("Instalasi");
            perSubInstalasi.getRow (0).createCell (1).setCellValue ("Sub Instalasi");
            perSubInstalasi.getRow (0).createCell (2).setCellValue ("Jumlah");

            int subInstrow = 0;
            int subInstSum = 0;
            for (Map.Entry<String, Integer> entry : subInstalasiCount.entrySet ()) {
                subInstrow++;
                perSubInstalasi.createRow (subInstrow);
                String[] splitValue = entry.getKey().split("~~~");
                String inst = splitValue[0];
                String subInst = splitValue[1];
                perSubInstalasi.getRow(subInstrow).createCell(0).setCellValue(inst);
                perSubInstalasi.getRow(subInstrow).createCell(1).setCellValue(subInst);
                perSubInstalasi.getRow(subInstrow).createCell(2).setCellValue(entry.getValue());
                subInstSum += entry.getValue ();
            }
            int subInstLastRow = perSubInstalasi.getLastRowNum () + 1;
            perSubInstalasi.createRow (subInstLastRow).createCell (0).setCellValue ("Grand Total");
            perSubInstalasi.getRow (subInstLastRow).createCell (2).setCellValue (subInstSum);

            System.out.println ("03. " + BookPertindakanNew.getSheetAt (5).getSheetName () + " Completed");


//          buat sheet 4 kelamin per instalasi
            BookPertindakanNew.createSheet ();
            Sheet klaminPerInst = BookPertindakanNew.getSheetAt (6);
            BookPertindakanNew.setSheetName (6, "5. Kelamin per Cara Bayar");
            System.out.println ("04. " + BookPertindakanNew.getSheetAt (6).getSheetName () + " Start");

            Map<String, Integer> klmnInstCount = new TreeMap<> ();
            for (int row = 1; row <= noDuplicate.getLastRowNum (); row++) {
                String instalasi = noDuplicate.getRow (row).getCell (24).getStringCellValue ();
                String kelamin = noDuplicate.getRow (row).getCell (31).getStringCellValue ();
                String instxTindakan = instalasi + "~~~" + kelamin; // I use T.T because i got no idea dot coma dash etc are used
                klmnInstCount.put (instxTindakan, klmnInstCount.getOrDefault (instxTindakan, 0) + 1);
            }

            klaminPerInst.createRow (0).createCell (0).setCellValue ("Instalasi");
            klaminPerInst.getRow (0).createCell (1).setCellValue ("Kelamin ");
            klaminPerInst.getRow (0).createCell (2).setCellValue ("Jumlah");

            int instlasirow = 0;
            int instalasiSum = 0;
            for (Map.Entry<String, Integer> entry : klmnInstCount.entrySet ()) {
                instlasirow++;
                klaminPerInst.createRow (instlasirow);
                String[] splitValue = entry.getKey ().split("~~~");
                String crByrSplit = splitValue[0];
                String tnd = splitValue[1];
                klaminPerInst.getRow(instlasirow).createCell(0).setCellValue(crByrSplit);
                klaminPerInst.getRow(instlasirow).createCell(1).setCellValue(tnd);
                klaminPerInst.getRow(instlasirow).createCell(2).setCellValue (entry.getValue ());
                instalasiSum += entry.getValue ();
            }
            int instLastRow = klaminPerInst.getLastRowNum () + 1;
            klaminPerInst.createRow (instLastRow).createCell (0).setCellValue ("Grand Total");
            klaminPerInst.getRow (instLastRow).createCell (2).setCellValue (instalasiSum);

            System.out.println ("04. " + BookPertindakanNew.getSheetAt (6).getSheetName () + " Completed");

//          buat sheet 6 Sub Inst per Hari
            BookPertindakanNew.createSheet ();
            Sheet subInstHari = BookPertindakanNew.getSheetAt (7);
            BookPertindakanNew.setSheetName (7, "6. Sub Inst per Hari");
            System.out.println ("05. " + BookPertindakanNew.getSheetAt (7).getSheetName () + " Start");
            subInstHari.createRow(0);
            subInstHari.createRow(1);

//          create place to store value
            nickInstXTindakan = new TreeSet<>();
            Set<String> tanggal = new TreeSet<>();
            countMap = new HashMap<>(); // new count map

//          mapping value
            for (int row = 1; row <= noDuplicate.getLastRowNum(); row++) {
                String cellNickInst = noDuplicate.getRow(row).getCell(24).getStringCellValue();
                String cellSubInst = noDuplicate.getRow(row).getCell(28).getStringCellValue();
                String instxTindakan = cellNickInst + "~~~" + cellSubInst; // I use T.T because i got no idea dot coma dash etc are used
                String cellTanggal = noDuplicate.getRow(row).getCell(9).getStringCellValue().substring (0, 10);

                tanggal.add(instxTindakan);//row
                nickInstXTindakan.add(cellTanggal);//cell

                // increment count in countMap
                if (!countMap.containsKey(cellTanggal)) {
                    countMap.put(cellTanggal, new HashMap<>());
                }
                Map<String, Integer> crBayarCountMap = countMap.get(cellTanggal);
                if (!crBayarCountMap.containsKey(instxTindakan)) {
                    crBayarCountMap.put(instxTindakan, 1);
                } else {
                    crBayarCountMap.put(instxTindakan, crBayarCountMap.get(instxTindakan) + 1);
                }
            }

//          writing cell
            subInstHari.createRow(0).createCell(0).setCellValue("Instalasi");
            subInstHari.getRow (0).createCell(1).setCellValue("Sub Instalasi");
            rowStart = 2;
            for (String konten : nickInstXTindakan) {
                subInstHari.getRow(0).createCell(rowStart).setCellValue(konten);
                rowStart++;
            }

//          filling row
            rowStart = 1;
            lastCol = nickInstXTindakan.size() + 2;
            for (String konten : tanggal) {
                String[] splitValue = konten.split("~~~");
                String instalasi = splitValue[0];
                String subInstalasi = splitValue[1];
                subInstHari.createRow(rowStart).createCell(0).setCellValue(instalasi);
                subInstHari.getRow(rowStart).createCell(1).setCellValue(subInstalasi);

                colStart = 2;
                int total = 0;
                for (String item : nickInstXTindakan) {
                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
                        int count = countMap.get(item).get(konten);
                        subInstHari.getRow(rowStart).createCell(colStart++).setCellValue(count);
                        total += count;
                    } else {
                        subInstHari.getRow(rowStart).createCell(colStart++).setCellValue(0);
                    }
                }
                subInstHari.getRow(rowStart).createCell(lastCol).setCellValue(total); // add row total
                rowStart++;
            }

//          add grand total to last row
            subInstHari.createRow(rowStart);
            lastCell = subInstHari.getRow (0).getLastCellNum ();
            subInstHari.getRow (0).createCell (lastCell).setCellValue ("Grand Total");
            subInstHari.getRow(rowStart).createCell(0).setCellValue("Grand Total");
            colStart = 2;
            for (String item : nickInstXTindakan) {
                int total = 0;
                for (String konten : tanggal) {
                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
                        total += countMap.get(item).get(konten);
                    }
                }
                subInstHari.getRow(rowStart).createCell(colStart++).setCellValue(total);
            }
            subInstHari.getRow(rowStart).createCell(lastCol).setCellValue(noDuplicate.getLastRowNum()); // add total number of rows


            System.out.println ("05. " + BookPertindakanNew.getSheetAt (7).getSheetName () + " Completed");



//          buat sheet 7 Pasien per Hari
            BookPertindakanNew.createSheet ();
            Sheet instalasiCrbyrHari = BookPertindakanNew.getSheetAt (8);
            BookPertindakanNew.setSheetName (8, "7. Jumlah Pasien per Hari");
            System.out.println ("06. " + BookPertindakanNew.getSheetAt (8).getSheetName () + " Start");
            instalasiCrbyrHari.createRow(0);
            instalasiCrbyrHari.createRow(1);

//          create place to store value
            tanggal = new TreeSet<>();
            nickInstXTindakan = new TreeSet<>();
            countMap = new HashMap<>(); // new count map

//          mapping value
            for (int row = 1; row <= noDuplicate.getLastRowNum(); row++) {
                String cellTanggal = noDuplicate.getRow(row).getCell(9).getStringCellValue().substring (0, 10);

                String cellNickInst = noDuplicate.getRow(row).getCell(24).getStringCellValue();
                String cellcrByr = noDuplicate.getRow(row).getCell(8).getStringCellValue();
                String instxTindakan = cellNickInst + "~~~" + cellcrByr; // I use T.T because i got no idea dot coma dash etc are used


                tanggal.add(cellTanggal);//row
                nickInstXTindakan.add(instxTindakan);//cell

                // increment count in countMap
                if (!countMap.containsKey(instxTindakan)) {
                    countMap.put(instxTindakan, new HashMap<>());
                }
                Map<String, Integer> crBayarCountMap = countMap.get(instxTindakan);
                if (!crBayarCountMap.containsKey(cellTanggal)) {
                    crBayarCountMap.put(cellTanggal, 1);
                } else {
                    crBayarCountMap.put(cellTanggal, crBayarCountMap.get(cellTanggal) + 1);
                }
            }

//          writing cell
            instalasiCrbyrHari.createRow(0).createCell(0).setCellValue("Tanggal");
            rowStart = 1;
            for (String konten : nickInstXTindakan) {
                String[] splitValue = konten.split("~~~");
                String instalasi = splitValue[0];
                String subInstalasi = splitValue[1];
                instalasiCrbyrHari.getRow(0).createCell(rowStart).setCellValue(instalasi);
                instalasiCrbyrHari.getRow(1).createCell(rowStart).setCellValue(subInstalasi);
                rowStart++;
            }

//          filling row
            rowStart = 2;
            lastCol = nickInstXTindakan.size() + 1;
            for (String konten : tanggal) {
                instalasiCrbyrHari.createRow(rowStart).createCell(0).setCellValue(konten);
                colStart = 1;
                int total = 0;
                for (String item : nickInstXTindakan) {
                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
                        int count = countMap.get(item).get(konten);
                        instalasiCrbyrHari.getRow(rowStart).createCell(colStart++).setCellValue(count);
                        total += count;
                    } else {
                        instalasiCrbyrHari.getRow(rowStart).createCell(colStart++).setCellValue(0);
                    }
                }
                instalasiCrbyrHari.getRow(rowStart).createCell(lastCol).setCellValue(total); // add row total
                rowStart++;
            }

//          add grand total to last row
            instalasiCrbyrHari.createRow(rowStart);
            lastCell = instalasiCrbyrHari.getRow (0).getLastCellNum ();
            instalasiCrbyrHari.getRow (0).createCell (lastCell).setCellValue ("Grand Total");
            instalasiCrbyrHari.getRow(rowStart).createCell(0).setCellValue("Grand Total");
            colStart = 1;
            for (String item : nickInstXTindakan) {
                int total = 0;
                for (String konten : tanggal) {
                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
                        total += countMap.get(item).get(konten);
                    }
                }
                instalasiCrbyrHari.getRow(rowStart).createCell(colStart++).setCellValue(total);
            }
            instalasiCrbyrHari.getRow(rowStart).createCell(lastCol).setCellValue(noDuplicate.getLastRowNum()); // add total number of rows

            System.out.println ("06. " + BookPertindakanNew.getSheetAt (8).getSheetName () + " Completed");



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

//          Sheet numbers 3 to 15
            for (int sheetNum = 2; sheetNum <= 8; sheetNum++) {
                Sheet currentSheet = BookPertindakanNew.getSheetAt(sheetNum);
                System.out.println (currentSheet.getSheetName ());
                for (int rightCell = 0; rightCell < currentSheet.getRow(0).getLastCellNum(); rightCell++) {
                    currentSheet.getRow(0).getCell(rightCell).setCellStyle(BorderCenterCellStyle);
                    currentSheet.autoSizeColumn(rightCell);
                    for (int downRow = 1; downRow <= currentSheet.getLastRowNum(); downRow++) {
                        if (currentSheet.getRow (downRow).getCell (rightCell)==null){
//                            System.out.println (downRow);
                            currentSheet.getRow (downRow).createCell (rightCell).setCellValue ("");
                        }
                        currentSheet.getRow(downRow).getCell(rightCell).setCellStyle(AllBorderCellStyle);
                    }
                }
            }

        } catch (Exception e) {
            e.printStackTrace ();
        }


        try {
            if (doneFinal) {
                outputStream = new FileOutputStream ("Done Lab " + localDate + ".xlsx");
            } else {
                outputStream = new FileOutputStream (fileNamePertindakanNew + " half done.xlsx");
            }
            BookPertindakanNew.write (outputStream);
        } catch (IOException e) {
            e.printStackTrace ();
        } finally {
            try {
                if (BookPertindakanNew != null) {
                    BookPertindakanNew.close ();
                }
                if (outputStream != null) {
                    outputStream.close ();
                }
            } catch (IOException e) {
                e.printStackTrace ();
            }
        }
    }

    private static int pertindakanNewRawLastRowNum(Sheet pertindakan_New_Raw) {
        return pertindakan_New_Raw.getLastRowNum ();
    }
}

