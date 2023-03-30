package LaporanIGD;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

public class LaporanIgd {

    public static void main(String[] args) throws IOException, AssertionError {
        new LaporanIgd ();
    }

    public LaporanIgd() throws IOException, AssertionError {
        String fileNameIgd = "23 02 igd register";
        String fileNameIrna = "23 02 irna register";
        InputStream LaporanIGD = new FileInputStream ("c:\\sat work\\test\\"+fileNameIgd+".xlsx");
        InputStream laporanIrna = new FileInputStream ("c:\\sat work\\test\\"+fileNameIrna+".xlsx");
        Workbook bookLaporanIGD = new XSSFWorkbook (LaporanIGD);
        Workbook bookLaporanIrna = new XSSFWorkbook (laporanIrna);
        Sheet register = bookLaporanIGD.getSheetAt (0);
        Sheet registerIrna = bookLaporanIrna.getSheetAt (0);

        //          Make Styling
        CellStyle AllBorderCellStyle = bookLaporanIGD.createCellStyle ();
        AllBorderCellStyle.setBorderBottom (BorderStyle.THIN);
        AllBorderCellStyle.setBottomBorderColor (IndexedColors.BLACK.getIndex ());
        AllBorderCellStyle.setBorderLeft (BorderStyle.THIN);
        AllBorderCellStyle.setLeftBorderColor (IndexedColors.BLACK.getIndex ());
        AllBorderCellStyle.setBorderRight (BorderStyle.THIN);
        AllBorderCellStyle.setRightBorderColor (IndexedColors.BLACK.getIndex ());
        AllBorderCellStyle.setBorderTop (BorderStyle.THIN);
        AllBorderCellStyle.setTopBorderColor (IndexedColors.BLACK.getIndex ());
        CellStyle BorderCenterCellStyle = bookLaporanIGD.createCellStyle ();
        BorderCenterCellStyle.setAlignment (HorizontalAlignment.CENTER);
        BorderCenterCellStyle.setBorderBottom (BorderStyle.THIN);
        BorderCenterCellStyle.setBottomBorderColor (IndexedColors.BLACK.getIndex ());
        BorderCenterCellStyle.setBorderLeft (BorderStyle.THIN);
        BorderCenterCellStyle.setLeftBorderColor (IndexedColors.BLACK.getIndex ());
        BorderCenterCellStyle.setBorderRight (BorderStyle.THIN);
        BorderCenterCellStyle.setRightBorderColor (IndexedColors.BLACK.getIndex ());
        BorderCenterCellStyle.setBorderTop (BorderStyle.THIN);
        BorderCenterCellStyle.setTopBorderColor (IndexedColors.BLACK.getIndex ());

        System.out.println ("01. Start doing " + bookLaporanIGD.getSheetName (0));

        String Year = register.getRow (2).getCell (5)
                .getStringCellValue ().substring (8, 10);
        String Month = register.getRow (2).getCell (5)
                .getStringCellValue ().substring (3, 5);
        int registerLastCell = register.getRow (0).getLastCellNum ();

        register.getRow (0).createCell (registerLastCell).setCellValue ("NOREG");
        for (int row = 1; row <= register.getLastRowNum (); row++) {
            StringBuilder sb = new StringBuilder ();
            for (int col = 0; col < 5; col++) {
                sb.append (register.getRow (row).getCell (col).getStringCellValue ());
            }
            String concatenated = sb.toString ();
            register.getRow (row).createCell (registerLastCell).setCellValue (concatenated);

            for (int cell =0;cell<register.getRow (0).getLastCellNum ();cell++) {
                if (register.getRow (row).getCell (cell) == null) {
                    register.getRow (row).createCell (cell).setCellValue ("");
                }
            }
        }

        //~~~~~~
        bookLaporanIGD.createSheet ("Sorted");
        Sheet sorted = bookLaporanIGD.getSheetAt (1);


        sorted.createRow (0).createCell (0).setCellValue ("NOREG");
        Integer[] SelectedArray = {5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,29,30,31,32,33,34,35,36,
                37,38,39,40,41,42,43,44,55,60,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83
        };
        Set<String> sortedNoreg = new TreeSet<> ();
        Row getRowRegister = register.getRow (0);
        Row headerRow = sorted.createRow(0);
        headerRow.createCell(0).setCellValue("NOREG");
        for (int i = 0; i < SelectedArray.length; i++) {
            headerRow.createCell(i+1).setCellValue(getRowRegister.getCell(SelectedArray[i]).getStringCellValue());
        }

        int sortedRow = 1;
        for (int row=1;row<=register.getLastRowNum ();row++){
            getRowRegister = register.getRow (row);
            if (getRowRegister.getCell (67).getStringCellValue ().equals ("Diagnosa Utama")){
                sorted.createRow (sortedRow);
                sorted.getRow (sortedRow).createCell (0).setCellValue (getRowRegister.getCell (registerLastCell).getStringCellValue ());
                for (int i = 0; i < SelectedArray.length; i++) {
                    if (getRowRegister.getCell(SelectedArray[i]) == null) {
                        sorted.getRow(sortedRow).createCell(i+1).setCellValue("");
                    } else if (getRowRegister.getCell(SelectedArray[i]).getCellType() == CellType.STRING) {
                        if (i == 0 || i==3) {
                            sorted.getRow(sortedRow).createCell(i+1).setCellValue(getRowRegister.getCell(SelectedArray[i]).getStringCellValue().substring(0, 10));
                        } else {
                            sorted.getRow(sortedRow).createCell(i+1).setCellValue(getRowRegister.getCell(SelectedArray[i]).getStringCellValue());
                        }
                    } else if (getRowRegister.getCell(SelectedArray[i]).getCellType() == CellType.NUMERIC) {
                        sorted.getRow(sortedRow).createCell(i+1).setCellValue(getRowRegister.getCell(SelectedArray[i]).getNumericCellValue());
                    }

                }
                sortedNoreg.add (getRowRegister.getCell (registerLastCell).getStringCellValue ());
                sortedRow++;
            }
        }

        int lastSortedRow= sorted.getLastRowNum ()+1;
        for (int row=1;row<=register.getLastRowNum ();row++) {
            getRowRegister = register.getRow (row);
            if (!sortedNoreg.contains (getRowRegister.getCell (registerLastCell).getStringCellValue ())){
                sorted.createRow (lastSortedRow);
                sorted.getRow (lastSortedRow).createCell (0).setCellValue (getRowRegister.getCell (registerLastCell).getStringCellValue ());
                for (int i = 0; i < SelectedArray.length; i++) {
                    if (getRowRegister.getCell(SelectedArray[i]) == null) {
                        sorted.getRow(lastSortedRow).createCell(i+1).setCellValue("");
                    } else if (getRowRegister.getCell(SelectedArray[i]).getCellType() == CellType.STRING) {
                        if (i == 0 || i==3) {
                            sorted.getRow(lastSortedRow).createCell(i+1).setCellValue(getRowRegister.getCell(SelectedArray[i]).getStringCellValue().substring(0, 10));
                        } else {
                            sorted.getRow(lastSortedRow).createCell(i+1).setCellValue(getRowRegister.getCell(SelectedArray[i]).getStringCellValue());
                        }
                    } else if (getRowRegister.getCell(SelectedArray[i]).getCellType() == CellType.NUMERIC) {
                        sorted.getRow(lastSortedRow).createCell(i+1).setCellValue(getRowRegister.getCell(SelectedArray[i]).getNumericCellValue());
                    }
                }
                sortedNoreg.add (getRowRegister.getCell (registerLastCell).getStringCellValue ());
                lastSortedRow++;
            }
        }
        System.out.println ("01. "+bookLaporanIGD.getSheetName (0)+" is done");


        bookLaporanIGD.createSheet ("2.KELAMIN PER TANGGAL");
        Sheet klaminPerTanggal = bookLaporanIGD.getSheetAt (2);
        System.out.println ("02. Start doing "+bookLaporanIGD.getSheetName (2));

        Set<String> tanggalRegist = new TreeSet<> ();
        Set<String> kelamin = new TreeSet<> ();
        Map<String, Map<String, Integer>> countMap = new HashMap<>(); // new count map
        for (int row=1;row<=sorted.getLastRowNum ();row++) {
            String cellTanggalReg = sorted.getRow (row).getCell (1).getStringCellValue ();
            String cellKelamin = sorted.getRow (row).getCell (6).getStringCellValue ();

            tanggalRegist.add (cellTanggalReg);
            kelamin.add (cellKelamin);

            // increment count in countMap
            if (!countMap.containsKey(cellKelamin)) {
                countMap.put(cellKelamin, new HashMap<>());
            }
            Map<String, Integer> kelaminCountMap = countMap.get(cellKelamin);
            if (!kelaminCountMap.containsKey(cellTanggalReg)) {
                kelaminCountMap.put(cellTanggalReg, 1);
            } else {
                kelaminCountMap.put(cellTanggalReg, kelaminCountMap.get(cellTanggalReg) + 1);
            }
        }

        klaminPerTanggal.createRow (0).createCell (0).setCellValue ("Kelamin");
        int rowStart=1;
        for (String tgl:tanggalRegist) {
            klaminPerTanggal.getRow (0).createCell (rowStart++).setCellValue (tgl);
        }
        rowStart=1;
        for (String klmn:kelamin){
            int colStart=1;
            klaminPerTanggal.createRow (rowStart).createCell (0).setCellValue (klmn);
            for (String tgl:tanggalRegist) {
                if (countMap.containsKey(klmn) && countMap.get(klmn).containsKey(tgl)) {
                    klaminPerTanggal.getRow (rowStart).createCell (colStart++).setCellValue(countMap.get(klmn).get(tgl));
                } else {
                    klaminPerTanggal.getRow (rowStart).createCell (colStart++).setCellValue(0);
                }
            }
            rowStart++;
        }
        System.out.println ("02. "+bookLaporanIGD.getSheetName (2)+" is done");



        bookLaporanIGD.createSheet ("3.KONDISI AKHIR PER TANGGAL");
        Sheet akhirPerTanggal = bookLaporanIGD.getSheetAt (3);
        System.out.println ("03. Start doing "+bookLaporanIGD.getSheetName (3));
        Set<String> kondisiAkhirTree = new TreeSet<> ();
        tanggalRegist = new TreeSet<> ();
        countMap = new HashMap<>(); // new count map
        String kondisiKeluar;
        for (int row=1;row<=sorted.getLastRowNum ();row++) {
            boolean kondisiAkhir = sorted.getRow (row).getCell (36) == null;
            boolean caraKeluar = sorted.getRow (row).getCell (37) == null;
            if (kondisiAkhir && caraKeluar) {
                kondisiKeluar = "";
            } else if (kondisiAkhir) {
                kondisiKeluar= register.getRow (row).getCell (43).getStringCellValue ();
            } else {
                kondisiKeluar= register.getRow (row).getCell (42).getStringCellValue ();
            }
            String cellKondisiAkhir = kondisiKeluar;
            String cellTanggalReg = sorted.getRow (row).getCell (1).getStringCellValue ();

            tanggalRegist.add (cellTanggalReg);
            kondisiAkhirTree.add (cellKondisiAkhir);

            // increment count in countMap
            if (!countMap.containsKey(cellKondisiAkhir)) {
                countMap.put(cellKondisiAkhir, new HashMap<>());
            }
            Map<String, Integer> kelaminCountMap = countMap.get(cellKondisiAkhir);
            if (!kelaminCountMap.containsKey(cellTanggalReg)) {
                kelaminCountMap.put(cellTanggalReg, 1);
            } else {
                kelaminCountMap.put(cellTanggalReg, kelaminCountMap.get(cellTanggalReg) + 1);
            }
        }

        akhirPerTanggal.createRow (0).createCell (0).setCellValue ("Kondisi Keluar");
        rowStart=1;
        for (String tgl:tanggalRegist) {
            akhirPerTanggal.getRow (0).createCell (rowStart++).setCellValue (tgl);
        }
        rowStart=1;
        for (String kondKlr:kondisiAkhirTree){
            int colStart=1;
            akhirPerTanggal.createRow (rowStart).createCell (0).setCellValue (kondKlr);
            for (String tgl:tanggalRegist) {
                if (countMap.containsKey(kondKlr) && countMap.get(kondKlr).containsKey(tgl)) {
                    akhirPerTanggal.getRow (rowStart).createCell (colStart++).setCellValue(countMap.get(kondKlr).get(tgl));
                } else {
                    akhirPerTanggal.getRow (rowStart).createCell (colStart++).setCellValue(0);
                }
            }
            rowStart++;
        }
        System.out.println ("03. "+bookLaporanIGD.getSheetName (3)+" is done");



        bookLaporanIGD.createSheet ("4.CARA BAYAR PER TANGGAL");
        Sheet crBayarPerTangal = bookLaporanIGD.getSheetAt (4);
        System.out.println ("04. Start doing "+bookLaporanIGD.getSheetName (4));
        Set<String> caraBayarTree = new TreeSet<> ();
        tanggalRegist = new TreeSet<> ();
        countMap = new HashMap<>(); // new count map
        for (int row=1;row<=sorted.getLastRowNum ();row++) {
            String cellCaraBayar = sorted.getRow (row).getCell (24).getStringCellValue ();
            String cellTanggalReg = sorted.getRow (row).getCell (1).getStringCellValue ();
            tanggalRegist.add (cellTanggalReg);
            caraBayarTree.add (cellCaraBayar);

            // increment count in countMap
            if (!countMap.containsKey(cellCaraBayar)) {
                countMap.put(cellCaraBayar, new HashMap<>());
            }
            Map<String, Integer> caraBayarCountMap = countMap.get(cellCaraBayar);
            if (!caraBayarCountMap.containsKey(cellTanggalReg)) {
                caraBayarCountMap.put(cellTanggalReg, 1);
            } else {
                caraBayarCountMap.put(cellTanggalReg, caraBayarCountMap.get(cellTanggalReg) + 1);
            }
        }

        crBayarPerTangal.createRow (0).createCell (0).setCellValue ("Kondisi Keluar");
        rowStart=1;
        for (String tgl:tanggalRegist) {
            crBayarPerTangal.getRow (0).createCell (rowStart++).setCellValue (tgl);
        }
        rowStart=1;
        for (String caraBayar:caraBayarTree){
            int colStart=1;
            crBayarPerTangal.createRow (rowStart).createCell (0).setCellValue (caraBayar);
            for (String tgl:tanggalRegist) {
                if (countMap.containsKey(caraBayar) && countMap.get(caraBayar).containsKey(tgl)) {
                    crBayarPerTangal.getRow (rowStart).createCell (colStart++).setCellValue(countMap.get(caraBayar).get(tgl));
                } else {
                    crBayarPerTangal.getRow (rowStart).createCell (colStart++).setCellValue(0);
                }
            }
            rowStart++;
        }
        System.out.println ("04. "+bookLaporanIGD.getSheetName (4)+" is done");



        bookLaporanIGD.createSheet ("5.PASIEN DOKTER TUGAS");
        Sheet pasienDokterTugas = bookLaporanIGD.getSheetAt (5);
        System.out.println ("05. Start doing "+bookLaporanIGD.getSheetName (5));

        Map<String, Integer> psnDktrTugas = new TreeMap<> ();
        for (int row = 1; row <= sorted.getLastRowNum (); row++) {
            String psnInstAsal = sorted.getRow (row).getCell (54).getStringCellValue ();
            psnDktrTugas.put (psnInstAsal, psnDktrTugas.getOrDefault (psnInstAsal, 0) + 1);
        }

        pasienDokterTugas.createRow (0).createCell (0).setCellValue ("Dokter Tugas");
        pasienDokterTugas.getRow (0).createCell (1).setCellValue ("Jumlah");

        int dokterTugasRow = 0;
        int dokterTugasSum = 0;
        for (Map.Entry<String, Integer> entry : psnDktrTugas.entrySet ()) {
            dokterTugasRow++;
            pasienDokterTugas.createRow (dokterTugasRow).createCell (0).setCellValue (entry.getKey ());
            pasienDokterTugas.getRow (dokterTugasRow).createCell (1).setCellValue (entry.getValue ());
            dokterTugasSum += entry.getValue ();
        }
        int PsnInstAsalLastRow = pasienDokterTugas.getLastRowNum () + 1;
        pasienDokterTugas.createRow (PsnInstAsalLastRow).createCell (0).setCellValue ("Grand Total");
        pasienDokterTugas.getRow (PsnInstAsalLastRow).createCell (1).setCellValue (dokterTugasSum);

        System.out.println ("05. "+bookLaporanIGD.getSheetName (5)+" is done");



        bookLaporanIGD.createSheet ("6.PASIEN DOKTER PER TANGGAL");
        Sheet pasienDokterPerTanggal = bookLaporanIGD.getSheetAt (6);
        System.out.println ("06. Start doing "+bookLaporanIGD.getSheetName (6));

        Set<String> psnDokterTugas = new TreeSet<> ();
        tanggalRegist = new TreeSet<> ();
        countMap = new HashMap<>(); // new count map
        for (int row=1;row<=sorted.getLastRowNum ();row++) {
            String cellDokterTugas = sorted.getRow (row).getCell (54).getStringCellValue ();
            String cellTanggalReg = sorted.getRow (row).getCell (1).getStringCellValue ();

            tanggalRegist.add (cellTanggalReg);
            psnDokterTugas.add (cellDokterTugas);

            // increment count in countMap
            if (!countMap.containsKey(cellDokterTugas)) {
                countMap.put(cellDokterTugas, new HashMap<>());
            }
            Map<String, Integer> drTugasCountMap = countMap.get(cellDokterTugas);
            if (!drTugasCountMap.containsKey(cellTanggalReg)) {
                drTugasCountMap.put(cellTanggalReg, 1);
            } else {
                drTugasCountMap.put(cellTanggalReg, drTugasCountMap.get(cellTanggalReg) + 1);
            }
        }

        pasienDokterPerTanggal.createRow (0).createCell (0).setCellValue ("Kondisi Keluar");
        rowStart=1;
        for (String tgl:tanggalRegist) {
            pasienDokterPerTanggal.getRow (0).createCell (rowStart++).setCellValue (tgl);
        }
        rowStart=1;
        for (String drTugas:psnDokterTugas){
            int colStart=1;
            pasienDokterPerTanggal.createRow (rowStart).createCell (0).setCellValue (drTugas);
            for (String tgl:tanggalRegist) {
                if (countMap.containsKey(drTugas) && countMap.get(drTugas).containsKey(tgl)) {
                    pasienDokterPerTanggal.getRow (rowStart).createCell (colStart++).setCellValue(countMap.get(drTugas).get(tgl));
                } else {
                    pasienDokterPerTanggal.getRow (rowStart).createCell (colStart++).setCellValue(0);
                }
            }
            rowStart++;
        }

        System.out.println ("06. "+bookLaporanIGD.getSheetName (6)+" is done");



        bookLaporanIGD.createSheet ("7.SUB INSTALASI PER CARA MASUK");
        Sheet subInstalasiCaraMasuk = bookLaporanIGD.getSheetAt (7);
        System.out.println ("07. Start doing "+bookLaporanIGD.getSheetName (7));

        Set<String> subInstalasi = new TreeSet<> ();
        Set<String> caraMasuk = new TreeSet<> ();
        countMap = new HashMap<>(); // new count map
        for (int row=1;row<=sorted.getLastRowNum ();row++) {
            String cellSubInst = sorted.getRow (row).getCell (30).getStringCellValue ();
            String cellCaraMasuk = sorted.getRow (row).getCell (32).getStringCellValue ();

            caraMasuk.add (cellCaraMasuk);
            subInstalasi.add (cellSubInst);

            // increment count in countMap
            if (!countMap.containsKey(cellSubInst)) {
                countMap.put(cellSubInst, new HashMap<>());
            }
            Map<String, Integer> drTugasCountMap = countMap.get(cellSubInst);
            if (!drTugasCountMap.containsKey(cellCaraMasuk)) {
                drTugasCountMap.put(cellCaraMasuk, 1);
            } else {
                drTugasCountMap.put(cellCaraMasuk, drTugasCountMap.get(cellCaraMasuk) + 1);
            }
        }

        subInstalasiCaraMasuk.createRow (0).createCell (0).setCellValue ("Sub Instalasi");
        rowStart=1;
        for (String crMsk:caraMasuk) {
            subInstalasiCaraMasuk.getRow (0).createCell (rowStart++).setCellValue (crMsk);
        }
        rowStart=1;
        for (String subInst:subInstalasi){
            int colStart=1;
            subInstalasiCaraMasuk.createRow (rowStart).createCell (0).setCellValue (subInst);
            for (String crMsk:caraMasuk) {
                if (countMap.containsKey(subInst) && countMap.get(subInst).containsKey(crMsk)) {
                    subInstalasiCaraMasuk.getRow (rowStart).createCell (colStart++).setCellValue(countMap.get(subInst).get(crMsk));
                } else {
                    subInstalasiCaraMasuk.getRow (rowStart).createCell (colStart++).setCellValue(0);
                }
            }
            rowStart++;
        }

        System.out.println ("07. "+bookLaporanIGD.getSheetName (7)+" is done");



        bookLaporanIGD.createSheet ("8.DIAGNOSA PER JENIS KELAMIN");
        Sheet diagPerKelamin = bookLaporanIGD.getSheetAt (8);
        System.out.println ("08. Start doing "+bookLaporanIGD.getSheetName (8));

        Set<String> diagnosa = new TreeSet<> ();
        Set<String> jenisKelamin = new TreeSet<> ();
        countMap = new HashMap<>(); // new count map
        for (int row=1;row<=sorted.getLastRowNum ();row++) {
            String celldiagnosa = sorted.getRow (row).getCell (35).getStringCellValue ();
            String cellJenisKelamin = sorted.getRow (row).getCell (6).getStringCellValue ();

            diagnosa.add (celldiagnosa);
            jenisKelamin.add (cellJenisKelamin);

            // increment count in countMap
            if (!countMap.containsKey(celldiagnosa)) {
                countMap.put(celldiagnosa, new HashMap<>());
            }
            Map<String, Integer> drTugasCountMap = countMap.get(celldiagnosa);
            if (!drTugasCountMap.containsKey(cellJenisKelamin)) {
                drTugasCountMap.put(cellJenisKelamin, 1);
            } else {
                drTugasCountMap.put(cellJenisKelamin, drTugasCountMap.get(cellJenisKelamin) + 1);
            }
        }

        diagPerKelamin.createRow (0).createCell (0).setCellValue ("Sub Instalasi");
        rowStart=1;
        for (String jnsKlamin:jenisKelamin) {
            diagPerKelamin.getRow (0).createCell (rowStart++).setCellValue (jnsKlamin);
        }
        rowStart=1;
        for (String subInst:diagnosa){
            int colStart=1;
            diagPerKelamin.createRow (rowStart).createCell (0).setCellValue (subInst);
            for (String jnsKlamin:jenisKelamin) {
                if (countMap.containsKey(subInst) && countMap.get(subInst).containsKey(jnsKlamin)) {
                    diagPerKelamin.getRow (rowStart).createCell (colStart++).setCellValue(countMap.get(subInst).get(jnsKlamin));
                } else {
                    diagPerKelamin.getRow (rowStart).createCell (colStart++).setCellValue(0);
                }
            }
            rowStart++;
        }

        System.out.println ("08. "+bookLaporanIGD.getSheetName (8)+" is done");



        bookLaporanIGD.createSheet ("9.PASIEN J-COVID");
        Sheet pasienJCovid = bookLaporanIGD.getSheetAt (9);
        System.out.println ("09. Start doing "+bookLaporanIGD.getSheetName (9));

        pasienJCovid.createRow (0);
        for (int cell = 0; cell < sorted.getRow(0).getLastCellNum(); cell++) {
            pasienJCovid.getRow (0).createCell (cell).setCellValue (sorted.getRow (0).getCell (cell).getStringCellValue ());
        }
        int covidRow = 1;
        for (int row = 1; row <= sorted.getLastRowNum(); row++) {
            if (sorted.getRow(row).getCell(23).getStringCellValue().equals("Jaminan COVID-19")) {
                Row sourceRow = sorted.getRow(row);
                Row targetRow = pasienJCovid.createRow(covidRow);
                for (int cell = 0; cell < sorted.getRow(0).getLastCellNum(); cell++) {
                    Cell sourceCell = sourceRow.getCell(cell);
                    Cell targetCell = targetRow.createCell(cell);
                    if (sourceCell.getCellType() == CellType.STRING) {
                        targetCell.setCellValue(sourceCell.getStringCellValue());
                    } else if (sourceCell.getCellType() == CellType.NUMERIC) {
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                    } else {
                        targetCell.setCellValue("");
                    }
                }
                covidRow++;
            }
        }

        System.out.println ("09. "+bookLaporanIGD.getSheetName (9)+" is done");



        bookLaporanIGD.createSheet ("10.PASIEN MAWAR");
        Sheet pasienMawar = bookLaporanIGD.getSheetAt (10);
        System.out.println ("10. Start doing "+bookLaporanIGD.getSheetName (10));

        pasienMawar.createRow (0).createCell (0).setCellValue ("NOREG");
        Set<String> sortedNoregIrna = new TreeSet<> ();
        Row getRowRegisterIrna = registerIrna.getRow (0);
        Row headerRowMawat = pasienMawar.createRow(0);
        headerRowMawat.createCell(0).setCellValue("NOREG");
        for (int i = 0; i < SelectedArray.length; i++) {
            headerRowMawat.createCell(i+1).setCellValue(getRowRegisterIrna.getCell(SelectedArray[i]).getStringCellValue());
        }

        sortedRow = 1;
        for (int row=1;row<=registerIrna.getLastRowNum ();row++){
            getRowRegister = registerIrna.getRow (row);
            String noreg = getRowRegister.getCell (0).getStringCellValue ()+
                    getRowRegister.getCell (1).getStringCellValue ()+
                    getRowRegister.getCell (2).getStringCellValue ()+
                    getRowRegister.getCell (3).getStringCellValue ()+
                    getRowRegister.getCell (4).getStringCellValue ();
//            if (getRowRegister.getCell (67).getStringCellValue ().equals ("Diagnosa Utama")){
            if (!sortedNoregIrna.contains (noreg)) {
                if (getRowRegister.getCell (34).getStringCellValue ().contains ("Mawar") || getRowRegister.getCell (36).getStringCellValue ().contains ("Mawar")) {
                    pasienMawar.createRow (sortedRow);
                    pasienMawar.getRow (sortedRow).createCell (0).setCellValue (
                            getRowRegister.getCell (0).getStringCellValue () +
                                    getRowRegister.getCell (1).getStringCellValue () +
                                    getRowRegister.getCell (2).getStringCellValue () +
                                    getRowRegister.getCell (3).getStringCellValue () +
                                    getRowRegister.getCell (4).getStringCellValue ()
                    );
                    for (int i = 0; i < SelectedArray.length; i++) {
                        if (getRowRegister.getCell (SelectedArray[i]) == null) {
                            pasienMawar.getRow (sortedRow).createCell (i + 1).setCellValue ("");
                        } else if (getRowRegister.getCell (SelectedArray[i]).getCellType () == CellType.STRING) {
                            if (i == 0 || i == 3) {
                                pasienMawar.getRow (sortedRow).createCell (i + 1).setCellValue (getRowRegister.getCell (SelectedArray[i]).getStringCellValue ().substring (0, 10));
                            } else {
                                pasienMawar.getRow (sortedRow).createCell (i + 1).setCellValue (getRowRegister.getCell (SelectedArray[i]).getStringCellValue ());
                            }
                        } else if (getRowRegister.getCell (SelectedArray[i]).getCellType () == CellType.NUMERIC) {
                            pasienMawar.getRow (sortedRow).createCell (i + 1).setCellValue (getRowRegister.getCell (SelectedArray[i]).getNumericCellValue ());
                        }

                    }
                    sortedNoregIrna.add (noreg);
                    sortedRow++;
                }
            }
        }

        // Define the sheets you want to apply the styles to
        Sheet[] sheets = {sorted, klaminPerTanggal, akhirPerTanggal, crBayarPerTangal, pasienDokterTugas,
                pasienDokterPerTanggal, subInstalasiCaraMasuk, diagPerKelamin, pasienJCovid, pasienMawar};

        // Loop through the sheets
        for (Sheet sheet : sheets) {
            // Loop through the cells in the first row
            for (int rightCell = 0; rightCell < sheet.getRow(0).getLastCellNum(); rightCell++) {
                sheet.getRow(0).getCell(rightCell).setCellStyle(BorderCenterCellStyle);
                sheet.autoSizeColumn(rightCell);
                // Loop through the cells in the remaining rows
                for (int downRow = 1; downRow <= sheet.getLastRowNum(); downRow++) {
                    sheet.getRow(downRow).getCell(rightCell).setCellStyle(AllBorderCellStyle);
                }
            }
        }


        System.out.println ("10. "+bookLaporanIGD.getSheetName (10)+" is done");
        bookLaporanIGD.removeSheetAt (0);




        FileOutputStream IGDHalfDone = new FileOutputStream (Year + " " + Month + " IGD Half Done.xlsx");
        bookLaporanIGD.write (IGDHalfDone);

    }

}

