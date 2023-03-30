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
        String fileName = "23 02 igd register";
        InputStream LaporanIGD = new FileInputStream ("c:\\sat work\\test\\"+fileName+".xlsx");
        Workbook bookLaporanIGD = new XSSFWorkbook (LaporanIGD);
        Sheet register = bookLaporanIGD.getSheetAt (0);

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

        System.out.println ("07. "+bookLaporanIGD.getSheetName (7)+" is done");



        bookLaporanIGD.createSheet ("8.DIAGNOSA PER JENIS KELAMIN");
        Sheet diagPerKelamin = bookLaporanIGD.getSheetAt (8);
        System.out.println ("08. Start doing "+bookLaporanIGD.getSheetName (8));

        System.out.println ("08. "+bookLaporanIGD.getSheetName (8)+" is done");



        bookLaporanIGD.createSheet ("9.PASIEN J-COVID");
        Sheet pasienJCovid = bookLaporanIGD.getSheetAt (9);
        System.out.println ("09. Start doing "+bookLaporanIGD.getSheetName (9));

        System.out.println ("09. "+bookLaporanIGD.getSheetName (9)+" is done");



        bookLaporanIGD.createSheet ("10.PASIEN MAWAR");
        Sheet pasienMawar = bookLaporanIGD.getSheetAt (10);
        System.out.println ("10. Start doing "+bookLaporanIGD.getSheetName (10));

        System.out.println ("10. "+bookLaporanIGD.getSheetName (10)+" is done");



        FileOutputStream IGDHalfDone = new FileOutputStream (Year + " " + Month + " IGD Half Done.xlsx");
        bookLaporanIGD.write (IGDHalfDone);

    }

}

