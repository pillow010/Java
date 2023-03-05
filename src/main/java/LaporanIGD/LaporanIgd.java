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
        final int DIAGNOSIS_CELL = 67;
        final int DIAGNOSIS_CODE_CELL = registerLastCell + 2;

        register.getRow (0).createCell (registerLastCell).setCellValue ("NOREG");
        register.getRow (0).createCell (registerLastCell+1).setCellValue ("Kondisi Keluar");
        register.getRow (0).createCell (DIAGNOSIS_CODE_CELL).setCellValue ("DIAGNOSIS_CODE_CELL");
        for (int row = 1; row <= register.getLastRowNum (); row++) {
            StringBuilder sb = new StringBuilder ();
            for (int col = 0; col < 5; col++) {
                sb.append (register.getRow (row).getCell (col).getStringCellValue ());
            }
            String concatenated = sb.toString ();
            register.getRow (row).createCell (registerLastCell).setCellValue (concatenated);

            boolean kondisiAkhir = register.getRow (row).getCell (42) == null;
            boolean caraKeluar = register.getRow (row).getCell (43) == null;
            if (kondisiAkhir && caraKeluar) {
                register.getRow (row).createCell (registerLastCell + 1).setCellValue ("");
            } else if (kondisiAkhir) {
                register.getRow (row).createCell (registerLastCell + 1).setCellValue (
                        register.getRow (row).getCell (43).getStringCellValue ()
                );
            } else {
                register.getRow (row).createCell (registerLastCell + 1).setCellValue (
                        register.getRow (row).getCell (42).getStringCellValue ()
                );
            }


            String diagnosis = Objects.requireNonNullElse (
                    register.getRow (row).getCell (DIAGNOSIS_CELL).getStringCellValue (), "");
            register.getRow (row).createCell (DIAGNOSIS_CODE_CELL)
                    .setCellValue (diagnosis.equals ("Diagnosa Utama") ? "AA" : diagnosis);


            for (int cell =0;cell<register.getRow (0).getLastCellNum ();cell++) {
                if (register.getRow (row).getCell (cell) == null) {
                    register.getRow (row).createCell (cell).setCellValue ("");
                }
            }
        }
//            // Sort sheet based on registerLastCell+2
//            DataFormatter formatter = new DataFormatter();
//            List<Row> rows = StreamSupport.stream (register.spliterator (), false)
//                    .sorted (Comparator.comparingInt (row -> {
//                        String value = formatter.formatCellValue (row.getCell (DIAGNOSIS_CODE_CELL));
//                        return value.charAt (0);
//                    })).toList ();
//            rows.stream ().forEach (System.out::println);


        bookLaporanIGD.createSheet ("Sorted");
        Sheet sorted = bookLaporanIGD.getSheetAt (1);

        sorted.createRow (0).createCell (0).setCellValue ("NOREG");
        sorted.getRow (0).createCell (1).setCellValue ("TANGGAL");
        sorted.getRow (0).createCell (2).setCellValue ("RM");
        sorted.getRow (0).createCell (3).setCellValue ("NAMA");
        sorted.getRow (0).createCell (4).setCellValue ("KELAMIN");
        sorted.getRow (0).createCell (5).setCellValue ("JENIS CARA BAYAR");
        sorted.getRow (0).createCell (6).setCellValue ("SUB INSTALASI");
        sorted.getRow (0).createCell (7).setCellValue ("CARA MASUK");
        sorted.getRow (0).createCell (8).setCellValue ("DIAGNOSA");
        sorted.getRow (0).createCell (9).setCellValue ("KONDISI AKHIR");
        sorted.getRow (0).createCell (10).setCellValue ("DOKTER TUGAS");
        sorted.getRow (0).createCell (11).setCellValue ("DIAGNOSA");
        sorted.getRow (0).createCell (12).setCellValue ("KET JENIS DIAGNOSA");

        Set<String> sortedNoreg = new TreeSet<> ();
        int sortedRow = 1;
        for (int row=1;row<=register.getLastRowNum ();row++){
            Row getRowRegister = register.getRow (row);
            if (getRowRegister.getCell (67).getStringCellValue ().equals ("Diagnosa Utama")){
                sorted.createRow (sortedRow);
                sorted.getRow (sortedRow).createCell (0).setCellValue (getRowRegister.getCell (registerLastCell).getStringCellValue ());
                sorted.getRow (sortedRow).createCell (1).setCellValue (getRowRegister.getCell (5).getStringCellValue ().substring (0, 10));
                sorted.getRow (sortedRow).createCell (2).setCellValue (getRowRegister.getCell (6).getStringCellValue ());
                sorted.getRow (sortedRow).createCell (3).setCellValue (getRowRegister.getCell (7).getStringCellValue ());
                sorted.getRow (sortedRow).createCell (4).setCellValue (getRowRegister.getCell (10).getStringCellValue ());
                sorted.getRow (sortedRow).createCell (5).setCellValue (getRowRegister.getCell (30).getStringCellValue ());
                sorted.getRow (sortedRow).createCell (6).setCellValue (getRowRegister.getCell (36).getStringCellValue ());
                sorted.getRow (sortedRow).createCell (7).setCellValue (getRowRegister.getCell (38).getStringCellValue ());
                sorted.getRow (sortedRow).createCell (8).setCellValue (getRowRegister.getCell (41).getStringCellValue ());
                sorted.getRow (sortedRow).createCell (9).setCellValue (getRowRegister.getCell (registerLastCell+1).getStringCellValue ());
                sorted.getRow (sortedRow).createCell (10).setCellValue (getRowRegister.getCell (78).getStringCellValue ());
                sorted.getRow (sortedRow).createCell (11).setCellValue (getRowRegister.getCell (41).getStringCellValue ());
                sorted.getRow (sortedRow).createCell (12).setCellValue (getRowRegister.getCell (67).getStringCellValue ());
                sortedNoreg.add (getRowRegister.getCell (registerLastCell).getStringCellValue ());
                sortedRow++;
            }
        }

        int lastSortedRow= sorted.getLastRowNum ()+1;
        for (int row=1;row<=register.getLastRowNum ();row++) {
            Row getRowRegister = register.getRow (row);
            if (!sortedNoreg.contains (getRowRegister.getCell (registerLastCell).getStringCellValue ())){
                sorted.createRow (lastSortedRow);
                sorted.getRow (lastSortedRow).createCell (0).setCellValue (getRowRegister.getCell (registerLastCell).getStringCellValue ());
                sorted.getRow (lastSortedRow).createCell (1).setCellValue (getRowRegister.getCell (5).getStringCellValue ().substring (0, 10));
                sorted.getRow (lastSortedRow).createCell (2).setCellValue (getRowRegister.getCell (6).getStringCellValue ());
                sorted.getRow (lastSortedRow).createCell (3).setCellValue (getRowRegister.getCell (7).getStringCellValue ());
                sorted.getRow (lastSortedRow).createCell (4).setCellValue (getRowRegister.getCell (10).getStringCellValue ());
                sorted.getRow (lastSortedRow).createCell (5).setCellValue (getRowRegister.getCell (30).getStringCellValue ());
                sorted.getRow (lastSortedRow).createCell (6).setCellValue (getRowRegister.getCell (36).getStringCellValue ());
                sorted.getRow (lastSortedRow).createCell (7).setCellValue (getRowRegister.getCell (38).getStringCellValue ());
                sorted.getRow (lastSortedRow).createCell (8).setCellValue (getRowRegister.getCell (41).getStringCellValue ());
                sorted.getRow (lastSortedRow).createCell (9).setCellValue (getRowRegister.getCell (registerLastCell+1).getStringCellValue ());
                sorted.getRow (lastSortedRow).createCell (10).setCellValue (getRowRegister.getCell (78).getStringCellValue ());
                sorted.getRow (lastSortedRow).createCell (11).setCellValue (getRowRegister.getCell (41).getStringCellValue ());
                sorted.getRow (lastSortedRow).createCell (12).setCellValue (getRowRegister.getCell (67).getStringCellValue ());
                sortedNoreg.add (getRowRegister.getCell (registerLastCell).getStringCellValue ());
                lastSortedRow++;
            }
        }
//        // Add sorted rows back to the sheet
//        for (Row sourceRow : rows) {
//            Row targetRow = sorted.createRow(sourceRow.getRowNum());
//
//            // Copy cell values from source row to target row
//            for (Cell sourceCell : sourceRow) {
//                Cell targetCell = targetRow.createCell(sourceCell.getColumnIndex());
//                if (sourceCell.getCellType ()==CellType.STRING){
//                    targetCell.setCellValue (sourceCell.getStringCellValue ());
//                } else {
//                    targetCell.setCellValue (sourceCell.getNumericCellValue ());
//                }
//
//            }
//        }




        System.out.println ("01. "+bookLaporanIGD.getSheetName (0)+" is done");


        bookLaporanIGD.createSheet ("without Duplicate");
        Sheet noDuplicate = bookLaporanIGD.getSheetAt (2);
        System.out.println ("02. Start doing" + bookLaporanIGD.getSheetName (2));


        noDuplicate.createRow (0).createCell (0).setCellValue ("NOREG");
        noDuplicate.getRow (0).createCell (1).setCellValue ("TANGGAL");
        noDuplicate.getRow (0).createCell (2).setCellValue ("RM");
        noDuplicate.getRow (0).createCell (3).setCellValue ("NAMA");
        noDuplicate.getRow (0).createCell (4).setCellValue ("KELAMIN");
        noDuplicate.getRow (0).createCell (5).setCellValue ("JENIS CARA BAYAR");
        noDuplicate.getRow (0).createCell (6).setCellValue ("SUB INSTALASI");
        noDuplicate.getRow (0).createCell (7).setCellValue ("CARA MASUK");
        noDuplicate.getRow (0).createCell (8).setCellValue ("DIAGNOSA");
        noDuplicate.getRow (0).createCell (9).setCellValue ("KONDISI AKHIR");
        noDuplicate.getRow (0).createCell (10).setCellValue ("DOKTER TUGAS");


        Set<String> noreg = new TreeSet<> ();
        for (int row = 1; row <= register.getLastRowNum (); row++) {
                noreg.add (register.getRow (row).getCell (registerLastCell).getStringCellValue ());
        }

        int i = 0;
        for (String value : noreg) {
            noDuplicate.createRow (++i).createCell (0).setCellValue (value);
        }

//        for (int row =1;row<=noDuplicate.getLastRowNum ();row++) {
//            String target = noDuplicate.getRow (row).getCell (0).getStringCellValue ();
//            for (int rowSource = 1; rowSource <= register.getLastRowNum (); rowSource++) {
//                String source = register.getRow (rowSource).getCell (0).getStringCellValue ();
//                if (target.equals (source)){
//                    String	tanggal	= register.getRow (rowSource).getCell (5).getStringCellValue ().substring (0, 10);
//                    String	rm = register.getRow (rowSource).getCell (6).getStringCellValue ();
//                    String	nama = register.getRow (rowSource).getCell (7).getStringCellValue ();
//                    String	kelamin	= register.getRow (rowSource).getCell (10).getStringCellValue ();
//                    String	jenisCaraBayar	= register.getRow (rowSource).getCell (30).getStringCellValue ();
//                    String	subInstalasi	= register.getRow (rowSource).getCell (36).getStringCellValue ();
//                    String	caraMasuk	= register.getRow (rowSource).getCell (38).getStringCellValue ();
//                    String	diagnosa	= register.getRow (rowSource).getCell (41).getStringCellValue ();
//                    String	kondisiAkhir	= register.getRow (rowSource).getCell (registerLastCell+1).getStringCellValue ();
//                    String	dokterTugas	= register.getRow (rowSource).getCell (78).getStringCellValue ();
//                    noDuplicate.getRow (row).createCell (1).setCellValue (tanggal);
//                    noDuplicate.getRow (row).createCell (2).setCellValue (rm);
//                    noDuplicate.getRow (row).createCell (3).setCellValue (nama);
//                    noDuplicate.getRow (row).createCell (4).setCellValue (kelamin);
//                    noDuplicate.getRow (row).createCell (5).setCellValue (jenisCaraBayar);
//                    noDuplicate.getRow (row).createCell (6).setCellValue (subInstalasi);
//                    noDuplicate.getRow (row).createCell (7).setCellValue (caraMasuk);
//                    noDuplicate.getRow (row).createCell (8).setCellValue (diagnosa);
//                    noDuplicate.getRow (row).createCell (9).setCellValue (kondisiAkhir);
//                    noDuplicate.getRow (row).createCell (10).setCellValue (dokterTugas);
//                }
//            }
//        }
        Map<String, String[]> registerData = new HashMap<> ();
        for (int row = 1; row <= register.getLastRowNum(); row++) {
            String source = register.getRow(row).getCell(registerLastCell).getStringCellValue();
            String[] data = new String[10];
            data[0] = register.getRow(row).getCell(5).getStringCellValue().substring(0, 10);
            data[1] = register.getRow(row).getCell(6).getStringCellValue();
            data[2] = register.getRow(row).getCell(7).getStringCellValue();
            data[3] = register.getRow(row).getCell(10).getStringCellValue();
            data[4] = register.getRow(row).getCell(30).getStringCellValue();
            data[5] = register.getRow(row).getCell(36).getStringCellValue();
            data[6] = register.getRow(row).getCell(38).getStringCellValue();
            data[7] = register.getRow(row).getCell(41).getStringCellValue();
            data[8] = register.getRow(row).getCell(registerLastCell + 1).getStringCellValue();
            data[9] = register.getRow(row).getCell(78).getStringCellValue();
            registerData.put(source, data);
        }

        List<String[]> noDuplicateData = new ArrayList<> ();
        for (int row = 1; row <= noDuplicate.getLastRowNum(); row++) {
            String target = noDuplicate.getRow(row).getCell(0).getStringCellValue();
            if (registerData.containsKey(target)) {
                String[] data = registerData.get(target);
                noDuplicateData.add(new String[]{data[0], data[1], data[2], data[3], data[4], data[5], data[6]
                        , data[7], data[8], data[9]});
            }
        }

        for (int ii = 0; ii < noDuplicateData.size(); ii++) {
            Row row = noDuplicate.getRow(ii+1);
            String[] data = noDuplicateData.get(ii);
            row.createCell(1).setCellValue(data[0]);
            row.createCell(2).setCellValue(data[1]);
            row.createCell(3).setCellValue(data[2]);
            row.createCell(4).setCellValue(data[3]);
            row.createCell(5).setCellValue(data[4]);
            row.createCell(6).setCellValue(data[5]);
            row.createCell(7).setCellValue(data[6]);
            row.createCell(8).setCellValue(data[7]);
            row.createCell(9).setCellValue(data[8]);
            row.createCell(10).setCellValue(data[9]);
        }

        System.out.println ("02. "+bookLaporanIGD.getSheetName (2)+" is done");


        FileOutputStream IGDHalfDone = new FileOutputStream (Year + " " + Month + " IGD Half Done.xlsx");
        bookLaporanIGD.write (IGDHalfDone);

    }

}

