package LaporanRadiologi;

import StylingLaporan.StylerRepo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;



import java.io.*;
import java.util.*;

import java.util.stream.IntStream;

public class LabHalfDone extends StylerRepo{
    public static void main(String[] args) {
        new LabHalfDone ();

    }
    private Workbook BookPertindakanNew;

    private FileOutputStream outputStream;

    public LabHalfDone(){
        Sheet SheetA = null;
        Sheet sheetB = null;
//        File pertindakanNew = new File("C:\\sat work\\test\\rad pertindakan new.xlsx");
        try {
//            POIFSFileSystem poifs = new POIFSFileSystem(pertindakanNew);
//            BookPertindakanNew = new HSSFWorkbook(poifs);

            InputStream pertindakanNew = new FileInputStream ("C:\\sat work\\test\\rad pertindakan new.xlsx");
            BookPertindakanNew = new XSSFWorkbook(pertindakanNew);


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
            System.out.println ("0. Doing "+ BookPertindakanNew.getSheetAt (0).getSheetName ());

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
                                +pertindakan_New_Raw.getRow (i).getCell (15).getStringCellValue ()
                        );
                    }
                }
            }
            System.out.println ("0. "+ BookPertindakanNew.getSheetAt (0).getSheetName ()+" Complete");


//          buat sheet 1 Ganjil
            Sheet Ganjil = BookPertindakanNew.createSheet();
            BookPertindakanNew.setSheetName(1, "Ganjil");
            System.out.println ("1. Doing "+ BookPertindakanNew.getSheetAt (1).getSheetName ());


            Set<String> uniqueValues = new HashSet<>();
            for (int row = 1; row <= pertindakan_New_Raw.getLastRowNum(); row++) {
                if (pertindakan_New_Raw.getRow(row) != null) { // check if row is not empty
                    Cell cell = pertindakan_New_Raw.getRow(row).getCell(29);
                    if (cell != null) { // check if cell is not empty
                        String cellValue = cell.getStringCellValue();
                        if (!cellValue.isBlank()) { // check if cell value is not blank
                            uniqueValues.add(cellValue);
                        }
                    }
                }
            }


            for (int i=0;i<= uniqueValues.size ();i++){
                Ganjil.createRow (i);
            }
            Ganjil.getRow (0).createCell (0).setCellValue ("NOREG");
            Ganjil.getRow (0).createCell (1).setCellValue ("JENIS CARA BAYAR");
            Ganjil.getRow (0).createCell (2).setCellValue ("TANGGAL MASUK");
            Ganjil.getRow (0).createCell (3).setCellValue ("NIC INST ASAL");

            List<String> sortedValues = uniqueValues.stream ().sorted ().toList ();

            IntStream.range(0, sortedValues.size())
                    .forEach(i -> {
                        String value = sortedValues.get(i);
                        Ganjil.getRow(i+1).createCell(0).setCellValue(value);
                    });

            for (int row = 1; row <= sortedValues.size (); row++) {
                String cellValue = Ganjil.getRow(row).getCell(0).getStringCellValue();
                for (int pertRow = 1; pertRow <= pertindakan_New_Raw.getLastRowNum(); pertRow++) {
                    String pertCellValue = pertindakan_New_Raw.getRow(pertRow).getCell(29).getStringCellValue();
                    if (cellValue.equals(pertCellValue)) {
                        String JnsCrByr = pertindakan_New_Raw.getRow(pertRow).getCell(8).getStringCellValue ();
                        String TglMsk = pertindakan_New_Raw.getRow(pertRow).getCell(9).getStringCellValue ().substring (0,10);
                        String NicInstAsal = pertindakan_New_Raw.getRow(pertRow).getCell(24).getStringCellValue ();
                        Ganjil.getRow(row).createCell(1).setCellValue(JnsCrByr);
                        Ganjil.getRow(row).createCell(2).setCellValue(TglMsk);
                        Ganjil.getRow(row).createCell(3).setCellValue(NicInstAsal);
                        break;
                    }
                }
            }

//          cek per row. sesuaikan width nya
            for (int columnIndex = 0; columnIndex < Ganjil.getRow (0).getLastCellNum (); columnIndex++) {
                Ganjil.autoSizeColumn(columnIndex);
            }
            System.out.println ("1. "+ BookPertindakanNew.getSheetAt (1).getSheetName ()+" Complete");


//          buat sheet 2 Genap
            Sheet Genap = BookPertindakanNew.createSheet();
            BookPertindakanNew.setSheetName(2, "Genap");
            System.out.println ("2. Doing "+ BookPertindakanNew.getSheetAt (2).getSheetName ());

//
            List<String> values = new ArrayList<> ();
            for (int row = 1; row <= pertindakan_New_Raw.getLastRowNum(); row++) {
                String cellValue =   pertindakan_New_Raw.getRow(row).getCell(30).getStringCellValue ();
                String Tindakan = pertindakan_New_Raw.getRow(row).getCell(15).getStringCellValue();
                if (!Tindakan.contains("PAKET")) {
                    values.add (cellValue);
                }
            }

            Genap.createRow (0);
            for (int cell=0;cell<pertindakan_New_Raw.getRow (0).getLastCellNum ();cell++){
                Genap.getRow (0).createCell (cell).setCellValue (
                        pertindakan_New_Raw.getRow (0).getCell (cell).getStringCellValue ()
                );
            }

            List<String> sortedGenapValues = values.stream ().sorted ().toList ();
            IntStream.range(0, sortedGenapValues.size())
                    .forEach(i -> {
                        String value = sortedGenapValues.get (i);
                        Genap.createRow (i + 1).createCell (30).setCellValue (value);
                    });

            for (int row = 1; row <= sortedGenapValues.size (); row++) {
                String cellValue = Genap.getRow(row).getCell(30).getStringCellValue();
                for (int pertRow = 1; pertRow <= pertindakan_New_Raw.getLastRowNum(); pertRow++) {
                    String pertCellValue = pertindakan_New_Raw.getRow(pertRow).getCell(30).getStringCellValue();
                    if (cellValue.equals(pertCellValue)) {
                        for (int cell = pertindakan_New_Raw.getRow (0).getLastCellNum (); cell >= 0; cell--) {
//                            Cell currentCell = pertindakan_New_Raw.getRow (row).getCell (cell);
                            if (pertindakan_New_Raw
                                    .getRow (pertRow)
                                    .getCell (cell)!= null) {
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

//          CT Scan, USG , RONTGENT, Konsul Dokter Spesialis
            for (int row = 1; row <= sortedGenapValues.size (); row++) {
                String Tindakan = Genap.getRow (row).getCell (15).getStringCellValue ();
                if (Tindakan.contains ("CT Scan")) {
                    Genap.getRow (row).createCell (15).setCellValue ("CT Scan");
                }else if (Tindakan.contains ("USG")) {
                    Genap.getRow (row).createCell (15).setCellValue ("USG");
                } else if (Tindakan.contains ("Konsul Dokter Spesialis")) {
                    Genap.getRow (row).createCell (15).setCellValue ("Konsul Dokter Spesialis");
                }else {
                    Genap.getRow (row).createCell (15).setCellValue ("RONTGENT");
                }
            }


//          cek per row. sesuaikan width nya
            for (int columnIndex = 0; columnIndex < Genap.getRow (0).getLastCellNum (); columnIndex++) {
                Genap.autoSizeColumn(columnIndex);
            }
            System.out.println ("2. "+ BookPertindakanNew.getSheetAt (2).getSheetName ()+" Complete");

//          buat sheet 3 pertindakan
            Sheet Pertindakan = BookPertindakanNew.createSheet();
            BookPertindakanNew.setSheetName(3, "1 Pertindakan");
            System.out.println ("3. Doing "+ BookPertindakanNew.getSheetAt (3).getSheetName ());

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

//          buat header center kemudian border semuanya
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
            System.out.println ("3. "+ BookPertindakanNew.getSheetAt (3).getSheetName ()+" Complete");




//        buat sheet 4 Jml tndakan per cr Byr pr hri
            BookPertindakanNew.createSheet ();
            Sheet TndkanCrByrHr = BookPertindakanNew.getSheetAt (4);
            BookPertindakanNew.setSheetName (4, "2.Jml tndakan per cr Byr pr hri");
            System.out.println ("4. Sheet "+ BookPertindakanNew.getSheetAt (4).getSheetName ()+" Created");


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
            Sheet PsnCrByrHr = BookPertindakanNew.getSheetAt (5);
            BookPertindakanNew.setSheetName (5, "3.Pasien per cara bayar pr hari");
            System.out.println ("5. Sheet "+ BookPertindakanNew.getSheetAt (5).getSheetName ()+" Created");


//        buat sheet 6 Jml tndakan per cr Byr pr hri
            BookPertindakanNew.createSheet ();
            Sheet TndCrByr = BookPertindakanNew.getSheetAt (6);
            BookPertindakanNew.setSheetName (6, "4.Tindakan Percara bayar ");
            System.out.println ("6. Doing "+ BookPertindakanNew.getSheetAt (6).getSheetName ());


            Map<String, Integer> tndkCount = new TreeMap<> ();
            for (int row = 1; row <= Genap.getLastRowNum(); row++) {
//                String tndk = Genap.getRow(row).getCell(8).getStringCellValue();
//                if (tndkCount.containsKey(tndk)) {
//                    tndkCount.put(tndk, tndkCount.get(tndk) + 1);
//                } else {
//                    tndkCount.put(tndk, 1);
//                }
                String tndk = Genap.getRow(row).getCell(8).getStringCellValue();
                tndkCount.put(tndk, tndkCount.getOrDefault(tndk, 0) + 1);
            }

            TndCrByr.createRow(0).createCell(0).setCellValue("Jenis Cara Bayar");
            TndCrByr.getRow(0).createCell(1).setCellValue("Jumlah");

            int TndCrByrrow = 0;
            int TndCrByrSum = 0;
            for (Map.Entry<String, Integer> entry : tndkCount.entrySet()) {
                TndCrByrrow++;
                TndCrByr.createRow(TndCrByrrow).createCell(0).setCellValue(entry.getKey());
                TndCrByr.getRow(TndCrByrrow).createCell(1).setCellValue(entry.getValue());
                TndCrByrSum+=entry.getValue ();
            }

            int TndCrByrLastRow = TndCrByr.getLastRowNum ()+1;
            TndCrByr.createRow (TndCrByrLastRow).createCell (0).setCellValue ("Grand Total");
            TndCrByr.getRow (TndCrByrLastRow).createCell (1).setCellValue (TndCrByrSum);

//          buat header center kemudian border semuanya ps. use'<' because return 2 but there is 0, and 1. no number 2.
            for (int rightCell = 0; rightCell<TndCrByr.getRow (0).getLastCellNum ();rightCell++){
                TndCrByr.getRow (0).getCell (rightCell).setCellStyle(BorderCenterCellStyle);
                for (int downRow = 1; downRow<= TndCrByr.getLastRowNum (); downRow++){
                    TndCrByr.getRow (downRow).getCell (rightCell).setCellStyle(AllBorderCellStyle);
                }
            }
//          cek per row. sesuaikan width nya
            int columnCountTndCrByr = TndCrByr.getRow (0).getLastCellNum();
            for (int columnIndex = 0; columnIndex < columnCountTndCrByr; columnIndex++) {
                TndCrByr.autoSizeColumn(columnIndex);
            }
            System.out.println ("6. "+ BookPertindakanNew.getSheetAt (6).getSheetName ()+" Completed");



//        buat sheet 7 Jml tndakan per cr Byr pr hri
            BookPertindakanNew.createSheet ();
            Sheet PsnCrByr = BookPertindakanNew.getSheetAt (7);
            BookPertindakanNew.setSheetName (7, "5.Pasien per cara bayar");
            System.out.println ("7. Doing "+ BookPertindakanNew.getSheetAt (7).getSheetName ());


            Map<String, Integer> PsncrByrCount = new TreeMap<> ();
            for (int row = 1; row <= Ganjil.getLastRowNum(); row++) {
//                String crByr = Ganjil.getRow(row).getCell(1).getStringCellValue();
//                if (PsncrByrCount.containsKey(crByr)) {
//                    PsncrByrCount.put(crByr, PsncrByrCount.get(crByr) + 1);
//                } else {
//                    PsncrByrCount.put(crByr, 1);
//                }
                String crByr = Ganjil.getRow(row).getCell(1).getStringCellValue();
                PsncrByrCount.put(crByr, PsncrByrCount.getOrDefault(crByr, 0) + 1);
            }

            PsnCrByr.createRow(0).createCell(0).setCellValue("Jenis Cara Bayar");
            PsnCrByr.getRow(0).createCell(1).setCellValue("Jumlah");

            int PsnCrByrrow = 0;
            int PsnCrByrSum = 0;
            for (Map.Entry<String, Integer> entry : PsncrByrCount.entrySet()) {
                PsnCrByrrow++;
                PsnCrByr.createRow(PsnCrByrrow).createCell(0).setCellValue(entry.getKey());
                PsnCrByr.getRow(PsnCrByrrow).createCell(1).setCellValue(entry.getValue());
                PsnCrByrSum += entry.getValue ();
            }
            int PsnCrByrLastRow = PsnCrByr.getLastRowNum ()+1;
            PsnCrByr.createRow (PsnCrByrLastRow).createCell (0).setCellValue ("Grand Total");
            PsnCrByr.getRow (PsnCrByrLastRow).createCell (1).setCellValue (PsnCrByrSum);

//          buat header center kemudian border semuanya ps. use'<' because return 2 but there is 0, and 1. no number 2.
            for (int rightCell = 0; rightCell<PsnCrByr.getRow (0).getLastCellNum ();rightCell++){
                PsnCrByr.getRow (0).getCell (rightCell).setCellStyle(BorderCenterCellStyle);
                for (int downRow = 1; downRow<= PsnCrByr.getLastRowNum (); downRow++){
                    PsnCrByr.getRow (downRow).getCell (rightCell).setCellStyle(AllBorderCellStyle);
                }
            }
//          cek per row. sesuaikan width nya
            int columnCountPsnCrByr = PsnCrByr.getRow (0).getLastCellNum();
            for (int columnIndex = 0; columnIndex < columnCountPsnCrByr; columnIndex++) {
                PsnCrByr.autoSizeColumn(columnIndex);
            }
            System.out.println ("7. "+ BookPertindakanNew.getSheetAt (0).getSheetName ()+" Completed");














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
    private static int findColumn(String columnName, Row row) {
        for (int column = 0; column < row.getLastCellNum(); column++) {
            Cell cell = row.getCell(column);
            if (cell.getStringCellValue().equals(columnName)) {
                return column;
            }
        }
        return -1;
    }

}
