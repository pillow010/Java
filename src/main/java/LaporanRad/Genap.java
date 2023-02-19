package LaporanRad;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTEdnDocProps;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Genap {
    public static void main(String[] args) {
    new Genap ();
    }

    private Workbook BookPertindakanNew;

    private FileOutputStream outputStream;

    public Genap(){
        try {
            InputStream pertindakanNew = new FileInputStream ("C:\\Users\\9a06s\\IdeaProjects\\Java\\pertindakanNew.xlsx");
            BookPertindakanNew = new XSSFWorkbook (pertindakanNew);
            Sheet Genap = BookPertindakanNew.getSheetAt (2);

//        buat sheet 4 Jml tndakan per cr Byr pr hri
            BookPertindakanNew.createSheet ();
            Sheet TndkanCrByrHr = BookPertindakanNew.getSheetAt (4);
            BookPertindakanNew.setSheetName (4, "2.Jml tndakan per cr Byr pr hri");

//            Map<String, Integer> countMap = new HashMap<> ();
//            String key;
//            for (int row = 1; row <= Genap.getLastRowNum(); row++) {
//                key = Genap.getRow(row).getCell(9).getStringCellValue().substring (0,10) + "_" + // TGL_MASUK
//                        Genap.getRow(row).getCell(15).getStringCellValue() + "_" + // NM_TINDAKAN
//                        Genap.getRow(row).getCell(8).getStringCellValue(); // JNS_CR_BYR
//                countMap.put(key, countMap.getOrDefault(key, 0) + 1); // increment the count for the key
//            }
//
//
//            // Create the list of objects
//            List<Object> master = new ArrayList<>();
//            for (Map.Entry<String, Integer> entry : countMap.entrySet()) {
//                String[] parts = entry.getKey().split("_");
//                String tglMsk = parts[0];
//                String tindakan = parts[1];
//                String jnsCrByr = parts[2];
//                int count = entry.getValue();
//                master.add(List.of(tglMsk, tindakan, jnsCrByr, count));
//            }

            // print the master list
//            System.out.println(master.stream().sorted ());
            List<String> masterCaraBayar = new ArrayList<>();
            List<String> masterTanggal = new ArrayList<>();
            List<String> masterTindakan = new ArrayList<>();

//            for (Row row : Genap) {
            for (Row row = Genap.getRow(1); row != null; row = Genap.getRow(row.getRowNum()+1)) {
//                System.out.println (row.getCell (9).getStringCellValue ());
                String caraBayar = row.getCell(8).getStringCellValue();
                String tglMsk = row.getCell(9).getStringCellValue().substring(0, 10);
                String tndk = row.getCell(15).getStringCellValue();

                if (!masterCaraBayar.contains(caraBayar)) {
                    masterCaraBayar.add(caraBayar);
                }
                if (!masterTanggal.contains(tglMsk)) {
                    masterTanggal.add(tglMsk);
                }
                if (!masterTindakan.contains(tndk)) {
                    masterTindakan.add(tndk);
                }
            }

            System.out.println("Cara Bayar:");
            masterCaraBayar.stream().sorted().forEach(System.out::println);

            TndkanCrByrHr.createRow (0).createCell(0).setCellValue("Tanggal");

            int rowNum = 1;
            for (String tglMsk : masterTanggal) {
                Row row = TndkanCrByrHr.createRow(rowNum++);
                row.createCell(0).setCellValue(tglMsk);
            }

            System.out.println("\nTindakan:");
            masterTindakan.stream().sorted().forEach(System.out::println);

























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

}