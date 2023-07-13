package LaporanJasa;

import java.io.IOException;

public class LaporanJasaCommandCenter {
    public static final String fileSource = "C:\\sat work\\test\\3. JKN SUSULAN NOV RAJAL\\";
    public static final String fileOutput = "C:\\sat work\\test\\";

    public static void main(String[] args) throws IOException {
        new A_RekapJasaDokterDanUnit ();
        new B_RincianTindakanJasaDokter();
        new C_RekapPasienJasaDokter();
        new D_RekapPasienJasaUnit();
        new E_RincianJasaNoname();
    }
}
