package LaporanJasa;

import java.io.File;
import java.io.IOException;

public class LaporanJasaCommandCenter {
    public static void main(String[] args) throws IOException {
        new A_RekapJasaDokterDanUnit ();
        new B_RincianTindakanJasaDokter();
        new C_RekapPasienJasaDokter();
        new D_RekapPasienJasaUnit();
        new E_RincianJasaNoname();

//        C:\sat work\test\a) LAPORAN REKAP PENERIMAAN JASA UNIT PER PASIEN.xls
//        C:\sat work\test\b) LAPORAN REKAP PENERIMAAN JASA PELAYANAN PER PASIEN.xls
//        C:\sat work\test\c) LAPORAN PENERIMAAN JASA PELAYANAN PER TINDAKAN.xls
//        C:\sat work\test\d) LAPORAN PENERIMAAN JASA PELAYANAN TANPA PEMILIK.xls
    }
}
