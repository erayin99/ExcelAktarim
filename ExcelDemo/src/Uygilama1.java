
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;



public class Uygilama1 {

    private static final String FILE_NAME = "MyfirstExcel.xlsx";

    public static void main(String[] args)  {

        XSSFWorkbook wb = new XSSFWorkbook ();
        XSSFSheet cs = wb.createSheet ("Ornek 1");
        //Sheet cs = wb.createSheet("Sayfa1");  //diger bir sheet tanimlama yontemi
        Object[][] tablo = {
                {"Veri_Yapisi", "Veri_Tipi", "Buyukluk"},
                {"int", "Ilkel", 2},
                {"float", "Ilkel", 4},
                {"String", "Gelismis", "Belirsiz"}
        };
        int satirNo = 0;
        System.out.println ( "Excel dosyasi olusturuluyor" );
        for (Object[] tablosatir : tablo) {
            Row satir = cs.createRow ( satirNo++ );
            int sutunNo = 0;
            for (Object tabloHucre : tablosatir) {
                Cell hucre = satir.createCell ( sutunNo++ );
                if (tabloHucre instanceof String) {
                    hucre.setCellValue ( (String) tabloHucre );

                } else if (tabloHucre instanceof Integer) {
                    hucre.setCellValue ( (Integer) tabloHucre );
                }
            }

        }
        try{
            FileOutputStream fos = new FileOutputStream(FILE_NAME);
            
                wb.write (fos);
                wb.close ();

        } catch (FileNotFoundException e) {
            System.out.println ( "dosya bulunamadi hatasi/n" + e.toString () );
        } catch (IOException e) {
            System.out.println ( "Dosya Yazdirma hatasi/n" + e.toString ());
        }
        System.out.println ( "Veriler Excel Dosyasina Yazdirildi" );

    }


}
