import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class Uygulama2 {

    public static void main(String[] args) throws IOException {

        String Dosya = "MyfirstExcel.xlsx";
        FileInputStream excel = new FileInputStream ( Dosya );
        XSSFWorkbook ck = new XSSFWorkbook (excel);
        XSSFSheet sayfa =  ck.getSheetAt ( 0 );
        Iterator <Row> kademe = sayfa.iterator ();
        while(kademe.hasNext ()){
            Row satir = kademe.next ();
            Iterator<Cell> kademesutun = satir.iterator ();


            while (kademesutun.hasNext ()){
                Cell hucre = kademesutun.next ();
                    if(hucre.getCellType () == CellType.STRING){
                        System.out.print (hucre.getStringCellValue () + "....");
                    }else if (hucre.getCellType () == CellType.NUMERIC){
                        System.out.print (hucre.getNumericCellValue () + "....");
                    }

            }
            System.out.println ("");
        }

    }
}
