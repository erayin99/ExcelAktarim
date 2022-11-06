import org.apache.logging.log4j.categories.Appenders;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

import java.util.Iterator;

import java.util.Vector;


public class JtableToExcel extends JFrame{


    private JTable table1;
    private JButton excelDosyaSecButton;


        JTable jt = new JTable ();
        JButton jb = new JButton ();
    public JtableToExcel() {

        add (excelDosyaSecButton);
        setSize ( 250, 250 );

        setDefaultCloseOperation ( JFrame.DISPOSE_ON_CLOSE);
        excelDosyaSecButton.addActionListener ( new ActionListener () {
            @Override
            public void actionPerformed(ActionEvent e) {

                Vector  tblbaslik = new Vector ();
                ;
                Vector tblveri = new Vector ();


                JFileChooser xlsxDosya = new JFileChooser ();
                int secim = xlsxDosya.showOpenDialog ( xlsxDosya );
                if (secim == JFileChooser.APPROVE_OPTION) {
                    File dosya = xlsxDosya.getSelectedFile ();
                    if (!dosya.getName ().endsWith ( "xlsx" )) {
                        JOptionPane.showMessageDialog ( xlsxDosya, "Lutfen excel dosya sec",
                                "hata", JOptionPane.ERROR_MESSAGE );
                    } else {
                        try {
                            XSSFWorkbook wb = new XSSFWorkbook ( dosya );
                            Sheet sayfa = wb.getSheetAt ( 0 );
                            tblbaslik.clear ();
                            Row ilksatir = sayfa.getRow ( 0 );
                            Iterator<Cell> kademebaslik = ilksatir.iterator ();
                            while (kademebaslik.hasNext ()) {
                                Cell hucre = kademebaslik.next ();
                                tblbaslik.add ( hucre.getStringCellValue () );
                                ;

                            }
                            tblveri.clear ();
                            for (int s = 1; s < sayfa.getLastRowNum (); s++) {
                                Row satir = sayfa.getRow ( s );
                                Vector birkayit = new Vector ();
                                Iterator<Cell> kademesatir = satir.iterator ();
                                while (kademebaslik.hasNext ()) {
                                    Cell hucre = kademesatir.next ();
                                    if (hucre.getCellType () == CellType.STRING) {
                                        birkayit.add ( hucre.getStringCellValue () );

                                    } else if (hucre.getCellType () == CellType.NUMERIC) {

                                    }
                                }
                                tblveri.add ( birkayit );
                            }
                            DefaultTableModel dtm = new DefaultTableModel ( tblveri,  tblbaslik );
                            table1.setModel (  dtm );

                        } catch (IOException ex) {
                            throw new RuntimeException ( ex );
                        } catch (InvalidFormatException ex) {
                            throw new RuntimeException ( ex );
                        }
                    }
                }
            }

        } );

}
    }

