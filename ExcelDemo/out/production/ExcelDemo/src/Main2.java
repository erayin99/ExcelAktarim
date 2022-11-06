import javax.swing.*;

public class Main2 {

    public static void main(String[] args) {

        try {
            UIManager.setLookAndFeel ( UIManager.getLookAndFeel () );
        } catch (UnsupportedLookAndFeelException e) {
            throw new RuntimeException ( e );
        }


        SwingUtilities.invokeLater ( new Runnable () {
            @Override
            public void run() {
                JtableToExcel jfe = new JtableToExcel ();
                jfe.setVisible ( true );
            }
        } );

    }

}
