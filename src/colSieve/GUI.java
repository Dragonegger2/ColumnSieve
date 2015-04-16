package colSieve;

import javax.swing.*;
import java.awt.*;

/**
 * Created by ctangney on 4/9/2015.
 */
public class GUI{
    private JPanel wrapper;
    private JLabel lblFiletype;
    private JLabel lblPaths;
    private JLabel lblOptions;
    private JRadioButton rdoXLS;
    private JRadioButton rdoXLSX;
    private JRadioButton rdoCSV;
    private JRadioButton rdoTXT;
    private JPanel paneFiletype;
    private JPanel panePath;
    private JPanel paneOptions;
    private JPanel paneTop;
    private JPanel panePathComponent;
    private JTextField txtTemplate;
    private JTextField txtInput;
    private JTextField txtOutput;
    private JLabel lblTemplate;
    private JLabel lblInput;
    private JLabel lblOutput;
    private JButton btnCompare;
    private JButton btnMap;
    private JButton btn03;
    private JButton btn04;
    private JPanel paneBottom;
    private JTextArea txtConsole;
    private JLabel lblConsole;

    public static void main(String[] args) {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (ClassNotFoundException e) {
                e.printStackTrace();
            } catch (InstantiationException e) {
                e.printStackTrace();
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            } catch (UnsupportedLookAndFeelException e) {
                e.printStackTrace();
            }
        //GUI Objects
        JFrame gui = new JFrame("Column Sieve");


        printToConsole("A string to add");
        gui.setContentPane(new GUI().wrapper);
        gui.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        gui.pack();
        gui.setVisible(true);
    }

    public static void printToConsole(String myStr){
        JTextArea console = new GUI().txtConsole;
        console.append(myStr);
    }
}
