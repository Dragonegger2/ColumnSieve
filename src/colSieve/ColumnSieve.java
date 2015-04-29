package colSieve;

import javax.swing.*;
import javax.swing.border.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;

/**
 * Created by ctangney on 4/29/2015.
 */
public class ColumnSieve extends JFrame{
    private JPanel wrapper;
    private JPanel pnlTop;
    private JPanel pnlBot;
    private JTabbedPane tabOptions;
    private JPanel pnlHome;
    private JPanel pnlCompare;
    private JPanel pnlSieve;
    private JTextArea txtConsole;
    private JLabel lblConsole;
    private JLabel lblWelcome;
    private JLabel lblDesc;
    private JLabel lblCompare;
    private JPanel pnlCompareFileType;
    private JLabel lblCompareFileType;
    private JPanel pnlFileTypeGroup;
    private JRadioButton rdoCompareXLS;
    private JRadioButton rdoCompareXLSX;
    private JRadioButton rdoCompareCSV;
    private JRadioButton rdoCompareTXT;
    private JPanel pnlCompareInputFile;
    private JLabel lblCompareInputProperties;
    private JTextField txtCompareInput;
    private JButton btnCompareInputBrowse;
    private JLabel lblCompareInput;
    private JTextField txtCompareInputSheet;
    private JLabel lblCompareInputSheet;
    private JPanel pnlCompareTemplateFile;
    private JTextField txtCompareTemplate;
    private JButton btnCompareTemplateBrowse;
    private JLabel lblCompareTemplateFile;
    private JTextField txtCompareTemplateSheet;
    private JLabel lblCompareTemplateSheet;
    private JLabel lblCompareTemplateProperties;
    private JButton btnCompareColumns;
    private JMenuBar menuBar;
    private JMenu menuFile, menuMore;
    private JMenuItem menuFile_EXIT, menuMore_ABOUT;

    public void setGui(){
        //set the initial elements to visible
        wrapper.setVisible(true);
        pnlTop.setVisible(true);
        pnlBot.setVisible(true);
        pnlHome.setVisible(true);
        tabOptions.setVisible(true);
        lblConsole.setVisible(true);

        //set up the menu bar
        menuBar = new JMenuBar();
        menuFile = new JMenu("File");
        menuMore = new JMenu("More");
        menuFile.setMnemonic(KeyEvent.VK_F);
        menuFile.getAccessibleContext().setAccessibleDescription("File menu.");
        menuMore.setMnemonic(KeyEvent.VK_M);
        menuMore.getAccessibleContext().setAccessibleDescription("More about the program.");

        //set up the file menu
        menuFile_EXIT = new JMenuItem("Exit",  KeyEvent.VK_F4);
        menuFile_EXIT.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_F4, ActionEvent.ALT_MASK));
        menuFile.add(menuFile_EXIT);
        //add the menuFile to the bar
        menuBar.add(menuFile);

        //set up the more menu
        menuMore_ABOUT = new JMenuItem("About", KeyEvent.VK_A);
        menuMore_ABOUT.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_A, ActionEvent.ALT_MASK));
        menuMore.add(menuMore_ABOUT);
        //add the menuMore to the bar
        menuBar.add(menuMore);

        //show the menu bar
        this.setJMenuBar(menuBar);

        //set the borders up
        Border border = BorderFactory.createLoweredBevelBorder();
        txtConsole.setBorder(border);
        txtConsole.setVisible(true);

        border = BorderFactory.createEtchedBorder();
        pnlCompareFileType.setBorder(border);
        pnlCompareInputFile.setBorder(border);
        pnlCompareTemplateFile.setBorder(border);

        //window behavior
        this.setContentPane(wrapper);
        this.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        this.setSize(new Dimension(1075, 910));
        this.setMinimumSize(new Dimension(1075, 910));
        this.setResizable(false);
    }

    public void showGui(){
        //update the console
        txtConsole.append("\n> waiting for input...");

        //set the title and show the frame
        this.setTitle("Column Sieve 1.0");
        this.setVisible(true);

        //thread to watch the exit menu item
        menuFile_EXIT.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                System.exit(0);
            }
        });
    }

    public void setConsole(String str){
        txtConsole.append(str);
    }

    //main method
    public static void main(String args[]) {
        //set the look and feel of the UI to the current system theme
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

        //make a new gui object
        ColumnSieve gui = new ColumnSieve();

        //set the gui up
        gui.setGui();

        //show the gui
        gui.showGui();
    }
}
