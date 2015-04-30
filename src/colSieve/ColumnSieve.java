package colSieve;

import colSieve.logic.ColSieve;

import javax.swing.*;
import javax.swing.border.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.io.File;

/**
 * Created by ctangney on 4/29/2015.
 */
public class ColumnSieve extends JFrame implements ActionListener{
    //form items
    private JPanel wrapper, pnlTop, pnlBot, pnlHome, pnlCompare, pnlNewTemplate, pnlCompareFileType, pnlCompareFileTypeGroup, pnlCompareInputFile, pnlCompareTemplateFile;
    private JLabel lblWelcome, lblConsole, lblDesc, lblCompare, lblCompareFileType, lblCompareInput, lblCompareInputProperties, lblCompareTemplateFile, lblCompareTemplateSheet, lblCompareTemplateProperties;
    private JTabbedPane tabOptions;
    private JTextArea txtConsole;
    private ButtonGroup grpCompareFileType, grpSieveColumns;
    private JRadioButton rdoCompareXLS, rdoCompareXLSX, rdoCompareCSV, rdoCompareTXT;
    private JTextField txtCompareInput, txtCompareTemplate;
    private JButton btnCompareInputBrowse, btnCompareTemplateBrowse, btnCompareColumns;
    private JPanel pnlSieveFileType;
    private JPanel pnlSieveInputProperties;
    private JPanel pnlSieveTemplateProperties;
    private JPanel pnlSieveOutputProperties;
    private JRadioButton rdoSieveXLS;
    private JRadioButton rdoSieveXLSX;
    private JRadioButton rdoSieveCSV;
    private JRadioButton rdoSieveTXT;
    private JPanel pnlSieve;
    private JLabel lblSieve;
    private JLabel lblSieveFileType;
    private JPanel pnlSieveFileTypeGroup;
    private JTextField txtSieveInput;
    private JButton txtSieveInputBrowse;
    private JLabel lblSieveInput;
    private JLabel lblSieveInputProperties;
    private JTextField txtSieveTemplate;
    private JButton btnSieveTemplateBrowse;
    private JLabel lblSieveTemplate;
    private JLabel lblSieveTemplateProperties;
    private JButton btnSieveColumns;
    private JTextField txtSieveOutput;
    private JLabel lblSieveOutput;
    private JLabel txtSieveOutputProperties;
    private JButton btnSieveOutSave;
    private JScrollPane sievePane;
    private JMenuBar menuBar;
    private JMenu menuFile, menuMore;
    private JMenuItem menuFile_EXIT, menuMore_ABOUT;

    //private variables
    private String input, template, output, inputSheet, templateSheet, outputSheet;
    private File inFile, templateFile, outputFile;
    private int confirmResult;

    //A ColSieve object to pass input to the FileCommand objects
    private ColSieve passInput = new ColSieve();

    public void setGui(){
        //set the initial elements to visible
        wrapper.setVisible(true);
        pnlTop.setVisible(true);
        pnlBot.setVisible(true);
        pnlHome.setVisible(true);
        tabOptions.setVisible(true);
        lblWelcome.setVisible(true);
        lblDesc.setVisible(true);

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
        //thread to watch the item
        menuFile_EXIT.addActionListener(this);
        menuFile.add(menuFile_EXIT);
        //add the menuFile to the bar
        menuBar.add(menuFile);

        //set up the about menu item
        menuMore_ABOUT = new JMenuItem("About", KeyEvent.VK_A);
        menuMore_ABOUT.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_A, ActionEvent.ALT_MASK));
        //thread to watch the item
        menuMore_ABOUT.addActionListener(this);
        menuMore.add(menuMore_ABOUT);
        //add the menuMore to the bar
        menuBar.add(menuMore);

        //show the menu bar
        this.setJMenuBar(menuBar);

        //set the borders up
        //Lowered Bevel
        Border border = BorderFactory.createLoweredBevelBorder();
        txtConsole.setBorder(border);
        txtConsole.setAutoscrolls(true);
        txtConsole.setVisible(true);

        //Etched
        border = BorderFactory.createEtchedBorder();
        pnlCompareFileType.setBorder(border);
        pnlCompareInputFile.setBorder(border);
        pnlCompareTemplateFile.setBorder(border);
        pnlSieveFileType.setBorder(border);
        pnlSieveInputProperties.setBorder(border);
        pnlSieveTemplateProperties.setBorder(border);
        pnlSieveOutputProperties.setBorder(border);

        //set up the Radio Button groups
        grpCompareFileType = new ButtonGroup();
        grpSieveColumns = new ButtonGroup();
        grpCompareFileType.add(rdoCompareXLS);
        grpCompareFileType.add(rdoCompareXLSX);
        grpCompareFileType.add(rdoCompareCSV);
        grpCompareFileType.add(rdoCompareTXT);
        grpSieveColumns.add(rdoSieveXLS);
        grpSieveColumns.add(rdoSieveXLSX);
        grpSieveColumns.add(rdoSieveCSV);
        grpSieveColumns.add(rdoSieveTXT);

        //items which are watched by thread
        btnCompareColumns.addActionListener(this);
        btnCompareInputBrowse.addActionListener(this);
        btnCompareTemplateBrowse.addActionListener(this);

        //window behavior
        this.setContentPane(wrapper);
        this.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        this.setSize(new Dimension(1200, 915));
        this.setMinimumSize(new Dimension(1200, 915));
    }

    public void showGui(){
        //update the console
        setConsole("> Welcome to the Column Sieve Tool!\n\n> Waiting for input...\n");

        //set the title and show the frame
        this.setTitle("Column Sieve 1.0");
        this.setVisible(true);
    }

    public void setConsole(String str){
        //updates the console with a new line of text
        txtConsole.append(str);
        txtConsole.setCaretPosition(txtConsole.getDocument().getLength());
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

    @Override
    public void actionPerformed(ActionEvent e){
        //local variables
        Boolean inputExists = false;
        Boolean templateExists = false;
        Boolean sheetsExist = false;

        if(e.getSource() == menuFile_EXIT){
            //file menu > exit clicked
            System.exit(0);
        }else if(e.getSource() == btnCompareColumns){
            //execute compare columns button clicked
            //set the file command
            passInput.setConsoleFileCommand("compareHeader");

            //if the user has not selected a file type
            if(grpCompareFileType.getSelection() == null){
                //make sure the user has entered files
                JOptionPane.showMessageDialog(null,"Please select a file type before proceeding.");
            }else {

                //if the user has not entered file properties
                if (txtCompareInput.getText().equals("") || txtCompareTemplate.getText().equals("")) {
                    //make sure the user has entered files
                    JOptionPane.showMessageDialog(null, "Please select an input and template file before proceeding.");
                } else {
                    //set the file type
                    if (rdoCompareXLS.isSelected()) {
                        passInput.setConsoleFileType("XLS");
                    } else if (rdoCompareXLSX.isSelected()) {
                        passInput.setConsoleFileType("XLSX");
                    } else if (rdoCompareCSV.isSelected()) {
                        passInput.setConsoleFileType("CSV");
                    } else if (rdoCompareTXT.isSelected()) {
                        passInput.setConsoleFileType("TXT");
                    }

                    //check that input file exists
                    if (new File(txtCompareInput.getText()).exists()) {
                        //if the input file exists
                        inputExists = true;
                        inFile = new File(txtCompareInput.getText());
                        passInput.setConsoleInFile(inFile.getAbsolutePath());
                    }

                    //check that the template file exists
                    if (new File(txtCompareTemplate.getText()).exists()) {
                        //if the template file exists
                        templateExists = true;
                        templateFile = new File(txtCompareTemplate.getText());
                        passInput.setConsoleTemplateFile(templateFile.getAbsolutePath());
                    }

                    //if the files exist, set the ColSieve information
                    if (inputExists && templateExists && !(txtCompareInput.getText().equals("")) && !(txtCompareTemplate.getText().equals(""))) {
                        passInput.setConsoleInSheet();
                        passInput.setConsoleTemplateSheet();

                        if (passInput.checkSheets()) {
                            //if the sheets exist
                            sheetsExist = true;
                        } else {
                            //sheets are missing
                            JOptionPane.showMessageDialog(null, "The tool is unable to open one or more of the sheets you have specified.\nPlease double check that you have entered both sheet names correctly before proceeding.");
                        }
                        if (sheetsExist) {
                            //If the sheets exist, then all other checks have been completed.
                            //Update the console, make a call to the compareHeader function
                            txtConsole.append("\n> The tool is now comparing your files.\n> Please wait...\n");
                            passInput.execute(passInput.getConsoleFileType(), this);
                        }
                    } else if ((!inputExists) && (!templateExists)) {
                        JOptionPane.showMessageDialog(null, "The tool is unable to locate one or more of the files you have entered.\nPlease confirm the location of the file before proceeding.");
                    } else if (!inputExists) {
                        JOptionPane.showMessageDialog(null, "The tool is unable to locate the input file you have selected.\nPlease confirm the location of the file before proceeding.");
                    } else if (!templateExists) {
                        JOptionPane.showMessageDialog(null, "The tool is unable to locate the template file you have selected.\nPlease confirm the location of the file before proceeding.");
                    }
                }
            }
        }else if(e.getSource() == btnCompareInputBrowse){
            //compare columns browse for input file clicked
            final JFileChooser fc = new JFileChooser("Please select an input file");
            int returnVal = fc.showOpenDialog(ColumnSieve.this);

            if(returnVal == JFileChooser.APPROVE_OPTION){
                inFile = fc.getSelectedFile();
                txtCompareInput.setText(inFile.getAbsolutePath());
            }
        }else if(e.getSource() == btnCompareTemplateBrowse){
            //compare columns browse for template file clicked
            final JFileChooser fc = new JFileChooser("Please select a template file");
            int returnVal = fc.showOpenDialog(ColumnSieve.this);

            if(returnVal == JFileChooser.APPROVE_OPTION){
                templateFile = fc.getSelectedFile();
                txtCompareTemplate.setText(templateFile.getAbsolutePath());
            }
        }
    }
}
