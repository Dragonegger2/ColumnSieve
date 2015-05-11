package colSieve;

import colSieve.logic.UserInput;

import javax.swing.*;
import javax.swing.border.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.io.File;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by ctangney on 4/29/2015.
 */
public class ColumnSieve extends JFrame implements ActionListener{
    //form items
    private JPanel wrapper,pnlTop,pnlBot,pnlHome,pnlCompare,pnlNewTemplate,pnlCompareFileType,pnlCompareFileTypeGroup,pnlCompareInputFile,pnlCompareTemplateFile,pnlSieveFileType,pnlSieveInputProperties,pnlSieveTemplateProperties,pnlSieveOutputProperties,pnlSieve,pnlSieveFileTypeGroup;
    private JLabel lblWelcome,lblConsole,lblDesc,lblCompare,lblCompareFileType,lblCompareInput,lblCompareInputProperties,lblCompareTemplateFile,lblCompareTemplateSheet,lblCompareTemplateProperties,lblSieve,lblSieveFileType,lblSieveInput,lblSieveInputProperties,lblSieveTemplate,lblSieveTemplateProperties,lblSieveOutput,lblSieveOutputProperties;
    private JTabbedPane tabOptions;
    private JTextArea txtConsole;
    private ButtonGroup grpCompareFileType,grpSieveColumns;
    private JRadioButton rdoCompareXLS,rdoCompareXLSX,rdoCompareCSV,rdoCompareTXT,rdoSieveXLS,rdoSieveXLSX,rdoSieveCSV,rdoSieveTXT;
    private JTextField txtCompareInput,txtCompareTemplate,txtSieveInput,txtSieveTemplate,txtSieveOutput;
    private JButton btnCompareInputBrowse,btnCompareTemplateBrowse,btnCompareColumns,btnSieveInputBrowse,btnSieveTemplateBrowse,btnSieveColumns,btnSieveOutSave;
    private JScrollPane sievePane;
    private JMenuBar menuBar;
    private JMenu menuFile, menuMore;
    private JMenuItem menuFile_EXIT, menuMore_ABOUT;

    //private variables
    private String input, template, output;
    private File inFile, templateFile, outputFile;
    private Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();

    //A UserInput object to pass input to the FileCommand objects
    private UserInput passInput = new UserInput();

    //File chooser to get files
    private JFileChooser fc = new JFileChooser();

    public void setGui(){
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
        rdoCompareXLS.addActionListener(this);
        rdoCompareXLSX.addActionListener(this);
        rdoCompareCSV.addActionListener(this);
        rdoCompareTXT.addActionListener(this);
        btnSieveColumns.addActionListener(this);
        btnSieveInputBrowse.addActionListener(this);
        btnSieveTemplateBrowse.addActionListener(this);
        btnSieveOutSave.addActionListener(this);
        rdoSieveXLS.addActionListener(this);
        rdoSieveXLSX.addActionListener(this);
        rdoSieveCSV.addActionListener(this);
        rdoSieveTXT.addActionListener(this);

        //items disabled on window load
        txtCompareInput.setEnabled(false);
        btnCompareInputBrowse.setEnabled(false);
        txtCompareTemplate.setEnabled(false);
        btnCompareTemplateBrowse.setEnabled(false);
        txtSieveInput.setEnabled(false);
        btnSieveInputBrowse.setEnabled(false);
        txtSieveTemplate.setEnabled(false);
        btnSieveTemplateBrowse.setEnabled(false);
        txtSieveOutput.setEnabled(false);
        btnSieveOutSave.setEnabled(false);

        //window behavior
        this.setContentPane(wrapper);
        this.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        this.setSize(new Dimension(1200, 915));
        this.setMinimumSize(new Dimension(1200, 915));
        URL iconURL = getClass().getResource("res/bars.png");
        ImageIcon icon = new ImageIcon(iconURL);
        this.setIconImage(icon.getImage());
        this.setLocation(screenSize.width/2-this.getSize().width/2, screenSize.height/2-this.getSize().height/2);
    }

    public void showGui(){
        //update the console
        setConsole("> Welcome to the Column Sieve Tool!\n\n> Waiting for input...\n");

        //set the title and show the frame
        this.setTitle("Column Sieve");
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
        }else if(e.getSource() == rdoCompareXLS) {
            //one of the compare radio buttons was clicked
            txtCompareInput.setEnabled(true);
            btnCompareInputBrowse.setEnabled(true);
            txtCompareTemplate.setEnabled(true);
            btnCompareTemplateBrowse.setEnabled(true);
            //set fileType
            passInput.setConsoleFileType("XLS");
        }else if(e.getSource() == rdoCompareXLSX) {
            //one of the compare radio buttons was clicked
            txtCompareInput.setEnabled(true);
            btnCompareInputBrowse.setEnabled(true);
            txtCompareTemplate.setEnabled(true);
            btnCompareTemplateBrowse.setEnabled(true);
            //set fileType
            passInput.setConsoleFileType("XLSX");
        }else if(e.getSource() == rdoCompareCSV) {
            //one of the compare radio buttons was clicked
            txtCompareInput.setEnabled(true);
            btnCompareInputBrowse.setEnabled(true);
            txtCompareTemplate.setEnabled(true);
            btnCompareTemplateBrowse.setEnabled(true);
            //set fileType
            passInput.setConsoleFileType("CSV");
        }else if(e.getSource() == rdoCompareTXT) {
            //one of the compare radio buttons was clicked
            txtCompareInput.setEnabled(true);
            btnCompareInputBrowse.setEnabled(true);
            txtCompareTemplate.setEnabled(true);
            btnCompareTemplateBrowse.setEnabled(true);
            //set fileType
            passInput.setConsoleFileType("TXT");
        }else if(e.getSource() == btnCompareColumns){
            //execute compare columns button clicked
            //set the file command
            passInput.setConsoleFileCommand("compareHeader");
            passInput.setRunMode(true);

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

                    //if the files exist, set the UserInput information
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
            fc.setDialogTitle("Please select an input file");
            String fileFilter = passInput.getConsoleFileType().toLowerCase();
            fc.setFileFilter(new FileNameExtensionFilter("Excel 97/2003 (."+fileFilter+")",fileFilter));
            fc.getCurrentDirectory();
            int returnVal = fc.showOpenDialog(ColumnSieve.this);

            if(returnVal == JFileChooser.APPROVE_OPTION){
                inFile = fc.getSelectedFile();
                fc.setCurrentDirectory(inFile);
                txtCompareInput.setText(inFile.getAbsolutePath());
            }
        }else if(e.getSource() == btnCompareTemplateBrowse){
            //compare columns browse for template file clicked
            fc.setDialogTitle("Please select a template file");
            String fileFilter = passInput.getConsoleFileType().toLowerCase();
            fc.setFileFilter(new FileNameExtensionFilter("Excel 97/2003 (."+fileFilter+")",fileFilter));
            fc.getCurrentDirectory();
            int returnVal = fc.showOpenDialog(ColumnSieve.this);

            if(returnVal == JFileChooser.APPROVE_OPTION){
                templateFile = fc.getSelectedFile();
                fc.setCurrentDirectory(templateFile);
                txtCompareTemplate.setText(templateFile.getAbsolutePath());
            }
        }else if(e.getSource() == rdoSieveXLS) {
            //one of the compare radio buttons was clicked
            txtSieveInput.setEnabled(true);
            btnSieveInputBrowse.setEnabled(true);
            txtSieveTemplate.setEnabled(true);
            btnSieveTemplateBrowse.setEnabled(true);
            txtSieveOutput.setEnabled(true);
            btnSieveOutSave.setEnabled(true);
            //set fileType
            passInput.setConsoleFileType("XLS");
        }else if(e.getSource() == rdoSieveXLSX) {
            //one of the compare radio buttons was clicked
            txtSieveInput.setEnabled(true);
            btnSieveInputBrowse.setEnabled(true);
            txtSieveTemplate.setEnabled(true);
            btnSieveTemplateBrowse.setEnabled(true);
            txtSieveOutput.setEnabled(true);
            btnSieveOutSave.setEnabled(true);
            //set fileType
            passInput.setConsoleFileType("XLSX");
        }else if(e.getSource() == rdoSieveCSV) {
            //one of the compare radio buttons was clicked
            txtSieveInput.setEnabled(true);
            btnSieveInputBrowse.setEnabled(true);
            txtSieveTemplate.setEnabled(true);
            btnSieveTemplateBrowse.setEnabled(true);
            txtSieveOutput.setEnabled(true);
            btnSieveOutSave.setEnabled(true);
            //set fileType
            passInput.setConsoleFileType("CSV");
        }else if(e.getSource() == rdoSieveTXT) {
            //one of the compare radio buttons was clicked
            txtSieveInput.setEnabled(true);
            btnSieveInputBrowse.setEnabled(true);
            txtSieveTemplate.setEnabled(true);
            btnSieveTemplateBrowse.setEnabled(true);
            txtSieveOutput.setEnabled(true);
            btnSieveOutSave.setEnabled(true);
            //set fileType
            passInput.setConsoleFileType("TXT");
        }else if(e.getSource() == btnSieveColumns){
            //execute sieve columns button pushed
            passInput.setConsoleFileCommand("mapColumnData");
            passInput.setRunMode(true);

            //if the user has not selected a file type
            if(grpSieveColumns.getSelection() == null){
                //make sure the user has entered files
                JOptionPane.showMessageDialog(null,"Please select a file type before proceeding.");
            }else {

                //if the user has not entered file properties
                if (txtSieveInput.getText().equals("") || txtSieveTemplate.getText().equals("")) {
                    //make sure the user has entered files
                    JOptionPane.showMessageDialog(null, "Please select an input file, template file, and choose a save location before proceeding.");
                } else {
                    //check that input file exists
                    if (new File(txtSieveInput.getText()).exists()) {
                        //if the input file exists
                        inputExists = true;
                        inFile = new File(txtSieveInput.getText());
                        passInput.setConsoleInFile(inFile.getAbsolutePath());
                    }

                    //check that the template file exists
                    if (new File(txtSieveTemplate.getText()).exists()) {
                        //if the template file exists
                        templateExists = true;
                        templateFile = new File(txtSieveTemplate.getText());
                        passInput.setConsoleTemplateFile(templateFile.getAbsolutePath());
                    }

                    //if the files exist, set the UserInput information
                    if (inputExists && templateExists && !(txtSieveInput.getText().equals("")) && !(txtSieveTemplate.getText().equals("")) && !(txtSieveOutput.getText().equals(""))) {
                        passInput.setConsoleInSheet();
                        passInput.setConsoleTemplateSheet();
                        passInput.setConsoleOutFile(txtSieveOutput.getText());

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
                            txtConsole.append("\n> The tool will now attempt to sieve the columns from your input file.\n> Please wait...\n");
                            passInput.execute(passInput.getConsoleFileType(), this);
                        }
                    } else if (inputExists && templateExists && !(txtSieveInput.getText().equals("")) && !(txtSieveTemplate.getText().equals("")) && txtSieveOutput.getText().equals("")){
                        final JOptionPane defaultSaveWarning = new JOptionPane();
                        int returnVal = defaultSaveWarning.showOptionDialog(null, "The tool has not detected a save location.\nBy default, the tool will create a new file on the desktop.", "Warning.",JOptionPane.YES_NO_CANCEL_OPTION,JOptionPane.WARNING_MESSAGE,null,null,null);
                        if(returnVal == JOptionPane.YES_OPTION){
                            passInput.setConsoleInSheet();
                            passInput.setConsoleTemplateSheet();
                            Date currentDate = new Date();
                            DateFormat format = new SimpleDateFormat("yyyy.MM.dd_HH.mm.ss");

                            if(passInput.getConsoleFileType().equals("XLS")) {
                                txtSieveOutput.setText("C:\\Users\\" + System.getProperty("user.name") + "\\Desktop\\Column_Sieve_OUT_" + format.format(currentDate) + ".xls");
                                passInput.setConsoleOutFile("C:\\Users\\" + System.getProperty("user.name") + "\\Desktop\\Column_Sieve_OUT_" + format.format(currentDate) + ".xls");
                            }else if(passInput.getConsoleFileType().equals("XLSX")) {
                                txtSieveOutput.setText("C:\\Users\\" + System.getProperty("user.name") + "\\Desktop\\Column_Sieve_OUT_" + format.format(currentDate) + ".xlsx");
                                passInput.setConsoleOutFile("C:\\Users\\" + System.getProperty("user.name") + "\\Desktop\\Column_Sieve_OUT_" + format.format(currentDate) + ".xlsx");
                            }

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
                                txtConsole.append("\n> The tool will now attempt to sieve the columns from your input file.\n> Please wait...\n");
                                passInput.execute(passInput.getConsoleFileType(), this);
                            }
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
        }else if(e.getSource() == btnSieveInputBrowse){
            //sieve columns browse for input file
            fc.setDialogTitle("Please select an input file");
            String fileFilter = passInput.getConsoleFileType().toLowerCase();
            fc.setFileFilter(new FileNameExtensionFilter("Excel 97/2003 (."+fileFilter+")",fileFilter));
            fc.getCurrentDirectory();
            int returnVal = fc.showOpenDialog(ColumnSieve.this);

            if(returnVal == JFileChooser.APPROVE_OPTION){
                inFile = fc.getSelectedFile();
                fc.setCurrentDirectory(inFile);
                txtSieveInput.setText(inFile.getAbsolutePath());
            }
        }else if(e.getSource() == btnSieveTemplateBrowse){
            //sieve columns browse for template file
            fc.setDialogTitle("Please select a template file");
            String fileFilter = passInput.getConsoleFileType().toLowerCase();
            fc.setFileFilter(new FileNameExtensionFilter("Excel 97/2003 (."+fileFilter+")",fileFilter));
            fc.getCurrentDirectory();
            int returnVal = fc.showOpenDialog(ColumnSieve.this);

            if(returnVal == JFileChooser.APPROVE_OPTION){
                templateFile = fc.getSelectedFile();
                fc.setCurrentDirectory(templateFile);
                txtSieveTemplate.setText(templateFile.getAbsolutePath());
            }
        }else if(e.getSource() == btnSieveOutSave){
            //sieve columns save button for output file
            fc.setDialogTitle("Please select a location to save your new list.");
            String fileFilter = passInput.getConsoleFileType().toLowerCase();
            fc.setFileFilter(new FileNameExtensionFilter("Excel 97/2003 (."+fileFilter+")",fileFilter));
            fc.getCurrentDirectory();
            int returnVal = fc.showSaveDialog(ColumnSieve.this);

            if(returnVal == JFileChooser.APPROVE_OPTION && passInput.getConsoleFileType().equals("XLS")){
                outputFile = fc.getSelectedFile();
                if(outputFile.getAbsolutePath().substring(outputFile.getAbsolutePath().length()-4).equals(".xls")){
                    txtSieveOutput.setText(outputFile.getAbsolutePath());
                    outputFile = new File(outputFile.getAbsolutePath());
                }else if(!(outputFile.getAbsolutePath().substring(outputFile.getAbsolutePath().length()-4).equals(".xls"))){
                    txtSieveOutput.setText(outputFile.getAbsolutePath() + ".xls");
                    outputFile = new File(outputFile.getAbsolutePath() + ".xls");
                }
                fc.setCurrentDirectory(outputFile);
            }else if(returnVal == JFileChooser.APPROVE_OPTION && passInput.getConsoleFileType().equals("XLSX")){
                outputFile = fc.getSelectedFile();
                if(outputFile.getAbsolutePath().substring(outputFile.getAbsolutePath().length()-5).equals(".xlsx")){
                    txtSieveOutput.setText(outputFile.getAbsolutePath());
                    outputFile = new File(outputFile.getAbsolutePath());
                }else if(!(outputFile.getAbsolutePath().substring(outputFile.getAbsolutePath().length()-5).equals(".xlsx"))){
                    txtSieveOutput.setText(outputFile.getAbsolutePath() + ".xlsx");
                    outputFile = new File(outputFile.getAbsolutePath() + ".xlsx");
                }
                fc.setCurrentDirectory(outputFile);
            }
        }
    }
}
