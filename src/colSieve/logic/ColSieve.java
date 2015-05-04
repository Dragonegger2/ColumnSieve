/*
Author: Connor Tangney
Published: 2015

. ColSieve class is responsible for handling all input.
. Included are things like BufferedReaders, etc. Support
. for direct input from the GUI will be implemented at a
. later date.

*/

package colSieve.logic;

import colSieve.ColumnSieve;
import org.apache.poi.POIXMLException;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import sun.awt.image.ImageWatched;

import javax.swing.*;
import java.awt.*;
import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedHashMap;

public class ColSieve {

    //BufferedReader for user input
    private BufferedReader br = new BufferedReader(new InputStreamReader(System.in));

    //Empty File Command objects
    private XLSFileCommands xlsFileCommands;
    private XLSXFileCommands xlsxFileCommands;

    //Boolean variable used to terminate the program
    private Boolean runFlag, runMode;

    //General variable declaration
    private String consoleFileType;
    private String consoleFileCommand;
    private String consoleInFile;
    private String consoleInSheet;
    private String consoleTemplateFile;
    private String consoleTemplateSheet;
    private String consoleOutFile;
    private String helpCommand;

    public void setRunFlag(String val) {
        /* ***
        Sets the program's run flag; halts the program when false.
        @PARAM - <STRING>
            -> @1 - String value which represents boolean state
        @RETURN -
        *** */

        if(val.toUpperCase().equals("TRUE")){
            this.runFlag = true;
        }else if(val.toUpperCase().equals("FALSE")){
            this.runFlag = false;
        }
    }

    public void setRunMode(Boolean val) {
        /* ***
        Sets the program's run mode; used primarily for error handling.
        @PARAM - <BOOLEAN>
            -> @1 - Boolean representing runMode value
                * 0 = command line mode
                * 1 = data operator mode
        @RETURN -
        *** */

        this.runMode = val;
    }

    public void setConsoleFileType(String val){
        /* ***
        Sets the file type captured from console input.
        @PARAM - <STRING>
            -> @1 - File type string value
        @RETURN -
        *** */

        this.consoleFileType = val;
    }

    public void setConsoleFileCommand(String val){
        /* ***
        Sets the file command captured from console input.
        @PARAM - <STRING>
            -> @1 - File command string value
        @RETURN -
        *** */

        this.consoleFileCommand = val;
    }

    public void setConsoleInFile(String val){
        /* ***
        Sets the input file captured from console.
        @PARAM - <STRING>
            -> @1 - Full path to input file, including extension
        @RETURN -
        *** */

        this.consoleInFile = val;
    }

    public void setConsoleInSheet(){
        /* ***
        Sets the input file sheet name captured from console.
        @PARAM -

        @RETURN -
        *** */
        try {
            if (consoleFileType.equals("XLS")) {
                xlsFileCommands = new XLSFileCommands();
                consoleInSheet = xlsFileCommands.getInputSheetName(this);
            } else if (consoleFileType.equals("XLSX")) {
                xlsxFileCommands = new XLSXFileCommands();
                consoleInSheet = xlsxFileCommands.getInputSheetName(this);
            }
        } catch(IOException e){
            e.printStackTrace();
        }
    }

    public void setConsoleTemplateFile(String val){
        /* ***
        Sets the template file captured from console.
        @PARAM - <STRING>
            -> @1 - Full path to template file, including extension
        @RETURN -
        *** */

        this.consoleTemplateFile = val;
    }

    public void setConsoleTemplateSheet(){
        /* ***
        Sets the input file sheet name captured from console.
        @PARAM - <STRING>
            -> @1 - Input file sheet name
        @RETURN -
        *** */
        try {
            if (consoleFileType.equals("XLS")) {
                xlsFileCommands = new XLSFileCommands();
                consoleTemplateSheet = xlsFileCommands.getTemplateSheetName(this);
            } else if (consoleFileType.equals("XLSX")) {
                xlsxFileCommands = new XLSXFileCommands();
                consoleTemplateSheet = xlsxFileCommands.getTemplateSheetName(this);
            }
        }catch(IOException e){
            e.printStackTrace();
        }
    }

    public void setConsoleOutFile(String val){
        /* ***
        Sets the output file captured from console.
        @PARAM - <STRING>
            -> @1 - Full path to output file, including extension
        @RETURN -
        *** */

        this.consoleOutFile = val;
    }

    public void setHelpCommand(String val){
        /* ***
        Sets the help command captured from console.
        @PARAM - <STRING>
            -> @1 - Captured helpCommand string
        @RETURN -
        *** */

        this.helpCommand = val;
    }

    public Boolean getRunFlag() {
        /* ***
        Accessor for private runFlag boolean
        @PARAM -

        @RETURN - <BOOLEAN>
            -> @1 - this.runFlag
        *** */

        return this.runFlag;
    }

    public Boolean getRunMode() {
        /* ***
        Accessor for private runMode boolean
        @PARAM -

        @RETURN - <BOOLEAN>
            -> @1 - this.runMode
        *** */

        return this.runMode;
    }

    public String getConsoleFileType() {
        /* ***
        Accessor for private consoleFileType string
        @PARAM -

        @RETURN - <STRING>
            -> @1 - this.consoleFileType
        *** */

        return this.consoleFileType;
    }

    public String getConsoleFileCommand() {
        /* ***
        Accessor for private consoleFileCommand string
        @PARAM -

        @RETURN - <STRING>
            -> @1 - this.consoleFileCommand
        *** */

        return this.consoleFileCommand;
    }

    public String getConsoleInFile() {
        /* ***
        Accessor for private consoleInFile string
        @PARAM -

        @RETURN - <STRING>
            -> @1 - this.consoleInFile
        *** */

        return this.consoleInFile;
    }

    public String getConsoleInSheet() {
        /* ***
        Accessor for private consoleInSheet string
        @PARAM -

        @RETURN - <STRING>
            -> @1 - this.consoleInSheet
        *** */

        return this.consoleInSheet;
    }

    public String getConsoleTemplateFile() {
        /* ***
        Accessor for private consoleTemplateFile string
        @PARAM -

        @RETURN - <STRING>
            -> @1 - this.consoleTemplateFile
        *** */

        return this.consoleTemplateFile;
    }

    public String getConsoleTemplateSheet() {
        /* ***
        Accessor for private consoleInSheet string
        @PARAM -

        @RETURN - <STRING>
            -> @1 - this.consoleInSheet
        *** */

        return this.consoleTemplateSheet;
    }

    public String getConsoleOutFile() {
        /* ***
        Accessor for private consoleOutFile string
        @PARAM -

        @RETURN - <STRING>
            -> @1 - this.consoleOutFile
        *** */

        return this.consoleOutFile;
    }

    public String getHelpCommand() {
        /* ***
        Accessor for private helpCommand string
        @PARAM -

        @RETURN - <STRING>
            -> @1 - this.helpCommand
        *** */

        return this.helpCommand;
    }

    public void run(ColSieve me) {
        /* ***
        Called from main to initialize the tool
        @PARAM - <COLSIEVE>
            -> @1 - A new UserInput object which contains information pertaining
                    to the current file.
        @RETURN -

        @EXIT -
            -> @1 - Creates a call to the proper UserInput function
        @THROWS -
            -> @1 - IOException; terminates program
        *** */

        /* ***** FILE COMMAND OPTIONS ***** */
        System.out.println("> Please enter the item number of the action you would like to perform: ");
        System.out.println(">");
        System.out.println(">\t1.\tCompare Data Layout");
        System.out.println(">\t2.\tSieve Columns");
        System.out.println(">");
        System.out.println("> Enter exit to close the application.");
        System.out.println("> Enter help for more information on any of these commands");
        System.out.print("\t> ");

        //Try block to set the command variable
        try {
            consoleFileCommand = br.readLine();

            //Logic to initialize the proper command object
            if (consoleFileCommand.equals("1")) {
                System.out.println();
                System.out.println("> You have selected: Compare Data Layout");
                consoleFileCommand = "compareHeader";
                compareDataLayout();
            } else if (consoleFileCommand.equals("2")) {
                System.out.println();
                System.out.println("> You have selected: Sieve Columns");
                consoleFileCommand = "mapColumnData";
                sieveColumns();
            } else if (consoleFileCommand.toLowerCase().equals("help")) {
                System.out.println();
                help();
            } else if (consoleFileCommand.toLowerCase().equals("exit")) {
                System.out.println();
                System.out.println("> Goodbye!");
                br.close();
                me.setRunFlag("FALSE");
            } else {
                System.out.println();
                System.out.println("> Please select an item from the list provided.");
                System.out.println();
                run(me);
            }
        } catch (IOException e) {
            System.out.println("! Java has encountered an IO exception.");
            System.out.println("! Application terminated abnormally.");
            System.exit(-1);
        }
    }

    public void help(){
        /* ***
        Called from UserInput.run(<ColSieve> <BufferedReader>) in order to initialize
        the help module.
        @PARAM -
        @RETURN -

        @EXIT -
            -> @1 - Return to main
        @THROWS -
            -> @1 - IOException; terminates program
        *** */

        //Help function for when a user requests help from the program menus
        System.out.println("> type ?help_compareHeader for more information on the Compare Data Layout function.");
        System.out.println("> type ?help_mapColumnData for more information on the Sieve Columns function.");
        System.out.println("> type ?help_cmd for more information on the tool's command line usage.");
        System.out.println("> type ?help_about for more information about the program.");
        System.out.print("\t> ");
        try {
            setHelpCommand(br.readLine());
        } catch(IOException e){
            System.out.println("! Java has encountered an IO exception.");
            System.out.println("! Application terminated abnormally.");
            System.exit(-1);
        }

        //Check the help command
        if(helpCommand.toLowerCase().equals("?help_compareheader")){
            System.out.println("\n\t> The Compare Data Layout function will confirm the column layout of a given data file.");
            System.out.println("\t> When prompted for file variables, please provide the full path to the program.");
            System.out.println("\t> If not provided with a sheet name, the program will default to \"Sheet1\".\n");
        } else if (helpCommand.toLowerCase().equals("?help_mapcolumndata")){
            System.out.println("\n\t> The Sieve Columns function will map list <X> to match the column layout in list <Y>.");
            System.out.println("\t> When prompted for file variables, please provide the full path to the program.");
            System.out.println("\t> If not provided with a sheet name, the program will default to \"Sheet1\".");
            System.out.println("\t> If not provided with an output path, the program will create a new file on the desktop by default.\n");
        }else if(helpCommand.toLowerCase().equals("?help_cmd")){
            System.out.println("\n\t> The Column Sieve tool supports command line functionality, allowing users to automate calls to the program.");
            System.out.println("\t> When running the tool in command line mode, ALL variables are required.");
            System.out.println("\t> Usage for calling the program via the command line can be found below.");
            System.out.println("\t>");
            System.out.println("\t> To call the Compare Data Layout function: ");
            System.out.println("\t\t> compareHeader <inputFile> <inputSheetName> <templateFile>");
            System.out.println("\t> ");
            System.out.println("\t> To call the Sieve Columns function: ");
            System.out.println("\t\t> mapColumnData <inputFile> <inputSheetName> <templateFile> <outputFile>\n");
        }else if(helpCommand.toLowerCase().equals("?help_about")) {
            System.out.println("\n\t> The Column Sieve tool was developed as a way for data operators to view certain elements of data files without first having to " +
                    "open the file.");
            System.out.println("\t> This is especially useful for workflows which implement automated procedures, as often times it is necessary that the data " +
                    "remain in a consistent format.");
            System.out.println("\t> For more information regarding program functionality, please view a specific topic.\n");
        }else{
            System.out.println("\n\t> Please select one of the menu items.\n");

            //Bounce back into help
            help();
        }
    }

    public void compareDataLayout(){
        /* ***
        Prompts the user to enter the necessary data for a call to FileCommands.compareHeader()
        @PARAM -
        @RETURN -

        @EXIT -
            -> @1 - Calls UserInput.fileType(<BufferedReader>)
        @THROWS -
            -> @1 - IOException; terminates program
        *** */

        try {
            //Capture necessary data
            System.out.print("\t> Enter <inputFile>: ");
            consoleInFile = br.readLine();
            System.out.print("\t> Enter <inputSheetName>: ");
            consoleInSheet = br.readLine();
            //Catch empty sheet variable
            if(consoleInSheet.equals("")){
                System.out.println("\t\t! No value for <inputSheetName> detected.");
                System.out.println("\t\t\t> File will use default: \"Sheet1\".");
                consoleInSheet = "Sheet1";
            }
            System.out.print("\t> Enter <templateFile>: ");
            consoleTemplateFile = br.readLine();
            System.out.println();

            //Call fileType() in order to proceed with the process
            fileType();
        } catch (IOException e){
            System.out.println("! Java has encountered an IO exception.");
            System.out.println("! Application terminated abnormally.");
            System.exit(-1);
        }
    }

    public void sieveColumns(){
        /* ***
        Prompts the user to enter the necessary data for a call to FileCommands.mapColumnData()
        @PARAM -
        @RETURN -

        @EXIT -
            -> @1 - Calls UserInput.fileType(<BufferedReader>)
        @THROWS -
            -> @1 - IOException; terminates program
        *** */

        try {
            //Capture necessary data
            System.out.print("\t> Enter <inputFile>: ");
            consoleInFile = br.readLine();
            System.out.print("\t> Enter <inputSheetName>: ");
            consoleInSheet = br.readLine();
            //Catch empty sheet variable
            if(consoleInSheet.equals("")){
                System.out.println("\t\t! No value for <inputSheetName> detected.");
                System.out.println("\t\t\t> File will use default: \"Sheet1\".");
                consoleInSheet = "Sheet1";
            }
            System.out.print("\t> Enter <templateFile>: ");
            consoleTemplateFile = br.readLine();
            System.out.print("\t> Enter <outputFile>: ");
            consoleOutFile = br.readLine();
            //Catch empty outfile variable
            if(consoleOutFile.equals("")){
                System.out.println("\t\t! No value for <outputFile> detected.");
                System.out.println("\t\t> File will default to output location: ");
                Date currentDate = new Date();
                DateFormat format = new SimpleDateFormat("yyyy.MM.dd_HH.mm.ss");
                System.out.println("\t\t\t> C:\\Users\\" + System.getProperty("user.name") + "\\Desktop\\Column_Sieve_OUT_" + format.format(currentDate) + ".xls");
                consoleOutFile = "C:\\Users\\" + System.getProperty("user.name") + "\\Desktop\\Column_Sieve_OUT_" + format.format(currentDate) + ".xls";
            }
            System.out.println();

            //Call fileType() in order to proceed with the process
            fileType();
        } catch (IOException e){
            System.out.println("! Java has encountered an IO exception.");
            System.out.println("! Application terminated abnormally.");
            System.exit(-1);
        }
    }

    public void fileType() {
        /* ***
        Called from ColSieve.compareDataLayout() or
        ColSieve.sieveColumns() in order to determine the
        file type the system should be using.
        the help module.
        @PARAM -
        @RETURN -

        @EXIT -
            -> @1 - Calls UserInput.execute(<String> <BufferedReader>)
        @THROWS -
            -> @1 - IOException; terminates program
        *** */

        /* ***** FILE TYPE OPTIONS ***** */
        System.out.println("> To select the proper file type, please enter the item number from the list of supported files: ");
        System.out.println(">");
        System.out.println(">\t1.\tExcel 97/2003 (.XLS)");
        System.out.println(">\t2.\tExcel (.XLSX)");
        System.out.println(">\t3.\tComma Separated (.CSV)");
        System.out.println(">\t4.\tTab Delimited (.TXT)");
        System.out.print("\t> ");

        //Try block to set the file type variable
        try {
            consoleFileType = br.readLine();
        } catch (IOException e) {
            System.out.println("! Error capturing file type.");
            System.out.println("! Application terminated abnormally.");
            System.exit(-1);
        }

        //Logic to initialize the proper file type
        if (consoleFileType.equals("1")) {
            this.consoleFileType = "XLS";
            execute(this.consoleFileType);
        } else if (consoleFileType.equals("2")) {
            this.consoleFileType = "XLSX";
            execute(this.consoleFileType);
        } else if (consoleFileType.equals("3")) {
            this.consoleFileType = "CSV";
            execute(this.consoleFileType);
        } else if (consoleFileType.equals("4")) {
            this.consoleFileType = "TXT";
            execute(this.consoleFileType);
        } else {
            System.out.println();
            System.out.println("> Please select an item from the list provided.");
            System.out.println();
            fileType();
        }
    }

    public void execute(String fileType) {
        /* ***
        Called from ColSieve.fileType() to create the correct
        FileCommand object and call for the data list.
        @PARAM - <STRING>
            -> @1 - String which represents file type.
        @RETURN -

        @EXIT -
            -> @1 - If an unsupported file type, call ColSieve.fileType()
            -> @2 - Return to main
        @THROWS -
            -> @1 - POIXMLException / OfficeXmlFileException; calls ColSieve.run()
        *** */

        try {
            //XLS Files
            if (fileType.equals("XLS")) {
                System.out.println();
                xlsFileCommands = new XLSFileCommands();
                System.out.println("\t> new XLSFileCommands object created");
                System.out.println();

                xlsFileCommands.setHeaderRows(this);

                //XLSX Files
            } else if (fileType.equals("XLSX")) {
                System.out.println();
                xlsxFileCommands = new XLSXFileCommands();
                System.out.println("\t> new XLSXFileCommands object created");
                System.out.println();

                //Determine the proper file command
                if (this.consoleFileCommand.equals("compareHeader")) {
                    xlsxFileCommands.compareHeader(this.consoleInFile, this.consoleInSheet, this.consoleTemplateFile, this.runMode);
                } else if (this.consoleFileCommand.equals("mapColumnData")) {
                    xlsxFileCommands.mapColumnData(this.consoleInFile, this.consoleInSheet, this.consoleTemplateFile, this.consoleOutFile, this.runMode, this);
                } else {
                    System.out.println("! An unknown error has occurred.");
                    System.out.println("! Application terminated abnormally.");
                }
                //CSV Files
            } else if (fileType.equals("CSV")) {
                System.out.println();
                System.out.println("> Support for CSV files has not yet been implemented. Check back later!");
                System.out.println();
                fileType();
                //TXT Files
            } else if (fileType.equals("TXT")) {
                System.out.println();
                System.out.println("> Support for TXT files has not yet been implemented. Check back later!");
                System.out.println();
                fileType();
            }
        }catch(POIXMLException e){
            System.out.println("! One or more of the files supplied was not in the expected file type.");
            System.out.println("! Please ensure all files are of the same format and that the proper file type is selected.\n");
            run(this);
        }catch(OfficeXmlFileException e){
            System.out.println("! One or more of the files supplied was not in the expected file type.");
            System.out.println("! Please ensure all files are of the same format and that the proper file type is selected.\n");
            run(this);
        }
    }

    public void execute(String fileType, ColumnSieve gui) {
        /* ***
        Called from ColSieve.GUI to create the correct
        FileCommand object and call for the data list.
        @PARAM - <STRING>
            -> @1 - String which represents file type.
            -> @2 - GUI to output data to
        @RETURN -

        @EXIT -
            -> @1 - If an unsupported file type, call ColSieve.fileType()
            -> @2 - Return to main
        @THROWS -
            -> @1 - POIXMLException / OfficeXmlFileException; calls ColSieve.run()
        *** */

        String resultString;

        try {
            //XLS Files
            if (fileType.equals("XLS")) {
                System.out.println();
                xlsFileCommands = new XLSFileCommands();
                gui.setConsole("\n\t> new XLSFileCommands object created\n\n");
                resultString = xlsFileCommands.setHeaderRows(this) + "\n\n> Waiting for input...\n";
                gui.setConsole(resultString);
                //XLSX Files
            } else if (fileType.equals("XLSX")) {
                System.out.println();
                xlsxFileCommands = new XLSXFileCommands();
                System.out.println("\t> new XLSXFileCommands object created");
                System.out.println();

                //Determine the proper file command
                if (this.consoleFileCommand.equals("compareHeader")) {
                    xlsxFileCommands.compareHeader(this.consoleInFile, this.consoleInSheet, this.consoleTemplateFile, this.runMode);
                } else if (this.consoleFileCommand.equals("mapColumnData")) {
                    xlsxFileCommands.mapColumnData(this.consoleInFile, this.consoleInSheet, this.consoleTemplateFile, this.consoleOutFile, this.runMode, this);
                } else {
                    System.out.println("! An unknown error has occurred.");
                    System.out.println("! Application terminated abnormally.");
                }
                //CSV Files
            } else if (fileType.equals("CSV")) {
                System.out.println();
                System.out.println("> Support for CSV files has not yet been implemented. Check back later!");
                System.out.println();
                fileType();
                //TXT Files
            } else if (fileType.equals("TXT")) {
                System.out.println();
                System.out.println("> Support for TXT files has not yet been implemented. Check back later!");
                System.out.println();
                fileType();
            }
        }catch(POIXMLException e){
            System.out.println("! One or more of the files supplied was not in the expected file type.");
            System.out.println("! Please ensure all files are of the same format and that the proper file type is selected.\n");
            run(this);
        }catch(OfficeXmlFileException e){
            System.out.println("! One or more of the files supplied was not in the expected file type.");
            System.out.println("! Please ensure all files are of the same format and that the proper file type is selected.\n");
            run(this);
        }
    }

    public String unknownField(String result){
        /* ***
        Called from a File Command object when an
        unknown field description has been detected
        @PARAM -
        @RETURN -

        @EXIT -

        @THROWS -
            -> @1 - IOException
        *** */
        try {
            String[] options = {"Abort current attempt","Add definition(s) to current template","Delete unknown definition(s)","Create a new template","Move unknown fields to end of file"};
            String unknownCommand = (String)JOptionPane.showInputDialog(null,"The tool has detected an unknown field definition.\nHow would you like to proceed?\n\n","Unknown Field Definition",JOptionPane.PLAIN_MESSAGE,null,options,"Select an option");

            if(unknownCommand != null) {
                if (unknownCommand.equals("1") || unknownCommand.equals("Abort current attempt")) {
                    result += "\n\n> You have opted to abort the current attempt to sieve the input columns.\n";
                } else if (unknownCommand.equals("2") || unknownCommand.equals("Add definition(s) to current template")) {
                    result += "\n\n> You have opted to add the unknown field(s) to the template file.\n";
                    if (consoleFileType.equals("XLS")) {
                        JOptionPane.showMessageDialog(null,"In a moment, the Column Sieve Tool will open your template file." +
                                "\nYou will then be prompted to enter the column index you would like your " +
                                "\nnew definition(s) to appear in." +
                                "\n\nIMPORTANT: Rows and columns in Excel files are zero-indexed!" +
                                "\nThis means that cell A1 would be represented by " +
                                "\ncoordinates [0, 0] -- Column 0, Row 0.");
                        File templateFile = new File(getConsoleTemplateFile());
                        Desktop.getDesktop().open(templateFile);
                        //existing unknown values
                        LinkedHashMap<Integer, String> unknownHeaderVal = xlsFileCommands.getUnknownHeaderVal();
                        //new locations for unknown values
                        LinkedHashMap<Integer, String> newDefinitionVal = new LinkedHashMap<Integer,String>();
                        for(int i = 0; i < xlsFileCommands.getLastCol(); i++) {
                            if (unknownHeaderVal.get(i) != null) {
                                newDefinitionVal.put(Integer.parseInt(JOptionPane.showInputDialog(null, "Please enter a column index for definition [" + unknownHeaderVal.get(i) + "]: \n")), unknownHeaderVal.get(i));
                            }
                        }
                        Runtime kill = Runtime.getRuntime();
                        Process p = kill.exec("taskkill /im EXCEL.exe");
                        p.waitFor();
                        p.destroy();

                        //get the windows tasklist
                        Process tasks = Runtime.getRuntime().exec(System.getenv("windir") + "\\system32\\tasklist.exe");

                        String line = "";
                        Boolean excelOpen = true;
                        BufferedReader taskList = new BufferedReader(new InputStreamReader(tasks.getInputStream()));
                        //while excel is open
                        while(excelOpen) {
                            //and while there is still an element in the task list
                            while (taskList.readLine() != null) {
                                //add the current task to a string dump
                                line += taskList.readLine();
                            }
                            //if the ID name EXCEL is not in the list, the program has been closed
                            if(!(line.contains("EXCEL.exe"))){
                                excelOpen = false;
                            }
                        }
                        //only add fields to the template if the file is closed.
                        if(!excelOpen) {
                            result = xlsFileCommands.addDefinition(this, result, newDefinitionVal);
                        }
                    } else if (consoleFileType.equals("XLSX")) {
                        xlsxFileCommands.addDefinition(this);
                    }
                } else if (unknownCommand.equals("3") || unknownCommand.equals("Delete unknown definition(s)")) {
                    result += "\n\n> You have opted to delete the unknown field(s).\n";
                    if (consoleFileType.equals("XLS")) {
                        result = xlsFileCommands.deleteColumn(this, result);
                    } else if (consoleFileType.equals("XLSX")) {
                        xlsxFileCommands.deleteColumn(this);
                    }
                } else if (unknownCommand.equals("4") || unknownCommand.equals("Create a new template")) {
                    result += "\n\n> You have opted to create a new template file containing the unknown field(s).\n";
                    newTemplateInfo();
                } else if (unknownCommand.equals("5") || unknownCommand.equals("Move unknown fields to end of file")) {
                    result += "\n\n> You have opted to move all field(s) to the end of the current input file.\n";
                    if (consoleFileType.equals("XLS")) {
                        result = xlsFileCommands.mapUnknownColumnToEOF(this, result);
                    } else if (consoleFileType.equals("XLSX")) {
                        xlsxFileCommands.mapUnknownColumnToEOF(this);
                    }
                } else {
                    System.out.println("\n> Please select an item from the list.\n");
                    unknownField(result);
                }
            } else {
                result += "\n\n> You have opted to abort the current attempt to sieve the input columns.\n";
            }

        }catch(IOException e){
            System.out.println("! Java has encountered an IO exception.");
            System.out.println("! Application terminated abnormally.\n");
            System.exit(-1);
        }catch(InterruptedException e){
            System.out.println("! Java has encountered an IO exception.");
            System.out.println("! Application terminated abnormally.\n");
            System.exit(-1);
        }

        return result;
    }

    public void newTemplateInfo(){
        try {
            System.out.print("\t> Please enter a new file name: ");
            String newTemplateName = br.readLine();

            if (consoleFileType.equals("XLS")) {
                if(newTemplateName.equals("")){
                    System.out.println("\t\t! No value for <newTemplateFile> detected.");
                    System.out.println("\t\t> File name will default to: ");
                    Date currentDate = new Date();
                    DateFormat format = new SimpleDateFormat("yyyy.MM.dd_HH.mm.ss");
                    System.out.println("\t\t\t> Column_Sieve_NEW_TEMPLATE_" + format.format(currentDate) + ".xls");
                    newTemplateName = "Column_Sieve_NEW_TEMPLATE_" + format.format(currentDate) + ".xls";
                }
                xlsFileCommands.createTemplate(newTemplateName,this);
            } else if (consoleFileType.equals("XLSX")) {
                if(newTemplateName.equals("")){
                    System.out.println("\t\t! No value for <newTemplateFile> detected.");
                    System.out.println("\t\t> File name will default to: ");
                    Date currentDate = new Date();
                    DateFormat format = new SimpleDateFormat("yyyy.MM.dd_HH.mm.ss");
                    System.out.println("\t\t\t> Column_Sieve_NEW_TEMPLATE_" + format.format(currentDate) + ".XLSX");
                    newTemplateName = "Column_Sieve_NEW_TEMPLATE_" + format.format(currentDate) + ".XLSX";
                }
                xlsxFileCommands.createTemplate(newTemplateName,this);
            }
        }catch(IOException e){
            System.out.println("! Java has encountered an IO exception.");
            System.out.println("! Application terminated abnormally.\n");
            System.exit(-1);
        }
    }

    public Integer newFieldDefinition(){
        String in;
        Integer newIndex = 0;

        try {
            in = br.readLine();
            newIndex = Integer.parseInt(in);
        }catch(IOException e){
            System.out.println("! Java has encountered an IO exception.");
            System.out.println("! Application terminated abnormally.\n");
            System.exit(-1);
        }
        return newIndex;
    }

    public Boolean checkSheets(){
        Boolean result = false;
        try {
            if (consoleFileType.equals("XLS")) {
                xlsFileCommands = new XLSFileCommands();
                result = xlsFileCommands.sheetExists(this);
            } else if (consoleFileType.equals("XLSX")) {
                xlsxFileCommands = new XLSXFileCommands();
                result = xlsxFileCommands.sheetExists(this);
            }
        }catch (IOException e){
            e.printStackTrace();
        }
        return result;
    }

}
