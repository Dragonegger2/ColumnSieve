/*
Author: Connor Tangney
Published: 2015
*/

package colSieve.CMD;

import colSieve.logic.UserInput;
import colSieve.logic.XLSFileCommands;
import colSieve.logic.XLSXFileCommands;
import org.apache.poi.POIXMLException;
import org.apache.poi.UnsupportedFileFormatException;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;

public class Main {

    public static void main(String[] args) {
        try {
            /* ***** WELCOME ***** */
            System.out.println("\n>> Welcome to the Column Sieve tool!");
            System.out.println();

            //Initialize UserInput object / runFlag
            UserInput me = new UserInput();
            me.setRunFlag("TRUE");

            //While the runFlag is true
            while (me.getRunFlag()) {
                //If statement to check that args is empty
                if (args.length != 0) {

                    /* ***** COMMAND LINE MODE ***** */
                    System.out.println("\n>> Running in command line mode.\n");

                    //Create variable to check file type
                    String cmdFileType;

                    //Set runMode to false
                    me.setRunMode(false);

                    //Args length 3 equals call to <compareHeader>
                    if (args.length == 3 || args.length == 4) {
                        cmdFileType = args[0];
                        cmdFileType = cmdFileType.substring(cmdFileType.length() - 4, cmdFileType.length());

                        //Create the correct File Command object, then call function
                        if (cmdFileType.toLowerCase().equals(".xls")) { // <---- XLS Files
                            XLSFileCommands fileCommand = initializeXLS(me, args);
                        } else if (cmdFileType.toLowerCase().equals("xlsx")) { // <---- XLSX Files
                            XLSXFileCommands fileCommand = initializeXLSX(me, args);
                        } else if (cmdFileType.toLowerCase().equals(".csv")) {
                            //Initialize CSV file command object
                        } else if (cmdFileType.toLowerCase().equals(".txt")) {
                            //Initialize TXT file command object
                        } else {
                            System.out.println("! You have entered an unsupported file type.");
                            System.out.println("! Application terminated abnormally.");
                            System.exit(-1);
                        }

                        //End program
                        me.setRunFlag("FALSE");

                        //Args length 4 equals call to <mapColumnData>
                    } else {
                        System.out.println("\n\t! An incorrect number of arguments has been entered.");
                        System.out.println("\t! Please view the help section of the program for more information");
                        System.out.println("\t! The program will now enter operator mode.");
                        System.out.println("\t\t! Usage: ");
                        System.out.println("\t\t! compareHeader <inputFile> <inputSheetName> <templateFile>");
                        System.out.println("\t\t! mapColumnData <inputFile> <inputSheetName> <templateFile> <outputFile>");
                        System.out.println("\t> ");
                        System.out.println("\t! Application terminated abnormally.\n");
                        System.exit(-1);
                    }
                } else {
                    /* ***** OPERATOR MODE ***** */
                    System.out.println("\n>> Running in operator mode.\n");
                    //Set runMode to true
                    me.setRunMode(true);
                    //Reset variables
                    me.setConsoleFileType("NaN");
                    me.setConsoleFileCommand("NaN");
                    me.setConsoleInFile("NaN");
                    me.setConsoleInSheet();
                    me.setConsoleTemplateFile("NaN");
                    me.setConsoleOutFile("NaN");
                    me.run(me);
                }
            }

            //Redundant call to setRunFlag
            //Ensures that the program will not run again, for any reason
            me.setRunFlag("FALSE");
            System.exit(0);
        } catch (UnsupportedFileFormatException e) {
            System.out.println("! You have entered an unsupported file type.");
            System.out.println("! Please refer to the /help section for information regarding supported file types.");
            System.out.println("! Application terminated abnormally.");
            System.exit(-1);
        }
    }
    public static XLSFileCommands initializeXLS(UserInput me, String[] args) {
        try {
            XLSFileCommands fileCommand = new XLSFileCommands();
            me.setConsoleInFile(args[0]);
            me.setConsoleInSheet();
            me.setConsoleTemplateFile(args[2]);
            if (args.length == 3)
                me.setConsoleFileCommand("compareHeader");
            else if (args.length == 4) {
                me.setConsoleOutFile(args[3]);
                me.setConsoleFileCommand("mapColumnData");
            }
            fileCommand.setHeaderRows(me);
            return fileCommand;
        } catch (POIXMLException e) {
            printPOIXMLExceptionAndDie();
        } catch (OfficeXmlFileException e) {
            printOfficeXmlFileExceptionAndDie();
        }
        return null;
    }

    public static XLSXFileCommands initializeXLSX(UserInput me, String[] args) {
        try {
            XLSXFileCommands fileCommand = new XLSXFileCommands();
            me.setConsoleInFile(args[0]);
            me.setConsoleInSheet();
            me.setConsoleTemplateFile(args[2]);
            if (args.length == 3) {
                me.setConsoleFileCommand("compareHeader");
            } else if (args.length == 4) {
                me.setConsoleOutFile(args[3]);
                me.setConsoleFileCommand("mapColumnData");
            }
            fileCommand.setHeaderRows(me);
            return fileCommand;
        } catch (POIXMLException e) {
            printOfficeXmlFileExceptionAndDie();
        } catch (OfficeXmlFileException e) {
            printOfficeXmlFileExceptionAndDie();
        }
        return null;
    }

    private void printOfficeXmlFileExceptionAndDie() {
        System.out.println("! One or more of the files supplied was not in the expected file type.");
        System.out.println("! Please ensure all files are of the same format and that the proper file type is selected.");
        System.out.println("! Application terminated abnormally.\n");
        System.exit(-1);
    }

    private void printPOIXMLExceptionAndDie() {
        System.out.println("! One or more of the files supplied was not in the expected file type.");
        System.out.println("! Please ensure all files are of the same format and that the proper file type is selected.");
        System.out.println("! Application terminated abnormally.\n");
        System.exit(-1);
    }
}
