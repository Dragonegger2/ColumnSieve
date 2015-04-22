/*
Author: Connor Tangney
Published: 2015
*/

package colSieve.logic;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.LinkedHashMap;

public class XLSFileCommands {
    //Declare storage maps for header values
    private LinkedHashMap<Integer,String> myHeaderVal = new LinkedHashMap<Integer,String>();
    private LinkedHashMap<Integer, String> templateHeaderVal = new LinkedHashMap<Integer,String>();

    //Create storage for bad column index / value pairs
    private LinkedHashMap<Integer,String> compareHeaderVal = new LinkedHashMap<Integer,String>();
    private LinkedHashMap<Integer,String> unknownHeaderVal = new LinkedHashMap<Integer,String>();

    private String compareResult = "";
    
    public String compareHeader(ColSieve userInput){
        try{
            //Excel file input stream information
            FileInputStream inFile = new FileInputStream(userInput.getConsoleInFile());
            HSSFWorkbook myBook = new HSSFWorkbook(inFile);
            HSSFSheet mySheet = myBook.getSheet(userInput.getConsoleInSheet());
            FileInputStream templateFile = new FileInputStream(userInput.getConsoleTemplateFile());
            HSSFWorkbook templateBook = new HSSFWorkbook(templateFile);
            HSSFSheet templateSheet = templateBook.getSheet("Sheet1");

            //Store excel header information
            HSSFRow myHeader = mySheet.getRow(0);
            HSSFRow templateHeader = templateSheet.getRow(0);

            //Blank excel objects to store data when necessary
            HSSFRow currentRow;
            HSSFCell currentCell;

            //Determine the number of header cells
            int lastCol = myHeader.getLastCellNum();
            int lastTemplateCol = templateHeader.getLastCellNum();

            //Determine the total number of rows in each file
            int lastTemplateRow = templateSheet.getLastRowNum();

            //Get file names for use with return strings
            String inFileName = userInput.getConsoleInFile();
            String templateFileName = userInput.getConsoleTemplateFile();

            while(inFileName.contains("/")){
                inFileName = inFileName.substring(inFileName.indexOf("/")+1);
            }

            while(templateFileName.contains("/")){
                templateFileName = templateFileName.substring(templateFileName.indexOf("/")+1);
            }

            //If the header rows contain the same number of entries...
            if(lastCol == lastTemplateCol) {
                //Loop through inFile header values
                for (int i = 0; i < lastCol; i++) {
                    //Get cell information
                    HSSFCell myCell = myHeader.getCell(i);
                    String cellVal = myCell.getStringCellValue();
                    HSSFCell templateCell = templateHeader.getCell(i);
                    String templateCellVal = templateCell.getStringCellValue();

                    //Add current cell to header maps
                    myHeaderVal.put(i, cellVal);
                    templateHeaderVal.put(i, templateCellVal);
                }

                //Loop through the header maps to determine if the
                //column values are equal
                for(int i = 0; i < lastCol; i++){
                    //If the header values do not match...
                    if(!myHeaderVal.get(i).equals(templateHeaderVal.get(i))){
                        //For each column in the file
                        for(int k = 0; k < lastTemplateCol; k++) {
                            //Loop through the entire current template column for
                            //additional data
                            for (int j = 1; j <= lastTemplateRow; j++) {
                                //Get cell(i) from the correct row
                                currentRow = templateSheet.getRow(j);
                                currentCell = currentRow.getCell(k);
                                //If current cell is empty, break from the column loop
                                if (currentCell == null) {
                                    break;
                                }
                                //If the headers match, it is a known value
                                if (currentCell.getStringCellValue().equals(myHeaderVal.get(i))) {
                                    myHeaderVal.put(i, currentCell.getStringCellValue());
                                    break;
                                }
                            }
                            //Put the updated myHeaderVal into the compareHeaderVal list
                            compareHeaderVal.put(i, myHeaderVal.get(i));
                        }
                    }
                }

                //If the bad column storage map is not empty...
                if(compareHeaderVal.size() != 0){
                    Boolean unknownBool;
                    String compareVal;
                    //Set compareResult and display all improperly mapped fields
                    compareResult = "1";
                    System.out.println("> The following columns from the input file \"" + inFileName + "\" are incorrectly mapped as determined by \"" + templateFileName + "\": \n");
                    for(int i = 0; i < lastCol; i++){
                        if(compareHeaderVal.get(i)!=null) {
                            System.out.println("\t> Column Index: " + i + "; Column Value: " + compareHeaderVal.get(i));
                        }
                    }

                    //for all columns in template
                    for(int j = 0; j < lastTemplateCol; j++){
                        if(compareHeaderVal.get(j)!=null) {
                            compareVal = compareHeaderVal.get(j);
                            //Initialize unknownBool to true; the tool will always assume
                            //a value in the compareHeaderVal map is unknown UNTIL a match is
                            //found. Value will only insert into unknownHeaderVal
                            //if the unknownBool is true.

                            //What this means:
                            //Value is in list -> unknownBool = false;
                            //Value not in list -> unknownBool = true;
                            unknownBool = true;

                            //for every row in the template file
                            for (int k = 0; k <= lastTemplateRow; k++) {
                                //Get the current row
                                currentRow = templateSheet.getRow(k);
                                //for all cells in the row
                                for (int l = 0; l < lastTemplateCol; l++) {
                                    currentCell = currentRow.getCell(l);
                                    //if the current cell equals the compareHeaderVal entry,
                                    //the item is not unknown; break from loop
                                    if (currentCell != null && compareVal.equals(currentCell.getStringCellValue())) {
                                        unknownBool = false;
                                        myHeaderVal.put(currentCell.getColumnIndex(), compareVal);
                                    }
                                    //If the value is already known, break from the current value
                                    if (!unknownBool) {
                                        break;
                                    }
                                }
                            }
                            //If a value is unknown, enter it into the unknownHeaderVal list.
                            //and remove it from the myHeaderVal list
                            if (unknownBool) {
                                unknownHeaderVal.put(j, compareVal);
                            }
                        }
                    }

                    System.out.println();

                    if(unknownHeaderVal.size()!=0){
                        compareResult = "-1";
                        System.out.println("> The tool has detected the following unknown fields:\n");
                        for(int i = 0; i < lastCol; i++){
                            if(unknownHeaderVal.get(i)!=null) {
                                System.out.println("\t> Column Value: " + unknownHeaderVal.get(i));
                            }
                        }
                        System.out.println();
                        if(!userInput.getRunMode()) {
                            System.out.println("! The input file contains unknown fields.");
                            System.out.println("! All fields must be known before attempting to map a file via the command line.");
                            System.out.println("! Application terminated abnormally.\n");
                            System.exit(-1);
                        }
                    }
                } else {
                    //Add string stating that file is correctly mapped to the compareResult
                    compareResult = "0";
                    System.out.println("\t> All columns from input file \"" + inFileName + "\" are in the correct location as determined by template file \"" + templateFileName + "\".\n");
                }
            } else if(lastTemplateCol < lastCol){
                if(userInput.getRunMode()) {
                    System.out.println("! The selected input file contains more than the expected number of columns.\n");
                }else {
                    System.out.println("! The selected input file contains more than the expected number of columns.");
                    System.out.println("! Application terminated abnormally.\n");
                    System.exit(-1);
                }
            } else if(lastTemplateCol > lastCol){
                if(userInput.getRunMode()) {
                    System.out.println("! The selected input file contains less than the expected number of columns.\n");
                }else {
                    System.out.println("! The selected input file contains more than the expected number of columns.");
                    System.out.println("! Application terminated abnormally.\n");
                    System.exit(-1);
                }
            }

            //Close both Excel files
            inFile.close();
            templateFile.close();
            //end Try block
        } catch(FileNotFoundException e){
            if(userInput.getRunMode()) {
                System.out.println("! One or more of the files have not been found.");
                System.out.println("! Please double check your file locations before trying again!\n");
            }else{
                System.out.println("! One or more of the files have not been found.");
                System.out.println("! Application terminated abnormally.\n");
                System.exit(-1);
            }
        } catch(IOException e){
            System.out.println("! Java has encountered an IO exception.");
            System.out.println("! Application terminated abnormally.\n");
            System.exit(-1);
        }
        return compareResult;
    }

    public void mapColumnData(ColSieve userInput){
        try{
            //Compare header values
            compareHeader(userInput);

            //Create input streams
            FileInputStream inFile = new FileInputStream(userInput.getConsoleInFile());
            HSSFWorkbook myBook = new HSSFWorkbook(inFile);
            HSSFSheet mySheet = myBook.getSheet(userInput.getConsoleInSheet());
            FileInputStream templateFile = new FileInputStream(userInput.getConsoleTemplateFile());
            HSSFWorkbook templateBook = new HSSFWorkbook(templateFile);
            HSSFSheet templateSheet = templateBook.getSheet("Sheet1");

            //Store excel header information
            HSSFRow myHeader = mySheet.getRow(0);
            HSSFRow templateHeader = templateSheet.getRow(0);

            //determine number of columns
            int numColumns = myHeader.getLastCellNum();

            //determine number of rows
            int numRows = mySheet.getLastRowNum();

            //Get file names for use with return strings
            String inFileName = userInput.getConsoleInFile();
            String templateFileName = userInput.getConsoleTemplateFile();
            String outFileName = userInput.getConsoleOutFile();

            while(inFileName.contains("\\")){
                inFileName = inFileName.substring(inFileName.indexOf("\\")+1);
            }

            while(templateFileName.contains("\\")){
                templateFileName = templateFileName.substring(templateFileName.indexOf("\\")+1);
            }

            while(outFileName.contains("\\")){
                outFileName = outFileName.substring(outFileName.indexOf("\\")+1);
            }

            String outPath = userInput.getConsoleOutFile().substring(0, (userInput.getConsoleOutFile().length()-outFileName.length()));


            if(!(outFileName.substring(outFileName.indexOf(".xls")).equals(inFileName.substring(inFileName.indexOf(".xls"))))){
                //If the program is running in command line mode, the program will terminate
                if(!userInput.getRunMode()) {
                    System.out.println("! The output file type does not match the input file type.");
                    System.out.println("! Please ensure that all your file types match before trying again.");
                    System.out.println("! Application terminated abnormally.\n");
                    System.exit(-1);
                    //If the program is running in operator mode, the program will return to the main
                }else{
                    System.out.println("! The output file type does not match the input file type.");
                    System.out.println("! Please ensure that all your file types match before trying again.\n");
                }
            } else {
                if (compareResult.equals("1")) {
                    //Create file objects to confirm output path / file existence
                    File myPath = new File(outPath);
                    File myFile = new File(userInput.getConsoleOutFile());

                    //Make sure output directory exists
                    //If it does not, create it
                    if (!myPath.exists()) {
                        System.out.println("\t! Output directory \"" + outPath + "\" has not been found.");
                        new File(outPath).mkdirs();
                        System.out.println("\t> Directory \"" + outPath + "\" has been successfully created.\n");
                    }

                    //Make sure the output file creates
                    //If it does not, create it
                    if (!myFile.exists()) {
                        System.out.println("\t! Output file \"" + outFileName + "\" has not been found.");
                        myFile.createNewFile();
                        System.out.println("\t> File \"" + userInput.getConsoleOutFile() + "\" has been successfully created.\n");
                    }

                    //Output workbook
                    Workbook outBook = new HSSFWorkbook();
                    FileOutputStream outFile = new FileOutputStream(userInput.getConsoleOutFile());
                    Sheet outSheet = outBook.createSheet(userInput.getConsoleInSheet());
                    Row outHeader = outSheet.createRow(0);

                    //Empty Excel objects
                    Cell headerValue;
                    Row outRow;
                    Cell outCell;

                    //Set the output sheet to contain the correct number of rows
                    for (int j = 1; j <= numRows; j++) {
                        outSheet.createRow(j);
                    }

                    //Loop through inFile header values
                    for (int i = 0; i < numColumns; i++) {
                        //Get cell information
                        HSSFCell myHeaderCell = myHeader.getCell(i);
                        String cellVal = myHeaderCell.getStringCellValue();
                        HSSFCell templateCell = templateHeader.getCell(i);
                        String templateCellVal = myHeaderVal.get(i);

                        //If the input header equals the template header
                        if (cellVal.equals(templateCellVal)) {
                            //Write header to file
                            headerValue = outHeader.createCell(i);
                            headerValue.setCellValue(cellVal);

                            //Loop through all the input rows
                            for (int j = 1; j <= numRows; j++) {
                                //Get the row data from the input file
                                Row currentRow = mySheet.getRow(j);
                                //Get the current cell from the row data
                                Cell currentCell = currentRow.getCell(i);
                                //Check to make sure current cell is not null
                                if (currentCell != null) {
                                    //Set the current cell to type: STRING
                                    currentCell.setCellType(Cell.CELL_TYPE_STRING);
                                    //Get the current row from the output file
                                    outRow = outSheet.getRow(j);
                                    //Create a new cell in the output sheet (col index I)
                                    outCell = outRow.createCell(i);
                                    //Set the outCell value to the current cell's string value
                                    outCell.setCellValue(currentCell.getStringCellValue());
                                    //If current cell is empty, print a cell with no value
                                } else {
                                    outRow = outSheet.getRow(j);
                                    outCell = outRow.createCell(i);
                                    outCell.setCellValue("");
                                }
                            }
                            //If the input header does not equal the template header
                        } else {
                            for (int k = 0; k < myHeaderVal.size(); k++) {
                                Cell currentCell = myHeader.getCell(k);
                                if (currentCell.getStringCellValue().equals(templateCellVal)) {
                                    //Store the correct column index
                                    int inCol = currentCell.getColumnIndex();
                                    int outCol = templateCell.getColumnIndex();

                                    //Write header to file
                                    headerValue = outHeader.createCell(outCol);
                                    headerValue.setCellValue(currentCell.getStringCellValue());

                                    //Loop through all the input rows
                                    for (int j = 1; j <= numRows; j++) {
                                        //Get the row data from the input file
                                        Row currentRow = mySheet.getRow(j);
                                        //Get the current cell from the row data
                                        currentCell = currentRow.getCell(inCol);
                                        //Check to make sure current cell is not null
                                        if (currentCell != null) {
                                            //Set the current cell to type: STRING
                                            currentCell.setCellType(Cell.CELL_TYPE_STRING);
                                            //Get the current row from the output file
                                            outRow = outSheet.getRow(j);
                                            //Create a new cell in the output sheet (col index I)
                                            outCell = outRow.createCell(i);
                                            //Set the outCell value to the current cell's string value
                                            outCell.setCellValue(currentCell.getStringCellValue());
                                            //If current cell is empty, print a cell with no value
                                        } else {
                                            outRow = outSheet.getRow(j);
                                            outCell = outRow.createCell(i);
                                            outCell.setCellValue("");
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //Write output file
                    try {
                        outBook.write(outFile);
                        System.out.println("> A new file has been created at the location: " + userInput.getConsoleOutFile() + "\n");
                    } catch (Throwable e) {
                        System.out.println("! The system encountered an error while trying to create the output file: " + userInput.getConsoleOutFile());
                        System.out.println("! Application terminated abnormally");
                        System.exit(-1);
                    }

                    //Close file streams
                    inFile.close();
                    templateFile.close();
                    outFile.close();
                } else if(compareResult.equals("-1") && userInput.getRunMode()){
                    userInput.unknownField();
                }
            }
        } catch(FileNotFoundException e){
            System.out.println("! One of the expected files has not been found. Please ensure you have entered the correct path to your" +
                    " input file, as well as your template file.");
            System.out.println("! Application terminated abnormally");
            System.exit(-1);
        } catch(IOException e){
            System.out.println("\n! A general IO Exception has occurred while trying to process the list: " + userInput.getConsoleInFile());
            System.out.println("! Application terminated abnormally");
            System.exit(-1);
        }
    }

    public void addDefinition(){

    }

    public void deleteColumn(){

    }

    public void createTemplate(String newName, ColSieve userInput) throws IOException{
        //Empty existing myHeaderVal data
        myHeaderVal = new LinkedHashMap<Integer, String>();

        //Set path variables
        String input = userInput.getConsoleInFile();
        String sheet = userInput.getConsoleInSheet();

        //Get the output file name
        String outName = newName;
        while(outName.contains("\\")){
            outName = outName.substring(outName.indexOf("\\")+1);
        }

        //Path variable to confirm output directory exists
        String newPath = newName.substring(0, (newName.length()-outName.length()));

        //Excel file input stream information
        FileInputStream inFile = new FileInputStream(input);
        HSSFWorkbook myBook = new HSSFWorkbook(inFile);
        HSSFSheet mySheet = myBook.getSheet(sheet);
        HSSFRow myRow = mySheet.getRow(0);
        HSSFCell myCell;
        int lastCol = myRow.getLastCellNum();

        //Get all the header cell values
        for(int i = 0; i < lastCol; i++){
            myCell = myRow.getCell(i);
            myHeaderVal.put(i,myCell.getStringCellValue());
        }

        //Validate that the output directory / file
        //exists. If they do not, create them
        File myPath = new File(newPath);
        File myFile = new File(newName);

        if (!myPath.exists()) {
            System.out.println("\t! Output directory \"" + newPath + "\" has not been found.");
            new File(newPath).mkdirs();
            System.out.println("\t> Directory \"" + newPath + "\" has been successfully created.\n");
        }

        if (!myFile.exists()) {
            System.out.println("\t! Output file \"" + newName + "\" has not been found.");
            myFile.createNewFile();
            System.out.println("\t> File \"" + newName + "\" has been successfully created.\n");
        }

        //Output workbook
        Workbook outBook = new HSSFWorkbook();
        FileOutputStream outFile = new FileOutputStream(newName);
        Sheet outSheet = outBook.createSheet(sheet);
        Row outRow = outSheet.createRow(0);
        Cell outCell;

        for(int i = 0; i < myHeaderVal.size(); i++){
            outCell = outRow.createCell(i);
            outCell.setCellValue(myHeaderVal.get(i));
        }

        outBook.write(outFile);

        outFile.close();
        inFile.close();
    }


}