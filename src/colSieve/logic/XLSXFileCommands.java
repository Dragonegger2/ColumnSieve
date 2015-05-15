/*
Author: Connor Tangney
Published: 2015
*/

package colSieve.logic;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.LinkedHashMap;

public class XLSXFileCommands {
    //Declare storage maps for header values
    private LinkedHashMap<Integer,String> myHeaderVal = new LinkedHashMap<Integer,String>();
    private LinkedHashMap<Integer,String> inputHeaderVal = new LinkedHashMap<Integer,String>();
    private LinkedHashMap<Integer, String> templateHeaderVal = new LinkedHashMap<Integer,String>();
    private LinkedHashMap<Integer, String> unknownHeaderVal = new LinkedHashMap<Integer,String>();
    private LinkedHashMap<Integer, String> badHeaderVal;
    private LinkedHashMap<Integer, String> outHeaderVal;

    //Blank excel objects to store data when necessary
    private XSSFSheet mySheet, templateSheet;
    private XSSFRow currentRow, myHeader, templateHeader;
    private XSSFCell myCell, myHeaderCell, templateCell;
    private Cell headerValue, outCell, currentCell;
    private Row outRow;

    //Miscellaneous variables
    private String compareResult = "";
    private String unknownCommand = "";
    private String cellVal = "";
    private String templateCellVal = "";
    private String inFileName, templateFileName, outFileName;
    private int lastCol, lastRow, lastTemplateCol, lastTemplateRow;

    public String setHeaderRows(UserInput userInput){
        //return string
        String result = "";

        try {
            //Excel file input stream information
            FileInputStream inFile = new FileInputStream(userInput.getConsoleInFile());
            XSSFWorkbook myBook = new XSSFWorkbook(inFile);
            mySheet = myBook.getSheet(userInput.getConsoleInSheet());
            FileInputStream templateFile = new FileInputStream(userInput.getConsoleTemplateFile());
            XSSFWorkbook templateBook = new XSSFWorkbook(templateFile);
            templateSheet = templateBook.getSheet(userInput.getConsoleTemplateSheet());

            //check to make sure that the sheet exists in the workbook
            if(mySheet == null){
                result += "\t! The tool has not found sheet \"" + userInput.getConsoleInSheet() + "\".\n";
                result += "\t! Please ensure you have entered the correct sheet name before proceeding.\n";
                result += ("\t! The process has terminated abnormally.\n\n");
            } else {

                //Store excel header information
                myHeader = mySheet.getRow(0);
                templateHeader = templateSheet.getRow(0);

                //Determine the number of header cells
                lastCol = myHeader.getLastCellNum();
                lastTemplateCol = templateHeader.getLastCellNum();

                //Determine the total number of rows in the template file
                lastTemplateRow = templateSheet.getLastRowNum();

                //Get file names for use with return strings
                inFileName = userInput.getConsoleInFile();
                templateFileName = userInput.getConsoleTemplateFile();

                while (inFileName.contains("/") || inFileName.contains("\\")) {
                    if (!inFileName.contains("\\")) {
                        inFileName = inFileName.substring(inFileName.indexOf("/") + 1);
                    } else if (!inFileName.contains("/")) {
                        inFileName = inFileName.substring(inFileName.indexOf("\\") + 1);
                    }
                }

                while (templateFileName.contains("/") || templateFileName.contains("\\")) {
                    if (!templateFileName.contains("\\")) {
                        templateFileName = templateFileName.substring(templateFileName.indexOf("/") + 1);
                    } else if (!templateFileName.contains("/")) {
                        templateFileName = templateFileName.substring(templateFileName.indexOf("\\") + 1);
                    }
                }

                //Loop through inFile header values
                for (int i = 0; i < lastCol; i++) {
                    //Get cell information
                    myCell = myHeader.getCell(i);
                    if (myCell != null) {
                        cellVal = myCell.getStringCellValue();
                        myHeaderVal.put(i, cellVal);
                        inputHeaderVal.put(i, cellVal);
                    } else {
                        myHeaderVal.put(i, null);
                    }
                    templateCell = templateHeader.getCell(i);
                    if (templateCell != null) {
                        templateCellVal = templateCell.getStringCellValue();
                        templateHeaderVal.put(i, templateCellVal);
                    } else {
                        templateHeaderVal.put(i, null);
                    }
                }
                inFile.close();
                templateFile.close();

                //initialize the outHeaderVal hash map
                outHeaderVal = new LinkedHashMap<Integer, String>();

                //for each item in the headerVal
                for(int i = 0; i < myHeaderVal.size(); i++){
                    //set each outHeaderVal entry to null
                    outHeaderVal.put(i, myHeaderVal.get(i));
                }

                if(lastCol == lastTemplateCol) {
                    //Get the current console command to determine the next step
                    if (userInput.getConsoleFileCommand().equals("compareHeader")) {
                        result = compareHeader(result);
                    } else if (userInput.getConsoleFileCommand().equals("mapColumnData")) {
                        result = mapColumnData(userInput, result);
                    }
                }else if(lastCol < lastTemplateCol) {
                    userInput.smallInFile();
                }else if(lastCol > lastTemplateCol) {
                    result = userInput.longInFile(result);
                }
            }
        } catch(FileNotFoundException e){
            if(userInput.getRunMode()) {
                result += "! One or more of the files have not been found.\n";
                result += "! Please double check your file locations before trying again!\n\n";
            }else{
                result += "! One or more of the files have not been found.\n";
                result += "! Application terminated abnormally.\n\n";
                System.exit(-1);
            }
        } catch(IOException e){
            result += "! Java has encountered an IO exception.\n";
            result += "! Application terminated abnormally.\n\n";
            System.exit(-1);
        }

        return result;
    }

    public Boolean sheetExists(UserInput userInput) throws IOException{
        //returns true if both sheets exist
        //returns false if either sheet is missing
        Boolean result;
        //assume the sheet does not exist
        result = false;

        //Open both of the Excel workbooks
        FileInputStream input = new FileInputStream(userInput.getConsoleInFile());
        FileInputStream template = new FileInputStream(userInput.getConsoleTemplateFile());
        XSSFWorkbook inBook = new XSSFWorkbook(input);
        XSSFWorkbook templateBook = new XSSFWorkbook(template);

        //set the sheets
        mySheet = inBook.getSheet(userInput.getConsoleInSheet());
        templateSheet = templateBook.getSheet(userInput.getConsoleTemplateSheet());

        //if both of the sheets are not null, they exist
        if(mySheet != null && templateSheet != null){
            result = true;
        }

        input.close();
        template.close();
        return result;
    }

    public String compareHeader(String result){
        //initialize the badHeaderVal list
        badHeaderVal  = new LinkedHashMap<Integer,String>();

        //a boolean for determining whether or not an input field is known.
        //the program will assume any field is unknown UNTIL it encounters a
        //match in the template file.
        Boolean unknownBool;

        //make a deep copy of the outHeaderVal into the myHeaderVal list
        for(int i = 0; i < myHeaderVal.size(); i++){
            myHeaderVal.put(i, outHeaderVal.get(i));
        }

        //for every item in the headerVal...
        for(int i = 0; i < myHeaderVal.size(); i++){
            //Reset the unknownBool to true
            unknownBool = true;

            //if the current item in the headerVal equals the current item in the templateVal...
            if(myHeaderVal.get(i).equals(templateHeaderVal.get(i))){
                //then the current headerVal item is in the correct index
                outHeaderVal.put(i, myHeaderVal.get(i));
            }else{
                //if the current headerVal item is contained in the templateVal AT ALL...
                if(templateHeaderVal.containsValue(myHeaderVal.get(i))){
                    //for every column in the templateVal...
                    for(int j = 0; j < templateHeaderVal.size(); j++){
                        //if the current templateVal item equals the current headerVal item...
                        if(templateHeaderVal.get(j).equals(myHeaderVal.get(i))){
                            //put the current headerVal item into the outVal list at the index it was
                            //located within templateVal
                            outHeaderVal.put(j, myHeaderVal.get(i));

                            //if the headerVal item was found in a column where it was not expected...
                            if(j!=i){
                                //put it in the badHeaderVal list
                                badHeaderVal.put(j, myHeaderVal.get(i));
                            }

                            //break from the second column loop
                            break;
                        }
                    }
                }else{
                    //for every column in the template file
                    for(int j = 0; j < templateHeaderVal.size(); j++){
                        //and for every row that is not the header...
                        for(int k = 1; k <= lastTemplateRow; k++){
                            //get the current template row
                            currentRow = templateSheet.getRow(k);
                            //then get the cell from the current column
                            currentCell = currentRow.getCell(j);
                            //if the current cell is not null, and the current
                            //templateVal item equals the current headerVal item
                            if(currentCell != null && currentCell.getStringCellValue().equals(myHeaderVal.get(i))){
                                //put the current headerVal item into outVal at the current
                                //templateVal index
                                outHeaderVal.put(j, myHeaderVal.get(i));

                                //Update the templateHeaderVal text to include the extended definition
                                templateHeaderVal.put(j, myHeaderVal.get(i));

                                //if the headerVal item was found in a column where it was not expected...
                                if(j!=i) {
                                    //put it in the badHeaderVal list
                                    badHeaderVal.put(j, myHeaderVal.get(i));
                                }

                                //set the unknownBool to false, as a match was found
                                unknownBool = false;
                                //break from the row loop
                                break;
                            }else if(currentCell == null){
                                //if the currentCell is null, break from the row loop
                                break;
                            }
                        }
                        //at the completion of a column, check to see that
                        //the headerVal item is still unknown; if it is not,
                        //break from the column loop, because it is known
                        if(!unknownBool){
                            break;
                        }
                    }
                    //after the program has looped through the entire template file,
                    //if the current headerVal is still unknown, put it in the
                    //unknownHeaderVal list
                    if(unknownBool){
                        unknownHeaderVal.put(i, myHeaderVal.get(i));
                    }
                }
            }
        }

        //if the badHeaderVal has items in it
        if(badHeaderVal.size()!=0) {
            //set the compareResult to 1
            compareResult = "1";
            result += ("> The tool has determined the following fields are improperly mapped:\n\n");
            //for everything in the badHeaderVal list...
            for (int i = 0; i < lastCol; i++) {
                //if the current value is not null
                if(badHeaderVal.get(i)!=null) {
                    //print to the console
                    result += ("\t> Column Value: " + badHeaderVal.get(i) + "\n");
                }
            }
            System.out.println();
        }else if(badHeaderVal.size()==0){
            //if there is nothing in the badHeaderVal, all fields are in the correct column
            compareResult = "0";
            result += "> The tool has determined that all fields have been properly mapped.\n\n";
        }

        //if the unknownHeaderVal has items in it
        if(unknownHeaderVal.size()!=0) {
            //set the compareResult to -1
            compareResult = "-1";
            result += "> The tool has detected the following unknown field(s):\n\n";
            //for everything in the unknownHeaderVal list...
            for (int i = 0; i < lastCol; i++) {
                //if the current value is not null
                if(unknownHeaderVal.get(i)!=null) {
                    //print to the console
                    result += "\t> Column Value: " + unknownHeaderVal.get(i) + "\n";
                }
                //if the current templateVal is not null
                if(templateHeaderVal.get(i)!=null) {
                    //check to make sure the current templateVal is equal to the current outVal
                    if (!templateHeaderVal.get(i).equals(outHeaderVal.get(i))) {
                        templateHeaderVal.put(i, null);
                        outHeaderVal.put(i, null);
                    }
                } else{
                    outHeaderVal.put(i, null);
                }
            }
            System.out.println();
        }
        return result;
    }

    public String mapColumnData(UserInput userInput, String result){
        try{
            //if the headers have not been compared
            if(unknownCommand.equals("")) {
                //Compare header values
                compareHeader(result);
            }

            //if the compare result returns -1, an unknown column was found;
            //call the unknownField function to determine how to proceed
            if(compareResult.equals("-1") && userInput.getRunMode() && unknownCommand.equals("")){
                result += userInput.unknownField(result);
            }else {

                FileInputStream inputFile = new FileInputStream(userInput.getConsoleInFile());
                XSSFWorkbook inputBook = new XSSFWorkbook(inputFile);
                mySheet = inputBook.getSheet(userInput.getConsoleInSheet());
                FileInputStream templateFile = new FileInputStream(userInput.getConsoleTemplateFile());
                XSSFWorkbook templateBook = new XSSFWorkbook(templateFile);
                templateSheet = templateBook.getSheet("Sheet1");

                //Store excel header information
                templateHeader = templateSheet.getRow(0);

                //determine number of columns
                lastCol = myHeader.getLastCellNum();

                //determine number of rows
                lastRow = mySheet.getLastRowNum();

                outFileName = userInput.getConsoleOutFile();

                while (outFileName.contains("\\") || outFileName.contains("/")) {
                    if (outFileName.contains("\\")) {
                        outFileName = outFileName.substring(outFileName.indexOf("\\") + 1);
                    } else if (outFileName.contains("/")) {
                        outFileName = outFileName.substring(outFileName.indexOf("/") + 1);
                    }

                }

                String outPath = userInput.getConsoleOutFile().substring(0, (userInput.getConsoleOutFile().length() - outFileName.length()));


                if (!(outFileName.substring(outFileName.indexOf(".xls")).equals(inFileName.substring(inFileName.indexOf(".xls"))))) {
                    //If the program is running in command line mode, the program will terminate
                    if (!userInput.getRunMode()) {
                        System.out.println("! The output file type does not match the input file type.");
                        System.out.println("! Please ensure that all your file types match before trying again.");
                        System.out.println("! Application terminated abnormally.\n");
                        System.exit(-1);
                        //If the program is running in operator mode, the program will return to the cmd
                    } else {
                        result += "\n! The output file type does not match the input file type.\n! Please ensure that all your file types match before trying again.\n";
                    }
                } else {
                    //Create file objects to confirm output path / file existence
                    File myPath = new File(outPath);
                    File myFile = new File(userInput.getConsoleOutFile());

                    //Make sure output directory exists
                    //If it does not, create it
                    if (!myPath.exists()) {
                        result += "\n\t! Output directory \"" + outPath + "\" has not been found.\n\t> Directory \"" + outPath + "\" has been successfully created.\n";
                        new File(outPath).mkdirs();
                    }

                    //Make sure the output file creates
                    //If it does not, create it
                    if (!myFile.exists()) {
                        result += "\n\t! Output file \"" + outFileName + "\" has not been found.\n\t> File \"" + userInput.getConsoleOutFile() + "\" has been successfully created.\n";
                        myFile.createNewFile();
                    }

                    //Output workbook
                    Workbook outBook = new XSSFWorkbook();
                    FileOutputStream outFile = new FileOutputStream(userInput.getConsoleOutFile());
                    Sheet outSheet = outBook.createSheet(userInput.getConsoleInSheet());
                    Row outHeader = outSheet.createRow(0);

                    //Set the output sheet to contain the correct number of rows
                    for (int j = 1; j <= lastRow; j++) {
                        outSheet.createRow(j);
                    }

                    //Loop through inFile header values
                    for (int i = 0; i < lastCol; i++) {
                        ///Get cell information
                        myHeaderCell = myHeader.getCell(i);
                        String cellVal = myHeaderCell.getStringCellValue();
                        String templateCellVal = templateHeaderVal.get(i);

                        //If the input header equals the template header
                        if (cellVal.equals(templateCellVal)) {
                            //Write header to file
                            headerValue = outHeader.createCell(i);
                            headerValue.setCellValue(cellVal);

                            //Loop through all the input rows
                            for (int j = 1; j <= lastRow; j++) {
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
                                currentCell = myHeader.getCell(k);
                                if (currentCell.getStringCellValue().equals(templateCellVal)) {
                                    //Write header to file
                                    headerValue = outHeader.createCell(i);
                                    headerValue.setCellValue(currentCell.getStringCellValue());

                                    //Loop through all the input rows
                                    for (int j = 1; j <= lastRow; j++) {
                                        //Get the row data from the input file
                                        Row currentRow = mySheet.getRow(j);
                                        //Get the current cell from the row data
                                        currentCell = currentRow.getCell(k);
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
                        result += "\n> A new file has been created at the location: " + userInput.getConsoleOutFile() + "\n\n";
                    } catch (Throwable e) {
                        System.out.println("! The system encountered an error while trying to create the output file: " + userInput.getConsoleOutFile());
                        System.out.println("! Application terminated abnormally");
                        System.exit(-1);
                    }

                    //Close file streams
                    templateFile.close();
                    outFile.close();
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
        return result;
    }

    public String mapUnknownColumnToEOF(UserInput userInput, String result){
        //make sure that all null-values in the template are at the end of the list
        for(int i = 0; i < templateHeaderVal.size(); i++){
            //if the current templateHeaderVal is null...
            if(templateHeaderVal.get(i)==null){
                //get the index of the next record
                int nextRecord = i+1;
                //loop through the templateHeaderVal
                for(int k = i; k < templateHeaderVal.size(); k++){
                    //set the current templateHeaderVal to the next templateHeaderVal
                    templateHeaderVal.put(k, templateHeaderVal.get(nextRecord));
                    //get the index of the next record
                    nextRecord++;
                }
            }
        }

        //For every non-null cell in the unknownHeaderVal
        for(int i = 0; i < unknownHeaderVal.size(); i++){
            //and for every entry in the templateHeaderVal
            for(int j = 0; j < templateHeaderVal.size(); j++){
                //if the templateHeaderVal is null, it needs replaced...
                if(templateHeaderVal.get(j) == null){
                    //loop through all the values in unknownHeader (nulls included)
                    for(int k = 0; k < templateHeaderVal.size(); k++){
                        //if the current, unknownHeaderVal is not null...
                        if(unknownHeaderVal.get(k)!=null){
                            //replace the current, null templateHeaderVal
                            templateHeaderVal.put(j, unknownHeaderVal.get(k));
                            //insert the unknownHeaderVal at the end of the outHeaderVal
                            outHeaderVal.put(j, unknownHeaderVal.get(k));
                            //null the current unknownHeaderVal
                            unknownHeaderVal.put(k, null);
                            //break from the current iteration through the unknownHeaderVal
                            break;
                        }
                    }
                    //break from the current iteration through templateHeaderVal
                    break;
                }
            }
        }
        //set unknownHeaderVal to a new linked hash map; this removes all entries from the list
        unknownHeaderVal = new LinkedHashMap<Integer, String>();
        //set the unknownCommand; this prevents the tool from trying to compare the columns again
        unknownCommand = "moveToEOF";
        //send updated information back to mapColumnData
        result = mapColumnData(userInput, result);
        return result;
    }

    public String addDefinition(UserInput userInput, String result, LinkedHashMap<Integer, String> newDefinitionVal) throws IOException, InterruptedException{
        //check to make sure excel has closed
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
            }else{
                Runtime.getRuntime().exec("taskkill /im EXCEL.exe");
                excelOpen = true;
            }
        }
        tasks.destroy();

        //miscellaneous variables
        int newLastRow;

        //Re-open the template file
        FileInputStream currentTemplate = new FileInputStream(userInput.getConsoleTemplateFile());
        Workbook templateBook = new XSSFWorkbook(currentTemplate);
        Sheet newTemplateSheet = templateBook.getSheet(userInput.getConsoleInSheet());
        Row newTemplateRow;
        Cell newTemplateCell;

        //set lastNewRow equal to the number of rows in the template
        newLastRow = lastTemplateRow;

        //for every newDefinitionVal
        for(int i = 0; i < lastTemplateCol; i++){
            //check to see if the current iteration is equal to a newDefinitionVal key
            if(newDefinitionVal.containsKey(i)){
                //using the newColIndex, cycle through every templateRow until you reach
                //a null value, which indicates the end of the column
                for(int j = 0; j <=newLastRow; j++){
                    newTemplateRow = newTemplateSheet.getRow(j);
                    newTemplateCell = newTemplateRow.getCell(i);
                    //if the currentCell is null, set it to the new key / value
                    if(newTemplateCell == null){
                        //add a new cell to the correct column
                        newTemplateCell = newTemplateRow.createCell(i);
                        //set the cell type to string
                        newTemplateCell.setCellType(Cell.CELL_TYPE_STRING);
                        //set the cell value
                        newTemplateCell.setCellValue(newDefinitionVal.get(i));
                        //break from the loop
                        break;
                    }else if(newTemplateCell != null && j == newLastRow){
                        //if the last row is not a null value, the template needs a new row
                        newLastRow++;

                        //add a new row
                        newTemplateRow = newTemplateSheet.createRow(newLastRow);
                        //add a new cell to the correct column
                        newTemplateCell = newTemplateRow.createCell(i);
                        //set the cell type to string
                        newTemplateCell.setCellType(Cell.CELL_TYPE_STRING);
                        //set the cell value
                        newTemplateCell.setCellValue(newDefinitionVal.get(i));
                        //as you re-assigned the newLastRow value, break from the loop
                        break;
                    }
                }
            }
        }

        //close the input stream
        currentTemplate.close();

        //FileOutputStream to update the template file
        FileOutputStream newTemplate = new FileOutputStream(userInput.getConsoleTemplateFile());
        //Write the file
        templateBook.write(newTemplate);
        //Close the file
        newTemplate.close();
        return result;
    }

    public String deleteColumn(UserInput userInput, String result){

        //make sure that all null-values in the template are at the end of the list
        for(int i = 0; i < templateHeaderVal.size(); i++){
            //if the current templateHeaderVal is null...
            if(templateHeaderVal.get(i)==null){
                //get the index of the next record
                int nextRecord = i+1;
                //loop through the templateHeaderVal
                for(int k = i; k < templateHeaderVal.size(); k++){
                    //set the current templateHeaderVal to the next templateHeaderVal
                    templateHeaderVal.put(k, templateHeaderVal.get(nextRecord));
                    //get the index of the next record
                    nextRecord++;
                }
            }
        }

        //for every entry in the templateHeaderVal
        for(int j = 0; j < templateHeaderVal.size(); j++){
            //if the templateHeaderVal is null, it needs removed...
            if(templateHeaderVal.get(j) == null){
                //remove the current templateHeader
                templateHeaderVal.remove(j);
                //remove the current outHeaderVal
                outHeaderVal.remove(j);
                //break from the current iteration through templateHeaderVal
                break;
            }
        }
        //set unknownHeaderVal to a new linked hash map; this removes all entries from the list
        unknownHeaderVal = new LinkedHashMap<Integer, String>();
        //set the unknownCommand; this prevents the tool from trying to compare the columns again
        unknownCommand = "delete";
        //send updated information back to mapColumnData
        result = mapColumnData(userInput, result);
        return result;

    }

    public void createTemplate(String newName, UserInput userInput) throws IOException{
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
        XSSFWorkbook myBook = new XSSFWorkbook(inFile);
        XSSFSheet mySheet = myBook.getSheet(sheet);
        XSSFRow myRow = mySheet.getRow(0);
        XSSFCell myCell;
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
        Workbook outBook = new XSSFWorkbook();
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

    public LinkedHashMap<Integer, String> getUnknownHeaderVal(){
        return unknownHeaderVal;
    }

    public int getLastCol(){
        return lastCol;
    }

    public String getInputSheetName(UserInput userInput) throws IOException{
        String result = "";
        FileInputStream inFile = new FileInputStream(userInput.getConsoleInFile());
        Workbook inputBook = new XSSFWorkbook(inFile);
        result += inputBook.getSheetName(0);
        inFile.close();
        return result;
    }

    public String getTemplateSheetName(UserInput userInput) throws IOException{
        String result = "";
        FileInputStream inFile = new FileInputStream(userInput.getConsoleTemplateFile());
        XSSFWorkbook templateBook = new XSSFWorkbook(inFile);
        result += templateBook.getSheetName(0);
        inFile.close();
        return result;
    }

    public int getCompareResult(){
        return Integer.parseInt(compareResult);
    }

}