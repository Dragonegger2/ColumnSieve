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
    private String compareResult = "";

    public String compareHeader(String input, String inputSheet, String template, Boolean runMode){
        try{
            //Excel file input stream information
            FileInputStream inFile = new FileInputStream(input);
            XSSFWorkbook myBook = new XSSFWorkbook(inFile);
            XSSFSheet mySheet = myBook.getSheet(inputSheet);
            FileInputStream templateFile = new FileInputStream(template);
            XSSFWorkbook templateBook = new XSSFWorkbook(templateFile);
            XSSFSheet templateSheet = templateBook.getSheet("Sheet1");

            //Store excel header information
            XSSFRow myHeader = mySheet.getRow(0);
            XSSFRow templateHeader = templateSheet.getRow(0);

            //Determine the number of header cells
            int lastCol = myHeader.getLastCellNum();
            int lastTemplateCol = templateHeader.getLastCellNum();

            //Get file names for use with return strings
            String inFileName = input;
            String templateFileName = template;

            while(inFileName.contains("/")){
                inFileName = inFileName.substring(inFileName.indexOf("/")+1);
            }

            while(templateFileName.contains("/")){
                templateFileName = templateFileName.substring(templateFileName.indexOf("/")+1);
            }

            //If the header rows contain the same number of entries...
            if(lastCol == lastTemplateCol) {
                //Declare storage maps for header values
                LinkedHashMap<Integer,String> myHeaderVal = new LinkedHashMap<Integer,String>();
                LinkedHashMap<Integer, String> templateHeaderVal = new LinkedHashMap<Integer,String>();

                //Loop through inFile header values
                for (int i = 0; i < lastCol; i++) {
                    //Get cell information
                    XSSFCell myCell = myHeader.getCell(i);
                    String cellVal = myCell.getStringCellValue();
                    XSSFCell templateCell = templateHeader.getCell(i);
                    String templateCellVal = templateCell.getStringCellValue();

                    //Add current cell to header maps
                    myHeaderVal.put(i, cellVal);
                    templateHeaderVal.put(i, templateCellVal);
                }

                //Create storage for bad column index / value pairs
                LinkedHashMap<Integer,String> compareHeaderVal = new LinkedHashMap<Integer,String>();

                //Loop through the header maps to determine if the
                //column values are equal
                for(int i = 0; i < lastCol; i++){
                    //If the header values do not match...
                    if(!myHeaderVal.get(i).equals(templateHeaderVal.get(i))){
                        //Add the incorrect header value to bad column storage
                        compareHeaderVal.put(i,myHeaderVal.get(i));
                    }
                }

                //If the bad column storage map is not empty...
                if(compareHeaderVal.size() != 0){
                    //Set the method compareResult
                    compareResult = "1";
                    System.out.println("> The following columns from the input file \"" + inFileName + "\" are incorrectly mapped as determined by \"" + templateFileName + "\": \n");
                    for(int i = 0; i < lastCol; i++){
                        if(compareHeaderVal.get(i)!=null) {
                            System.out.println("\t> Column Index: " + i + "; Column Value: " + compareHeaderVal.get(i));
                        }
                    }
                    System.out.println();
                } else {
                    //Add string stating that file is correctly mapped to the compareResult
                    compareResult = "0";
                    System.out.println("\t> All columns from input file \"" + inFileName + "\" are in the correct location as determined by template file \"" + templateFileName + "\".\n");
                }
            } else if(lastTemplateCol < lastCol){
                if(runMode) {
                    System.out.println("! The selected input file contains more than the expected number of columns.\n");
                }else {
                    System.out.println("! The selected input file contains more than the expected number of columns.");
                    System.out.println("! Application terminated abnormally.\n");
                    System.exit(-1);
                }
            } else if(lastTemplateCol > lastCol){
                if(runMode) {
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
            if(runMode) {
                System.out.println("! One or more of the files have not been found.");
                System.out.println("! Please double check your file locations before trying again!\n");
            }else{
                System.out.println("! One or more of the files have not been found.");
                System.out.println("! Application terminated abnormally.\n");
                System.exit(-1);
            }
        } catch(IOException e){
            System.out.println("! One or more of the files have not been found.");
            System.out.println("! Application terminated abnormally.\n");
            System.exit(-1);
        }
        return compareResult;
    }

    public void mapColumnData(String input, String inputSheet, String template, String output, Boolean runMode, ColSieve userInput){
        try{
            //Create input streams
            FileInputStream inFile = new FileInputStream(input);
            XSSFWorkbook myBook = new XSSFWorkbook(inFile);
            XSSFSheet mySheet = myBook.getSheet(inputSheet);
            FileInputStream templateFile = new FileInputStream(template);
            XSSFWorkbook templateBook = new XSSFWorkbook(templateFile);
            XSSFSheet templateSheet = templateBook.getSheet("Sheet1");

            //Store excel header information
            XSSFRow myHeader = mySheet.getRow(0);
            XSSFRow templateHeader = templateSheet.getRow(0);

            //determine number of columns
            int numColumns = myHeader.getLastCellNum();

            //determine number of rows
            int numRows = mySheet.getLastRowNum();

            //Get file names for use with return strings
            String inFileName = input;
            String templateFileName = template;
            String outFileName = output;

            while(inFileName.contains("\\")){
                inFileName = inFileName.substring(inFileName.indexOf("\\")+1);
            }

            while(templateFileName.contains("\\")){
                templateFileName = templateFileName.substring(templateFileName.indexOf("\\")+1);
            }

            while(outFileName.contains("\\")){
                outFileName = outFileName.substring(outFileName.indexOf("\\")+1);
            }

            //Output file path
            String outPath = output.substring(0,(output.length()-outFileName.length()));

            //Output workbook
            Workbook outBook = new XSSFWorkbook();

            //Create file objects to confirm output path / file existence
            File myPath = new File(outPath);
            File myFile = new File(output);

            //Make sure output directory exists
            //If it does not, create it
            if(!myPath.exists()){
                System.out.println("\t! Output directory \"" + outPath + "\" has not been found.");
                new File(outPath).mkdirs();
                System.out.println("\t> Directory \"" + outPath + "\" has been successfully created.\n");
            }

            //Make sure the output file creates
            //If it does not, create it
            if(!myFile.exists()){
                System.out.println("\t! Output file \"" + outFileName + "\" has not been found.");
                myFile.createNewFile();
                System.out.println("\t> File \"" + output + "\" has been successfully created.\n");
            }

            FileOutputStream outFile = new FileOutputStream(output);
            Sheet outSheet = outBook.createSheet(inputSheet);
            Row outHeader = outSheet.createRow(0);

            //Empty Excel objects
            Cell headerValue;
            Row outRow;
            Cell outCell;

            if (compareHeader(input, inputSheet, template, runMode).contains("1")) {
                //Set the output sheet to contain the correct number of rows
                for (int j = 1; j <= numRows; j++) {
                    outSheet.createRow(j);
                }

                //Loop through inFile header values
                for (int i = 0; i < numColumns; i++) {
                    //Get cell information
                    XSSFCell myHeaderCell = myHeader.getCell(i);
                    String cellVal = myHeaderCell.getStringCellValue();
                    XSSFCell templateCell = templateHeader.getCell(i);
                    String templateCellVal = templateCell.getStringCellValue();

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
                        for (Cell currentCell : myHeader) {
                            if (currentCell.getStringCellValue().equals(templateCellVal)) {
                                //Store the correct column index
                                int inCol = currentCell.getColumnIndex();
                                int outCol = templateCell.getColumnIndex();

                                //Write header to file
                                headerValue = outHeader.createCell(outCol);
                                headerValue.setCellValue(templateCellVal);

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
                    System.out.println("\n> A new file has been created at the location: " + output + "\n");
                } catch (Throwable e) {
                    System.out.println("! The system encountered an error while trying to create the output file: " + output);
                    System.out.println("! Application terminated abnormally");
                    System.exit(-1);
                }

                //Close file streams
                inFile.close();
                templateFile.close();
                outFile.close();
            }
        } catch(FileNotFoundException e){
            System.out.println("! One of the expected files has not been found. Please ensure you have entered the correct path to your" +
                    " input file, as well as your template file.");
            System.out.println("! Application terminated abnormally");
            System.exit(-1);
        } catch(IOException e){
            System.out.println("/n! A general IO Exception has occurred while trying to process the list: " + input);
            System.out.println("! Application terminated abnormally");
            System.exit(-1);
        }
    }

    public void mapUnknownColumnToEOF(ColSieve userInput){

    }

    public void addDefinition(ColSieve userInput){

    }

    public void deleteColumn(ColSieve userInput){

    }

    public void createTemplate(String newName, ColSieve userInput) throws IOException{

    }
}