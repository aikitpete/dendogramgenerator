package com.piglin;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

/**
 * Hello world!
 */
public class App {

    //List<String> indicators = new ArrayList<String>();
    static String[] indicators = {
            "Trustworthiness",//1
            "Attractiveness", //2
            "Reliability", //3
            "Responsibility",//4
            "Power", //5
            "Professionalism",//6
            "Expertise", //7
            "Dependability",//8
            "Honesty", //9
            "Dominance",//10
            "Confidence",//11
            "Influence",//12
            "Knowledge",//13
            "Qualification",//14
            "Awareness",//15
            "Informed",//16
            "Competence",//17
            "Intelligence"};//18

    static double[][] differences = new double[indicators.length][indicators.length];

    public static void main(String[] args) {

        try {
            System.out.println("Hello World!");

            File myFile = new File("/Users/swyna/Projects/LUND/THESIS/DifferencesData.xlsx");
            FileInputStream fis = null;

            fis = new FileInputStream(myFile);

            // Finds the workbook instance for XLSX file
            XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

            // Return first sheet from the XLSX workbook
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);

            //Workbook wb = new HSSFWorkbook(fis); //or new XSSFWorkbook("c:/temp/test.xls")
            //Sheet sheet = wb.getSheetAt(0);
            XSSFFormulaEvaluator evaluator = myWorkBook.getCreationHelper().createFormulaEvaluator();

            // Get iterator to all the rows in current sheet
            Iterator<Row> rowIterator = mySheet.iterator();

            rowIterator.next();

            // Traversing over each row of XLSX file
            //for (int i=0; i<indicators.length; i++) {
            int x=0;
            int y=0;
            int iteration=0;
            int offset=0;
            for (int i=0; i<indicators.length; i++) {
                differences[i][i] = 0;
            }
            //x=x+1;
            for (int i=0; i<153; i++) {

                x = (iteration);
                y = (iteration+1+offset);



                //for (int j=i+1; j<indicators.length; j++) {
                Row row = rowIterator.next();

                    // For each row, iterate through each columns
                    Iterator<Cell> cellIterator = row.cellIterator();
                    for (int k=0; k<6; k++) {

                        Cell cell = cellIterator.next();

                        //if (i==10||i==13||i==16||j==10||j==13||j==16) {
                        //    continue;
                        //}


                        CellValue cellValue = evaluator.evaluate(cell);
                        //System.out.print(cellValue.getNumberValue() + "\t");


                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_STRING:
                                System.out.print(cell.getStringCellValue() + "\t");
                                break;
                            case Cell.CELL_TYPE_NUMERIC:
                                System.out.print(cell.getNumericCellValue() + "\t");
                                break;
                            case Cell.CELL_TYPE_BOOLEAN:
                                System.out.print(cell.getBooleanCellValue() + "\t");
                                break;
                            case Cell.CELL_TYPE_FORMULA:
                                //CellValue cellValue = evaluator.evaluate(cell);
                                System.out.print(cellValue.getNumberValue() + "\t");
                                break;
                            default:
                                System.out.print("unsupported: "+cell.getCellType()+"\t");
                        }



                        if (k==4) {
                            differences[x][y]=cellValue.getNumberValue();
                            differences[y][x]=cellValue.getNumberValue();
                        }


                    }
                    if (y==indicators.length-1) {
                        offset = offset + 1;
                        iteration = 0;
                    } else {
                        iteration = iteration + 1;
                    }

                    System.out.print(x + "\t");
                    System.out.print(y + "\t");

                    System.out.println("");
                //}

            }

            System.out.print("Name"+",");
            for (int i=0;i<indicators.length;i++) {

                //if (i==10||i==13||i==16) {
                //    continue;
                //}

                System.out.print(indicators[i]);
                if (i!=indicators.length-1) {
                    System.out.print(",");
                }
            }
            System.out.println();
            for (int i=0;i<differences.length;i++) {

                //if (i==10||i==13||i==16) {
                //    continue;
                //}

                System.out.print(indicators[i]+",");
                for (int j=0;j<differences[i].length;j++) {

                    //if (i==10||i==13||i==16||j==10||j==13||j==16) {
                    //    continue;
                    //}

                    double value = 1-differences[i][j];
                    String number = String.valueOf(value);
                    if (number.length()==1) {
                        number = number + ".";
                    }
                    while (number.length()<5) {
                        number = number + "0";
                    }
                    System.out.print(number);
                    if (j!=differences[i].length-1) {
                        System.out.print(",");
                    }
                }
                System.out.println();
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
