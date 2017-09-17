package com.innovative.runexcelreader;

import com.innovative.excelfilereader.*;
import org.apache.commons.cli.*;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;


public class Cli {

    //private data members
    private static final Logger log = Logger.getLogger(Cli.class.getName());
    private String[] commandLineArguments = null;
    //creating option object
    public Options excelReaderOptions = new Options();

    public Cli(String[] args) {
        this.commandLineArguments = args;

        //add option: three parameters :The first parameter is a java.lang.String that represents the option.
        // The second parameter is a boolean that specifies whether the option requires an argument or not.
        //The third parameter is the description of the option
        //By specifying the argument true you need to pass the arguments compulsary at execution time.

        Option filePath = new Option("f", "filepath");
        filePath.setArgs(1);
        excelReaderOptions.addOption(filePath);

        Option worksheet = new Option("w", "worksheetNumber");
        worksheet.setArgs(1);
        excelReaderOptions.addOption(worksheet);

        Option query = new Option("q", "coloumn");
        query.setArgs(3);
        excelReaderOptions.addOption(query);

        Option help = new Option("h", "help");
        help.setArgs(1);
        excelReaderOptions.addOption(help);
    }

    /*
        * This method prints Metadata of the Excel File
        * Input: filePath : it takes file path of the Excel File
        * Output: Prints Total No of worksheets, List of worksheets, summary of all worksheets
     */
    public void MetadataDetails(String filePath) throws
            Exception {

        MetaData md = new MetaData(filePath);
        md.getDashLine(2);
        System.out.println("\t\t\t\t\t\t\t\tMETADATA");
        md.getDashLine(2);
        System.out.println("Total number of worksheets in a file:" + md.getNoOfWorksheets() + "\n");
        System.out.println("Worksheets: ");
        for (int s = 0; s < md.getSheetNames().size(); s++) {
            System.out.println(md.getSheetNames().get(s));
        }
        System.out.println();
        ArrayList<worksheet> wsList = md.getSheets();
        ArrayList<String> wsNames = md.getSheetNames();
        for (int i = 0; i < md.getNoOfWorksheets(); i++) {

            System.out.println(wsNames.get(i));
            System.out.println("Row Counts:" + wsList.get(i).getRowCounts());
            System.out.println("Colomn Counts:" + wsList.get(i).getColoumnCounts());
            System.out.print("Column DataTypes : ");

            int count = 0;
            for (String dt : wsList.get(i).getColoumnDataTypes()) {
                count++;
                System.out.print("(" + count + ")" + dt + "\t");
            }
            System.out.println();
        }
    }


    /*
        *This method displays the worksheet in table format and the summary.
        * Input: file path, worksheet number
        * Output: prints Table, Row Counts, Column Counts, Data Types of each column
     */
    public void getWorksheet(String filePath, int sheetIndex) throws Exception
    {

        try {

            ExcelReader rs = new ExcelReader();
            ExcelInfo excelInfo = rs.readExcel(filePath, sheetIndex);
            ArrayList<RowString> rows = excelInfo.getResultSet();
            rs.getDashLine(excelInfo.getColoumnCounts());

            for (int r = 0; r < rows.size(); r++) {
                RowString row = rows.get(r);
                for (int i = 0; i < row.getRow().size(); i++) {
                    System.out.format("%-35s", ExcelReader.wraptext(row.getRow().get(i)));
                    System.out.format("%-5s", "|");
                }
                System.out.println();
                if (r == 0) {
                    rs.getDashLine(excelInfo.getColoumnCounts());
                }
            }
            rs.getDashLine(excelInfo.getColoumnCounts());
            System.out.println("No of Rows :" + excelInfo.getRowCounts());
            System.out.println("No of Columns :" + excelInfo.getColoumnCounts());
            System.out.print("Column DataTypes : ");

            int count = 0;
            for (String dt : excelInfo.getColoumnDataTypes()) {
                count++;
                System.out.print("(" + count + ")" + dt + "\t");
            }
            System.out.println();

        } catch (IllegalArgumentException e) {
            System.out.println("One of the arguments is incorrect.");
            MetaData.printUsage();
        } catch (IndexOutOfBoundsException e) {
            System.out.println("No arguments are given.\n");
            MetaData.printUsage();
        }
    }

    /*
        * This method filters the given worksheet by the column and condition and prints it.
        * input:
        * @param filePath : Path of Excel file
        * @param sheetIndex: Sheet number
        * @param columnNum: column number in the worksheet
        * @operator: '=' or '<' or '>' (for String DataType, only '=')
        * @operand: any number or string
        * output: prints filtered table
     */


    public void getQueryPrint(String filePath, int sheetIndex,int columnNum, char operator, String operand)
            throws Exception {

        QueryInfo qInfo = new QueryInfo();
        ExcelReader rs = new ExcelReader();
        ExcelInfo excelInfo = qInfo.queryInfo(filePath, sheetIndex, columnNum, operator, operand);
        ArrayList<RowString> rows = excelInfo.getResultSet();
        rs.getDashLine(excelInfo.getColoumnCounts());

        for (int r = 0; r < rows.size(); r++) {
            RowString row = rows.get(r);
            for (int i = 0; i < row.getRow().size(); i++) {
                System.out.format("%-35s", ExcelReader.wraptext(row.getRow().get(i)));
                System.out.format("%-5s", "|");

            }
            System.out.println();
            if (r == 0) {
                rs.getDashLine(excelInfo.getColoumnCounts());
            }
        }
        rs.getDashLine(excelInfo.getColoumnCounts());
        System.out.println("No of Rows :" + excelInfo.getRowCounts());
        System.out.println("No of Columns :" + excelInfo.getColoumnCounts());
        System.out.print("Column DataTypes : ");

        int count = 0;
        for (String dt : excelInfo.getColoumnDataTypes()) {
            count++;
            System.out.print("(" + count + ")" + dt + "\t");
        }
        System.out.println();

    }

    //Parsing the command line argument

   public void useExcelReaderParser() {
       CommandLineParser parser = new BasicParser();
       CommandLine cmd = null;
       try {
           cmd = parser.parse(excelReaderOptions, commandLineArguments);
           if (cmd.hasOption("h")) {
               PrintHelP();
           }
           if (cmd.hasOption("f") && !cmd.hasOption("w") && !cmd.hasOption("q")) {
               System.out.println("\n=============================================");
               log.log(Level.INFO, "Selected file path -f=" + cmd.getOptionValues("f")[0]);
               MetadataDetails(cmd.getOptionValues("f")[0]);
           }

           if (cmd.hasOption("w") && !cmd.hasOption("q")) {
               if (cmd.hasOption("f")) {
                   System.out.println("\n=============================================");
                   log.log(Level.INFO, "Selected file path & worksheet -w=" + cmd.getOptionValues("w")[0]);
                   getWorksheet(cmd.getOptionValues("f")[0], Integer.parseInt(cmd.getOptionValues("w")[0]));
               } else {
                   System.out.println("\n=============================================");
                   log.log(Level.INFO, "Please provide the filepath -f=");
               }
           }

           if (cmd.hasOption("q")) {
               if (cmd.hasOption("f") && cmd.hasOption("w")) {
                   System.out.println("\n=============================================");
                   log.log(Level.INFO, "Selected file path -w=" + cmd.getOptionValues("w")[0]);
                   getQueryPrint(cmd.getOptionValues("f")[0], Integer.parseInt(cmd.getOptionValues("w")[0]), Integer.parseInt(cmd.getOptionValues("q")[0]), (cmd.getOptionValues("q")[1]).charAt(0), cmd.getOptionValues("q")[2]);
               } else {
                   System.out.println("\n=============================================");
                   log.log(Level.INFO, "Please provide the worksheetnumber -w=");
               }
           }

       } catch (ParseException parseException) {
           System.out.println("Invalid Command.Please check the usage");
           PrintHelP();
       } catch (IllegalArgumentException e) {
           System.out.println("One of the arguments is incorrect.");
       } catch (IndexOutOfBoundsException e) {
           System.out.println("Provide valid arguments.\n");
       } catch (FileNotFoundException fe) {
           System.out.println("Provide valid FilePath.\n");
       } catch (IOException io) {
           System.out.println("Provide valid FilePath.\n");
       } catch (Exception e) {
           e.printStackTrace();
       }
   }

    // Generate help information with Apache Commons CLI.
    public void PrintHelP() {
        System.out.println("\n============================================");

        HelpFormatter formatter = new HelpFormatter();
        System.out.println("\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t");
        formatter.printHelp(" ", excelReaderOptions);
        System.out.println("\n=============================================");
        System.exit(0);
    }
}
