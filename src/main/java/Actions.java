/**
 * @author Michael Smith
 * 6/10/2021
 */

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.*;
import java.time.LocalDate;
import java.util.ArrayList;

public class Actions {

    public static void main(String[] args) throws IOException {


        /**
         * Get file from working directory and store it in a reference
         */
        String filename = "JTF PERSTAT " + LocalDate.now().getDayOfMonth() + " " + LocalDate.now().getMonth() + " " + (LocalDate.now().getYear() - 2000) + ".xlsx";
        File file = new File(filename);
        FileInputStream fip = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fip);
        if (file.isFile() && file.exists()) {
            System.out.println("\"" + file.getName() + "\"" + " is open");
        }
        else {
            System.out.println("file doesnt exist or cannot open");
        }


        /**
         * Create output file
         */
        File output = new File("output.txt");


        /**
         * Create information related variables
         */
        int totalPax = 0;
        int adminCMDCount = 0;
        int logSupportCount = 0;
        int testingCount = 0;
        int VAXDispersionSupportCount = 0;

        ArrayList<String> SMsComingOnOrdersToday = new ArrayList<>();
        ArrayList<String> SMsComingOffOrdersTomorrow = new ArrayList<>();
        ArrayList<String> SMsComingOffOrders2Weeks = new ArrayList<>();
        ArrayList<String> SMsOnLeave = new ArrayList<>();
        ArrayList<String> SMsOnQuarantine = new ArrayList<>();
        ArrayList<String> SMsOnQuarters = new ArrayList<>();


        /**
         * Create workbook related variable references
         */
        DataFormatter formatter = new DataFormatter();
        Sheet perstatSheet = workbook.getSheetAt(1);


        /**
         * Create cell reference specific variables
         */
        int nameColumnIndex = 2;
        int statusColumnIndex = 15;
        int ordersColumnIndex = 7;
        int rankColumnIndex = 3;
        int startDateColumnIndex = 10;
        int endDateColumnIndex = 11;
        int taskForceColumnIndex = 13;

        Cell currentNameCell;
        Cell currentStatusCell;
        Cell currentOrdersCell;
        Cell currentRankCell;
        Cell currentEndDateCell;
        Cell currentStartDateCell;
        Cell currentTaskForceCell;


        /**
         * loop through every row
         */
        for(Row row : perstatSheet) {

            /**
             * Skip header rows in workbook
             */
            if(row.getRowNum() < 3) continue;

            /**
             * Create Cell references and raw string formats of cell contents for each row
             */
            currentNameCell = CellUtil.getCell(row, nameColumnIndex);
            currentStatusCell = CellUtil.getCell(row, statusColumnIndex);
            currentOrdersCell = CellUtil.getCell(row, ordersColumnIndex);
            currentRankCell = CellUtil.getCell(row, rankColumnIndex);
            currentStartDateCell = CellUtil.getCell(row, startDateColumnIndex);
            currentEndDateCell = CellUtil.getCell(row, endDateColumnIndex);
            currentTaskForceCell = CellUtil.getCell(row, taskForceColumnIndex);

            String currentName = formatter.formatCellValue(currentNameCell).trim().toUpperCase();
            String currentStatus = formatter.formatCellValue(currentStatusCell).trim().toUpperCase();
            String currentOrders = formatter.formatCellValue(currentOrdersCell).trim().toUpperCase();
            String currentRank = formatter.formatCellValue(currentRankCell).trim().toUpperCase();
            String currentTaskForce = formatter.formatCellValue(currentTaskForceCell).trim().toUpperCase();

            String currentStartDateText = formatter.formatCellValue(currentStartDateCell).trim().toUpperCase();
            String currentEndDateText = formatter.formatCellValue(currentEndDateCell).trim().toUpperCase();


            /**
             * If SM is on orders they will count towards the total PAX count. They will also count towrds
             * a total count for each task force/mission set
             */
            if(currentOrders.equals("ON")) {
                totalPax++;

                if(currentTaskForce.equals("TOC")) adminCMDCount++;
                if(currentTaskForce.equals("MED OPS")) VAXDispersionSupportCount++;
                if(currentTaskForce.equals("RAPTOR")) testingCount++;
                if(currentTaskForce.equals("POWER")) logSupportCount++;


            }

            //TODO
            /**
             * If end date has bad formatting skip this row. Will also skip SAD people
             */
            if(currentEndDateText.indexOf('/') == -1 || currentStartDateText.indexOf('/') == -1) continue;


            /**
             * Take current date string text and store it into integers so that a LocalDate object can be created
             * from it.
             */
            int endMonth = Integer.parseInt(currentEndDateText.substring(0, currentEndDateText.indexOf('/')));
            currentEndDateText = currentEndDateText.substring(currentEndDateText.indexOf('/') + 1);
            int endDay = Integer.parseInt(currentEndDateText.substring(0, currentEndDateText.indexOf('/')));
            currentEndDateText = currentEndDateText.substring(currentEndDateText.indexOf('/') + 1);
            int endYear = Integer.parseInt(currentEndDateText);

            LocalDate currentEndDate = LocalDate.of(2000 + endYear, endMonth, endDay);

            /**
             * Take current start date string text and store it into integers so that a LocalDate object can be created
             * from it.
             */
            int startMonth = Integer.parseInt(currentStartDateText.substring(0, currentStartDateText.indexOf('/')));
            currentStartDateText = currentStartDateText.substring(currentStartDateText.indexOf('/') + 1);
            int startDay = Integer.parseInt(currentStartDateText.substring(0, currentStartDateText.indexOf('/')));
            currentStartDateText = currentStartDateText.substring(currentStartDateText.indexOf('/') + 1);
            int startYear = Integer.parseInt(currentStartDateText);

            LocalDate currentStartDate = LocalDate.of(2000 + startYear, startMonth, startDay);

            /**
             * Add SMs coming on today
             */
            if(currentStartDate.equals(LocalDate.now())) {
                SMsComingOnOrdersToday.add(currentRank + " " + currentName);
            }

            /**
             * Add SMs names to either comingoff2weeks or comingofftomorrow based on end date mutually exclusively
             */
            if(currentEndDate.equals(LocalDate.now()) || currentEndDate.isBefore(LocalDate.now()) && currentOrders.equals("ON")) {
                SMsComingOffOrdersTomorrow.add(currentRank + " " + currentName);
            }

            if(currentEndDate.minusWeeks(2).isBefore(LocalDate.now()) && currentOrders.equals("ON") && !currentEndDate.equals(LocalDate.now())) {
                SMsComingOffOrders2Weeks.add(currentRank + " " + currentName);
            }


            /**
             * Add SM to appropriate category based on status
             */
            if( currentStatus.equals("LEAVE") &
                    currentOrders.equals("ON")) {
                SMsOnLeave.add(currentRank + " " + currentName);
            }


            if( currentStatus.equals("QUARANTINE") &
                    currentOrders.equals("ON")) {
                SMsOnQuarantine.add(currentRank + " " + currentName);
            }

            if( currentStatus.equals("QUARTERS") &
                    currentOrders.equals("ON")) {
                SMsOnQuarters.add(currentRank + " " + currentName);
            }


        }


        /**
         * reference to current date
         */
        LocalDate currentDate = LocalDate.now();

        /**
         * Write to output file in the specific format for the email utilizing appropraite arraylists and variables
         * to fill in information
         */
        FileWriter writer = new FileWriter(output.getName());
        writer.write(   "ALCON,\n\n" +
                "Please see attachment for the JTF PERSTAT for " + currentDate.getDayOfMonth() + " " + currentDate.getMonth() + " " + currentDate.getYear() + "\n" +
                "Roll-up is as follows:\n\n" +
                "Total PAX: " + totalPax + "\n\n" +
                "JTF breakdown:\n\n" +

                adminCMDCount + " Admin/CMD\n" +
                logSupportCount + " Log Support\n" +
                testingCount + " Testing\n" +
                VAXDispersionSupportCount + " VAX Dispersion/Support\n\n" +

                "SM(s) coming on orders as of today:\n");

        if(SMsComingOnOrdersToday.size() <= 0) writer.write("NONE REPORTED\n");
        else for(String s : SMsComingOnOrdersToday) writer.write(s + "\n");



        writer.write(   "Total: " + SMsComingOnOrdersToday.size() + "\n\n" +
                "SM(s) coming off orders as of tomorrow:\n");

        for(String s : SMsComingOffOrdersTomorrow) writer.write(s + "\n");

        writer.write(   "Total: " + SMsComingOffOrdersTomorrow.size() + "\n\n" +
                "SM(s) coming off orders within 2 weeks\n");

        for(String s : SMsComingOffOrders2Weeks) writer.write(s + "\n");

        writer.write(   "Total:" + SMsComingOffOrders2Weeks.size() + "\n\n" +
                "SM(s) on Leave:\n");

        for(String s : SMsOnLeave) writer.write(s + "\n");

        writer.write(
                "Total: " + SMsOnLeave.size() + " \n\n" +
                        "SM(s) on Quarantine:\n");

        for(String s : SMsOnQuarantine) writer.write(s + "\n");

        writer.write(   "Total: " + SMsOnQuarantine.size() + " \n\n" +
                "SM(s) on Quarters:\n");

        for(String s : SMsOnQuarters) writer.write(s + "\n");

        writer.write(
                "Total: " + SMsOnQuarters.size() + " \n\n" +
                        "Michael W Smith\n" +
                        "2LT, IN\n" +
                        "1-200th IN D CO PLT LDR\n" +
                        "(575)499-9245\n" +
                        "p173939@nmsu.edu\n" +
                        "Smithsonian64@yahoo.com\n" +
                        "michael.w.smith910.mil@mail.mil"
        );

        writer.close();
        System.out.println("\"" + file.getName() + "\"" + " is closed");
        System.out.println("Generated Email");

        Sheet CUBSheet = workbook.getSheetAt(0);

        int startRow = 16;

        Row currentRow;
        Cell nameCell;
        Cell TFCell;
        XSSFCellStyle TFCellStyle;
        for (int i = 0; i < SMsOnLeave.size(); i++) {
            CUBSheet.shiftRows(startRow, CUBSheet.getLastRowNum(), 1, true, true);
            currentRow = CUBSheet.createRow(startRow);

            nameCell = currentRow.createCell(1);
            nameCell.setCellValue(SMsOnLeave.get(i));
            TFCell = currentRow.createCell(5);
            TFCell.setCellValue("POWER");
            TFCellStyle = workbook.createCellStyle();
            TFCellStyle.setFillBackgroundColor(new XSSFColor(Color.GREEN));
            TFCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            CUBSheet.addMergedRegion(new CellRangeAddress(currentRow.getRowNum(), currentRow.getRowNum(), 1, 1+3));

            i++;
            if(i + 1 >= SMsOnLeave.size()) break;

            nameCell = currentRow.createCell(1 + 6);
            nameCell.setCellValue(SMsOnLeave.get(i));
            TFCell = currentRow.createCell(5 + 6);
            TFCell.setCellValue("POWER");
            TFCellStyle = workbook.createCellStyle();
            TFCellStyle.setFillBackgroundColor(new XSSFColor(Color.GREEN));
            TFCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            CUBSheet.addMergedRegion(new CellRangeAddress(currentRow.getRowNum(), currentRow.getRowNum(), 1 + 6, 1+3 + 6));
        }

        OutputStream fileOut = new FileOutputStream("test" + file.getName());
        workbook.write(fileOut);




    }

}