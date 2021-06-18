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
import java.util.Arrays;

public class Actions {

    public static void main(String[] args) throws IOException {


        /**
         * Get file from working directory and store it in a reference
         */

        String filename = "./src/PERSTATs/JTF PERSTAT " + LocalDate.now().getDayOfMonth() + " " + LocalDate.now().getMonth() + " " + (LocalDate.now().getYear() - 2000) + ".xlsx";
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

        ArrayList<ServiceMember> SMsComingOnOrdersToday = new ArrayList<ServiceMember>();
        ArrayList<ServiceMember> SMsComingOffOrdersTomorrow = new ArrayList<ServiceMember>();
        ArrayList<ServiceMember> SMsComingOffOrders2Weeks = new ArrayList<ServiceMember>();
        ArrayList<ServiceMember> SMsOnLeave = new ArrayList<ServiceMember>();
        ArrayList<ServiceMember> SMsOnQuarantine = new ArrayList<ServiceMember>();
        ArrayList<ServiceMember> SMsOnQuarters = new ArrayList<ServiceMember>();

        ArrayList<ServiceMember> SMs = new ArrayList<ServiceMember>();


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
        int MOSColumnIndex = 6;
        


        /**
         * loop through every row
         */
        for(Row row : perstatSheet) {

            /**
             * Skip header rows in workbook
             */
            if (row.getRowNum() < 3) continue;
            if (formatter.formatCellValue((CellUtil.getCell(row, startDateColumnIndex))).trim().toUpperCase().indexOf('-') == -1) continue;
            if (formatter.formatCellValue((CellUtil.getCell(row, endDateColumnIndex))).trim().toUpperCase().indexOf('-') == -1) continue;

            /**
             * Create Cell references and raw string formats of cell contents for each row
             */

            SMs.add(new ServiceMember(
                    formatter.formatCellValue(CellUtil.getCell(row, nameColumnIndex)).trim().toUpperCase(),
                    formatter.formatCellValue(CellUtil.getCell(row, rankColumnIndex)).trim().toUpperCase(),
                    formatter.formatCellValue(CellUtil.getCell(row, MOSColumnIndex)).trim().toUpperCase(),
                    formatter.formatCellValue(CellUtil.getCell(row, ordersColumnIndex)).trim().toUpperCase(),
                    LocalDate.parse(formatter.formatCellValue(CellUtil.getCell(row, startDateColumnIndex)).trim().toUpperCase()),
                    LocalDate.parse(formatter.formatCellValue(CellUtil.getCell(row, endDateColumnIndex)).trim().toUpperCase()),
                    formatter.formatCellValue(CellUtil.getCell(row, taskForceColumnIndex)).trim().toUpperCase(),
                    formatter.formatCellValue(CellUtil.getCell(row, statusColumnIndex)).trim().toUpperCase()
            ));

        }
        
        for(ServiceMember sm : SMs) {
            
            /**
             * If SM is on orders they will count towards the total PAX count. They will also count towards
             * a total count for each task force/mission set
             */
            if(!sm.status.equals("OFF")) {
                totalPax++;

                if(sm.taskForce.equals("TOC")) adminCMDCount++;
                if(sm.taskForce.equals("MED OPS")) VAXDispersionSupportCount++;
                if(sm.taskForce.equals("RAPTOR")) testingCount++;
                if(sm.taskForce.equals("POWER")) logSupportCount++;


            }

            //TODO
            /**
             * If end date has bad formatting skip this row. Will also skip SAD people
             */
            //if(sm..indexOf('/') == -1 || currentStartDateText.indexOf('/') == -1) continue;

            
            /**
             * Add SMs coming on today
             */
            if(sm.startDate.equals(LocalDate.now())) {
                SMsComingOnOrdersToday.add(sm);
            }

            /**
             * Add SMs names to either comingoff2weeks or comingofftomorrow based on end date mutually exclusively
             */
            if(sm.endDate.equals(LocalDate.now()) || sm.endDate.isBefore(LocalDate.now()) && !sm.status.equals("OFF")) {
                SMsComingOffOrdersTomorrow.add(sm);
            }

            if(sm.endDate.minusWeeks(2).isBefore(LocalDate.now()) && !sm.status.equals("OFF") && !sm.endDate.equals(LocalDate.now())) {
                SMsComingOffOrders2Weeks.add(sm);
            }


            /**
             * Add SM to appropriate category based on status
             */
            if( sm.status.equals("LEAVE") &
                    !sm.status.equals("OFF")) {
                SMsOnLeave.add(sm);
            }


            if( sm.status.equals("QUARANTINE") &
                    !sm.status.equals("OFF")) {
                SMsOnQuarantine.add(sm);
            }

            if( sm.status.equals("QUARTERS") &
                    !sm.status.equals("OFF")) {
                SMsOnQuarters.add(sm);
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
        else for(ServiceMember sm : SMsComingOnOrdersToday) writer.write(sm.rank + " " + sm.name + "\n");



        writer.write(   "Total: " + SMsComingOnOrdersToday.size() + "\n\n" +
                "SM(s) coming off orders as of tomorrow:\n");

        for(ServiceMember sm : SMsComingOffOrdersTomorrow) writer.write(sm.rank + " " + sm.name + "\n");

        writer.write(   "Total: " + SMsComingOffOrdersTomorrow.size() + "\n\n" +
                "SM(s) coming off orders within 2 weeks\n");

        for(ServiceMember sm : SMsComingOffOrders2Weeks) writer.write(sm.rank + " " + sm.name + "\n");

        writer.write(   "Total:" + SMsComingOffOrders2Weeks.size() + "\n\n" +
                "SM(s) on Leave:\n");

        for(ServiceMember sm : SMsOnLeave) writer.write(sm.rank + " " + sm.name + "\n");

        writer.write(
                "Total: " + SMsOnLeave.size() + " \n\n" +
                        "SM(s) on Quarantine:\n");

        for(ServiceMember sm : SMsOnQuarantine) writer.write(sm.rank + " " + sm.name + "\n");

        writer.write(   "Total: " + SMsOnQuarantine.size() + " \n\n" +
                "SM(s) on Quarters:\n");

        for(ServiceMember sm : SMsOnQuarters) writer.write(sm.rank + " " + sm.name + "\n");

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
            nameCell.setCellValue(SMsOnLeave.get(i).name);
            TFCell = currentRow.createCell(5);
            TFCell.setCellValue(SMsOnLeave.get(i).taskForce);
            TFCellStyle = workbook.createCellStyle();

            TFCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            switch (SMsOnLeave.get(i).taskForce) {
                case "TOC":
                    TFCellStyle = (XSSFCellStyle) CellUtil.getCell(CUBSheet.getRow(1), 6).getCellStyle();
                    break;
                case "MED OPS":
                    TFCellStyle = (XSSFCellStyle) CellUtil.getCell(CUBSheet.getRow(2), 6).getCellStyle();
                    break;
                case "RAPTOR":
                    TFCellStyle = (XSSFCellStyle) CellUtil.getCell(CUBSheet.getRow(3), 6).getCellStyle();
                    break;
                case "POWER":
                    TFCellStyle = (XSSFCellStyle) CellUtil.getCell(CUBSheet.getRow(4), 6).getCellStyle();
                    break;
            }

            TFCell.setCellStyle(TFCellStyle);
            CUBSheet.addMergedRegion(new CellRangeAddress(currentRow.getRowNum(), currentRow.getRowNum(), 1, 1+3));

            i++;
            if(i + 1 >= SMsOnLeave.size()) break;

            nameCell = currentRow.createCell(1 + 6);
            nameCell.setCellValue(SMsOnLeave.get(i).name);
            TFCell = currentRow.createCell(5 + 6);
            TFCell.setCellValue(SMsOnLeave.get(i).taskForce);
            TFCellStyle = workbook.createCellStyle();


            TFCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);


            switch (SMsOnLeave.get(i).taskForce) {
                case "TOC":
                    TFCellStyle = (XSSFCellStyle) CellUtil.getCell(CUBSheet.getRow(1), 6).getCellStyle();
                    break;
                case "MED OPS":
                    TFCellStyle = (XSSFCellStyle) CellUtil.getCell(CUBSheet.getRow(2), 6).getCellStyle();
                    break;
                case "RAPTOR":
                    TFCellStyle = (XSSFCellStyle) CellUtil.getCell(CUBSheet.getRow(3), 6).getCellStyle();
                    break;
                case "POWER":
                    TFCellStyle = (XSSFCellStyle) CellUtil.getCell(CUBSheet.getRow(4), 6).getCellStyle();
                    break;
            }


            TFCell.setCellStyle(TFCellStyle);



            CUBSheet.addMergedRegion(new CellRangeAddress(currentRow.getRowNum(), currentRow.getRowNum(), 1 + 6, 1+3 + 6));
        }

        OutputStream fileOut = new FileOutputStream("./src/PERSTATs/Output " + file.getName());
        workbook.write(fileOut);




    }

}