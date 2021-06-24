/**
 * @author Michael Smith
 * 6/10/2021
 */

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.*;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;

public class Actions {

    static String filename = "./src/PERSTATs/JTF PERSTAT " + LocalDate.now().getDayOfMonth() + " " + LocalDate.now().getMonth() + " " + (LocalDate.now().getYear() - 2000) + ".xlsx";
    static File file;
    static FileInputStream fip;
    static XSSFWorkbook workbook;
    static String outputFileName = "output.txt";
    static Sheet CUBSheet;
    static Sheet personnelSheet;

    static ArrayList<ServiceMember> SMsComingOnOrdersToday;
    static ArrayList<ServiceMember> SMsComingOffOrdersTomorrow;
    static ArrayList<ServiceMember> SMsComingOffOrders2Weeks;
    static ArrayList<ServiceMember> SMsOnLeave;
    static ArrayList<ServiceMember> SMsOnQuarantine;
    static ArrayList<ServiceMember> SMsOnQuarters;

    static ArrayList<ServiceMember> SMs = new ArrayList<ServiceMember>();

    public static void main(String[] args) throws IOException {

        fetchFile();
        generateEmail();
        deleteOldLeave();
        inputLeave();
        inputDTG();
        outputFile();

    }

    static void fetchFile() throws IOException {
        System.out.print("Fetching file data...");

        file = new File(filename);
        fip = new FileInputStream(file);
        workbook = new XSSFWorkbook(fip);



        if (file.isFile() && file.exists()) {

        }
        else {
            System.out.println("file doesnt exist or cannot open");
        }

        CUBSheet = workbook.getSheetAt(0);
        personnelSheet = workbook.getSheetAt(1);

        System.out.println("\t\tDone!");

    }

    static void generateEmail() throws IOException {
        File output = new File("output.txt");



        /**
         * Create output file
         */



        /**
         * Create information related variables
         */
        int totalPax = 0;
        int adminCMDCount = 0;
        int logSupportCount = 0;
        int testingCount = 0;
        int VAXDispersionSupportCount = 0;

        SMsComingOnOrdersToday = new ArrayList<ServiceMember>();
        SMsComingOffOrdersTomorrow = new ArrayList<ServiceMember>();
        SMsComingOffOrders2Weeks = new ArrayList<ServiceMember>();
        SMsOnLeave = new ArrayList<ServiceMember>();
        SMsOnQuarantine = new ArrayList<ServiceMember>();
        SMsOnQuarters = new ArrayList<ServiceMember>();

        SMs = new ArrayList<ServiceMember>();


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

        int extendingCellColumn = 10;
        int notExtendingCellColumn = 12;



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
                    formatter.formatCellValue(CellUtil.getCell(row, statusColumnIndex)).trim().toUpperCase(),
                    CellUtil.getCell(row, endDateColumnIndex)

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
                if(sm.endDateCellReference.getCellStyle().getFillForegroundColor() == CellUtil.getCell(perstatSheet.getRow(0), notExtendingCellColumn).getCellStyle().getFillForegroundColor()) {
                    SMsComingOffOrders2Weeks.add(sm);
                }
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
        System.out.print("Generating Email...");
        FileWriter writer = new FileWriter(outputFileName);
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
                        "PFC Aguino, Raquel\n" +
                        "1115th Transportation Company\n" +
                        "JTF S-1\n" +
                        "Raquel.a.aguino.mil@mail.mil\n" +
                        "\"Don't count the days. Make the days count.\"\n"
        );

        writer.close();
        System.out.println("\t\t\tDone!");
    }

    static void deleteOldLeave() {
        DataFormatter formatter = new DataFormatter();

        //Sheet CUBSheet = workbook.getSheetAt(0);

        int startRow = 16;

        Row currentRow;
        Cell nameCell;
        Cell TFCell;
        XSSFCellStyle TFCellStyle;

        int currentRowNumber;

        currentRow = CUBSheet.getRow(startRow);
        //while(!formatter.formatCellValue(CellUtil.getCell(currentRow, 1)).trim().toUpperCase().equals("QUARANTINE/QUARTERS")) {

        System.out.print("Deleting old leave...");

        while(!formatter.formatCellValue(CellUtil.getCell(CUBSheet.getRow(startRow + 1), 1)).trim().toUpperCase().equals("QUARANTINE")) {

            currentRow = CUBSheet.getRow(startRow);

            for (int j = 0; j < CUBSheet.getNumMergedRegions(); j++) {
                if (CUBSheet.getMergedRegion(j).isInRange(startRow, 1)) {
                    CUBSheet.removeMergedRegion(j);
                }

            }

            for (int j = 0; j < CUBSheet.getNumMergedRegions(); j++) {
                if (CUBSheet.getMergedRegion(j).isInRange(startRow, 7))
                    CUBSheet.removeMergedRegion(j);

            }

            CUBSheet.removeRow(currentRow);
            CUBSheet.shiftRows(startRow + 1, CUBSheet.getLastRowNum(), -1);

        }

        for (int j = 0; j < CUBSheet.getNumMergedRegions(); j++) {
            if (CUBSheet.getMergedRegion(j).isInRange(startRow, 1)) {
                CUBSheet.removeMergedRegion(j);

            }


        }

        for (int j = 0; j < CUBSheet.getNumMergedRegions(); j++) {

            if (CUBSheet.getMergedRegion(j).isInRange(startRow, 7))
                CUBSheet.removeMergedRegion(j);

        }

        System.out.println("\t\tDone!");
    }

    static void inputLeave() {
        System.out.print("Inputting Leave...");

        Row currentRow;
        Cell nameCell;
        Cell TFCell;
        XSSFCellStyle TFCellStyle;

        int startRow = 16;

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

        System.out.println("\t\t\tDone!");
    }

    static void inputDTG() {
        System.out.print("Inputting Date Time Group...");
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
        LocalDateTime now = LocalDateTime.now();
        Row currentRow = personnelSheet.getRow(1);
        CellStyle DTGCellStyle = CellUtil.getCell(personnelSheet.getRow(0), 0).getCellStyle();
        Cell DTGCell = CellUtil.getCell(currentRow, 0);
        DTGCell.setCellValue("Updated: " + now.toString());
        DTGCell.setCellStyle(DTGCellStyle);
        System.out.println("Done!");
    }

    static void outputFile() throws IOException {
        OutputStream fileOut = new FileOutputStream("./src/PERSTATs/" + file.getName());
        workbook.write(fileOut);

        System.out.println("\nDone with process check file.");
    }

}