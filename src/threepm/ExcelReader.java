package threepm;

import java.io.File;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;

import com.opencsv.CSVWriter;

/**
 *
 * @author the_fegati
 */
public class ExcelReader {

    private static File file;
    DataFormatter formatter = new DataFormatter();

    public ExcelReader(File file) {
        ExcelReader.file = file;
    }

    public void getRowAsListFromExcel() {
        List<String[]> csvList = new ArrayList<>();
        FileInputStream fis;
        Workbook workbook;
        int maxDataCount = 0;
        try {
            String fileExtension = file.toString().substring(file.toString().indexOf("."));
            OPCPackage oPCPackage = OPCPackage.open(file);
            //use xssf for xlsx format else hssf for xls format
            switch (fileExtension) {
                case ".xlsx":
                    workbook = new XSSFWorkbook(oPCPackage);
                    break;
                case ".xls":
//                    workbook = new HSSFWorkbook(new POIFSFileSystem(fis));
                    System.err.println("Wrong file type selected!");
                    return;
                default:
                    System.err.println("Wrong file type selected!");
                    return;
            }

            //get number of worksheets in the workbook
            int numberOfSheets = 1;
            String[] dataRows = new String[8];
            dataRows[0] = "dataelementUID";
            dataRows[1] = "period";
            dataRows[2] = "orgUnitUID";
            dataRows[3] = "categoryOptionComboUID";
            dataRows[4] = "ImplementingMechanismUID";
            dataRows[5] = "dataValue";
//                        System.out.println(cell);
//            System.out.printf("%4s%16s%8s%17s%17s%17s%10s\n", "", dataRows[0], dataRows[1], dataRows[2], dataRows[3], dataRows[4], dataRows[5]);

            csvList.add(dataRows);

            //iterating over each workbook sheet
            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);

                int period = (int) sheet.getRow(1).getCell(3, Row.CREATE_NULL_AS_BLANK).getNumericCellValue();
                String dataelementName = sheet.getRow(1).getCell(2, Row.CREATE_NULL_AS_BLANK).getStringCellValue();

                int numberOfRows = sheet.getLastRowNum();

                for (int row = 6; row <= numberOfRows; row++) {
                    Row currentRow = sheet.getRow(row);
                    int numberOfCells = currentRow.getLastCellNum();

                    String facility = currentRow.getCell(0, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
//                    System.out.println(facility);
//                    System.exit(0);

                    String attributeOptionCombo = currentRow.getCell(3, Row.CREATE_NULL_AS_BLANK).getStringCellValue();

                    for (int cell = 5; cell < numberOfCells; cell++) {
                        String dataelementUID = sheet.getRow(2).getCell(cell, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
                        String categoryOptionCombo = sheet.getRow(3).getCell(cell, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
                        int dataValue = 0;
                        try {
                            dataValue = (int) currentRow.getCell(cell, Row.CREATE_NULL_AS_BLANK).getNumericCellValue();
                        } catch (Exception e) {
                        }
//                        if (dataValue == 0) {
//                            continue;
//                        }
                        dataRows = new String[8];
                        dataRows[0] = String.valueOf(dataelementUID);
                        dataRows[1] = String.valueOf(period);
                        dataRows[2] = facility;
                        dataRows[3] = categoryOptionCombo;
                        dataRows[4] = attributeOptionCombo;
                        dataRows[5] = String.valueOf(dataValue);
//                        System.out.println(cell);
                        csvList.add(dataRows);
//                        Pick from here after church
                        System.out.printf("%4d%16s%8s%17s%17s%17s%10d\n", row, dataelementUID, period, facility, categoryOptionCombo, attributeOptionCombo, dataValue);

                    }
//                    if (row == 6) {
//                        System.out.println();
//                        return;
//                    }
                }
            }

            System.out.println("");

            workbook.close();
            writeRowToCSVFile(csvList);
        } catch (IOException | InvalidFormatException e) {
        }
    }

    /*
	 * Write the rows into the CSV file
     */
    private static void writeRowToCSVFile(List<String[]> cleanRows)
            throws IOException {
        File newFile = new File("/home/siech/Documents/open_heaven/intelliSOFT/3PM/3pm_V2/targets/" + file.getName().substring(0, file.getName().indexOf(".")) + ".csv");
        try (CSVWriter csvWriter = new CSVWriter(new FileWriter(newFile))) {
            csvWriter.writeAll(cleanRows);
        }
    }

}
