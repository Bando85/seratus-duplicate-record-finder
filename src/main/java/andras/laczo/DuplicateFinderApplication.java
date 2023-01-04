package andras.laczo;

import andras.laczo.services.ExcelWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.util.*;

public class DuplicateFinderApplication {

    private static final String FILE_WITH_NEW_DATA_PATH = "src/main/resources/newDatasTestFile.xlsx";
    private static final String FILE_WITH_OLD_DATA_PATH = "src/main/resources/oldDatasTestFile.xlsx";
    private static final String OUTPUT_FILE_PATH = "src/main/resources";
    private static XSSFWorkbook fileWithNewData;
    private static XSSFWorkbook fileWithOldData;
    private static Short fileWithNewDataLastCellColNum;
    private static Integer fileWithNewDataLastRowNumber;


    public static void main(String[] args) throws Exception {

        initialize();
        compareHeaders();
        setLastCellColNum();
        int[] hashArray = generateHashCodesOfRows(fileWithOldData);
        //fileWithOldData.close();
        markMatchesInNewDataFileByHashCodes(hashArray);
        saveToNewWorkbook();
        //fileWithNewData.close();
    }

    private static void initialize() {
        try {

            fileWithNewData = new XSSFWorkbook(new File(FILE_WITH_NEW_DATA_PATH));
            fileWithOldData = new XSSFWorkbook(new File(FILE_WITH_OLD_DATA_PATH));
            fileWithNewDataLastRowNumber = fileWithNewData.getSheetAt(0).getLastRowNum();

        } catch (Exception e) {
            System.out.println("Exception in intializeFiles!");
        }
    }

    private static void compareHeaders() throws Exception {

        XSSFRow headerRowInNew = fileWithNewData.getSheetAt(0).getRow(0);
        XSSFRow headerRowInOld = fileWithOldData.getSheetAt(0).getRow(0);

        if (headerRowInNew.getLastCellNum() != headerRowInOld.getLastCellNum()) {
            throw new Exception("Headers dont match!");
        }

        for (int i = 0; i < headerRowInNew.getLastCellNum(); i++) {
            if (!headerRowInNew.getCell(i).getStringCellValue().equals(
                    headerRowInOld.getCell(i).getStringCellValue()
            )) throw new Exception("Headers dont match!");
        }
        System.out.println("Headers match!");
    }

    private static void setLastCellColNum() {
        fileWithNewDataLastCellColNum = fileWithNewData.getSheetAt(0).getRow(0).getLastCellNum();
    }

    private static void markMatchesInNewDataFileByHashCodes(int[] hashCodes) {
        System.out.println("Marking duplicates...");
        Arrays.sort(hashCodes);

        // start from the second row -> i=1
        for (int i = 1; i < fileWithNewDataLastRowNumber + 1; i++) {
            int rowHash;
            rowHash = generateHashCodeOfRow(fileWithNewData.getSheetAt(0).getRow(i));
            int pos = Arrays.binarySearch(hashCodes,rowHash);
            if (pos>0) {
                markRow(fileWithNewData.getSheetAt(0).getRow(i));
                hashCodes[pos]=0;
            }
        }
        System.out.println("Marking duplicates done!");
    }

    private static void markRow(XSSFRow row) {
        row.createCell(fileWithNewDataLastCellColNum + 1).setCellValue("isDuplicate");
    }

    private static String getStringValueByCellType(Cell cell) {
        String value = "";
        if (cell == null) return value;
        if (cell.getCellType().equals(CellType.NUMERIC)) {
            value = Integer.toString((int) cell.getNumericCellValue());
        }
        if (cell.getCellType().equals(CellType.STRING)) {
            value = cell.getStringCellValue();
        }
        return value;
    }

    private static int[] generateHashCodesOfRows(XSSFWorkbook file) {
        System.out.println("Generating hash codes...");
        int lastRowNum = file.getSheetAt(0).getLastRowNum();
        int[] intArray = new int[lastRowNum];
        int pos=0;
        // start from the second row -> i=1
        for (int i = 1; i < lastRowNum; i++) {
            intArray[pos] = generateHashCodeOfRow(file.getSheetAt(0).getRow(i));
            pos++;
        }
        System.out.println("Generating completed!");
        return intArray;
    }

    private static int generateHashCodeOfRow(Row row) {
        List<String> stringList = new ArrayList<>();
        for (int i = 0; i < fileWithNewDataLastCellColNum; i++) {
            stringList.add(getStringValueByCellType(row.getCell(i)));
        }
        return Objects.hash(stringList);
    }

    private static void saveToNewWorkbook() {
        System.out.println("Saving to workbook...");
        ExcelWriter.initialize(OUTPUT_FILE_PATH);
        for (int i = 0; i < fileWithNewDataLastRowNumber + 1; i++) {
            ExcelWriter.putData(fileWithNewData.getSheetAt(0).getRow(i),i);
        }
        ExcelWriter.saveWorkbook();
        System.out.println("Done!");
    }

}


