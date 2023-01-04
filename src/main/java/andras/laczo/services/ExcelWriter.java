package andras.laczo.services;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import java.io.File;
import java.io.FileOutputStream;


public class ExcelWriter {

    private static String outputFilePath;
    private static SXSSFWorkbook myWorkbook;

    public static void initialize(String path) {
        outputFilePath = path;
        myWorkbook = new SXSSFWorkbook(100);
        myWorkbook.createSheet();
    }

    public static void putData(XSSFRow xssfRow, int rowNumber) {
        Integer starterColPos = 0;
        SXSSFSheet newSXSSFSheet = myWorkbook.getSheetAt(0);
        SXSSFRow newSXSSFRow = newSXSSFSheet.createRow(rowNumber);

        for (int i = 0; i < xssfRow.getLastCellNum(); i++) {

            CellType cType = xssfRow.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getCellType();
            switch (cType) {
                case STRING:
                case FORMULA:
                    newSXSSFRow.getCell(starterColPos + i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                            .setCellValue(xssfRow.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue());
                    break;
                case NUMERIC:
                    newSXSSFRow.getCell(starterColPos + i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                            .setCellValue(xssfRow.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getNumericCellValue());
                    break;
            }
        }
    }

    public static void saveWorkbook() {
        try {
            File file = OutputFileCreator.createFile(new File(outputFilePath));
            FileOutputStream out = new FileOutputStream(file);
            myWorkbook.write(out);
            out.close();
            // dispose of temporary files backing this workbook on disk
            myWorkbook.dispose();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
