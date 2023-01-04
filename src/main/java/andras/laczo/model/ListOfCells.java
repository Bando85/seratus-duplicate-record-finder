package andras.laczo.model;

import andras.laczo.services.ExcelColumn;

import java.util.LinkedList;

public class ListOfCells extends LinkedList<RawCellObject> {

    public void addCell(int cellRow, String cellCol, String sheetName) {
        this.add(new RawCellObject(cellRow - 1, ExcelColumn.toNumber(cellCol) - 1, sheetName));
    }

}
