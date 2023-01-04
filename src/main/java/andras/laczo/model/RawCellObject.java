package andras.laczo.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.CellType;

@Builder
@AllArgsConstructor
@Getter
@Setter
public class RawCellObject {
    private String name;
    private String sheetName;
    private String workbookName;
    private int row;
    private int col;
    private Object value;
    private CellType valueType;

    public RawCellObject(int r, int c, String sname) {
        this.row = r;
        this.col = c;
        this.sheetName = sname;
    }

}
