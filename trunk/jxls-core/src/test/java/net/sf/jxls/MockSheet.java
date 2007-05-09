package net.sf.jxls;

import net.sf.jxls.transformer.Sheet;

/**
 * @author Leonid Vysochyn
 */
public class MockSheet extends Sheet {
    String sheetName;

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }


}
