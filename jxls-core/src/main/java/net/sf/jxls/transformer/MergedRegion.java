package net.sf.jxls.transformer;

import org.apache.poi.hssf.util.Region;
import net.sf.jxls.transformer.RowCollection;

/**
 * Represents merged region
 * @author Leonid Vysochyn
 */
public class MergedRegion {
    private Region region;
    private RowCollection rowCollection;
    private int index;

    public MergedRegion(Region region, RowCollection rowCollection) {
        this.region = region;
        this.rowCollection = rowCollection;
    }

    public MergedRegion(Region region, int index) {
        this.region = region;
        this.index = index;
    }

    public int getIndex() {
        return index;
    }

    public Region getRegion() {
        return region;
    }

    public void setRegion(Region region) {
        this.region = region;
    }

    public RowCollection getRowCollection() {
        return rowCollection;
    }

    public void setRowCollection(RowCollection rowCollection) {
        this.rowCollection = rowCollection;
    }
}
