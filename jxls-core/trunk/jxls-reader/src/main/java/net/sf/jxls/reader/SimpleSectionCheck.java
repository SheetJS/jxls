package net.sf.jxls.reader;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class SimpleSectionCheck implements SectionCheck {

    List offsetRowChecks = new ArrayList();

    public SimpleSectionCheck() {
    }

    public SimpleSectionCheck(List relativeRowChecks) {
        this.offsetRowChecks = relativeRowChecks;
    }

    public boolean isCheckSuccessful(XLSRowCursor cursor) {
        for (int i = 0; i < offsetRowChecks.size(); i++) {
            OffsetRowCheck offsetRowCheck = (OffsetRowCheck) offsetRowChecks.get(i);
            if( !offsetRowCheck.isCheckSuccessful( cursor ) ){
                return false;
            }
        }
        return true;
    }

    public void addRowCheck(OffsetRowCheck offsetRowCheck) {
        offsetRowChecks.add( offsetRowCheck );
    }


    public List getOffsetRowChecks() {
        return offsetRowChecks;
    }

    public void setOffsetRowChecks(List offsetRowChecks) {
        this.offsetRowChecks = offsetRowChecks;
    }
}
