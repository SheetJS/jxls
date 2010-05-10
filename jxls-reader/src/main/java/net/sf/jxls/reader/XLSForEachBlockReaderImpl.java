package net.sf.jxls.reader;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

import org.apache.commons.jexl2.JexlContext;
import org.apache.commons.jexl2.MapContext;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

/**
 * @author Leonid Vysochyn
 */
public class XLSForEachBlockReaderImpl extends BaseBlockReader implements XLSLoopBlockReader {
    protected final Log log = LogFactory.getLog(getClass());

    String items;
    String var;
    Class varType;
    List innerBlockReaders = new ArrayList();

    SectionCheck loopBreakCheck;

    public XLSForEachBlockReaderImpl() {
    }

    public XLSForEachBlockReaderImpl(int startRow, int endRow, String items, String var, Class varType) {
        this.startRow = startRow;
        this.endRow = endRow;
        this.items = items;
        this.var = var;
        this.varType = varType;
    }

    public XLSReadStatus read(XLSRowCursor cursor, Map beans) {
        readStatus.clear();
        JexlContext context = new MapContext(beans);
        ExpressionCollectionParser parser = new ExpressionCollectionParser(context, items + ";", true);
        Collection itemsCollection = parser.getCollection();
        while (!loopBreakCheck.isCheckSuccessful(cursor)) {
            createNewCollectionItem(itemsCollection, beans);
            readInnerBlocks(cursor, beans);
        }
        cursor.moveBackward();
        return readStatus;

    }

    private void createNewCollectionItem(Collection itemsCollection, Map beans) {
        Object obj = null;
        try {
            obj = varType.newInstance();
        } catch (Exception e) {
            String message = "Can't create a new collection item for " + items + ". varType = " + varType.getName();
            readStatus.addMessage(new XLSReadMessage(message));
            if (!ReaderConfig.getInstance().isSkipErrors()) {
                readStatus.setStatusOK(false);
                throw new XLSDataReadException(message, readStatus);
            }
            if (log.isWarnEnabled()) {
                log.warn(message);
            }
        }
        itemsCollection.add(obj);
        beans.put(var, obj);
    }

    private void readInnerBlocks(XLSRowCursor cursor, Map beans) {
        for (int i = 0; i < innerBlockReaders.size(); i++) {
            XLSBlockReader xlsBlockReader = (XLSBlockReader) innerBlockReaders.get(i);
            readStatus.mergeReadStatus(xlsBlockReader.read(cursor, beans));
            cursor.moveForward();
        }
    }

    public void addBlockReader(XLSBlockReader reader) {
        innerBlockReaders.add(reader);
    }

    public List getBlockReaders() {
        return innerBlockReaders;
    }

    public SectionCheck getLoopBreakCondition() {
        return loopBreakCheck;
    }

    public void setLoopBreakCondition(SectionCheck sectionCheck) {
        this.loopBreakCheck = sectionCheck;
    }

    public void setItems(String items) {
        this.items = items;
    }

    public void setVar(String var) {
        this.var = var;
    }

    public void setVarType(Class varType) {
        this.varType = varType;
    }

    public String getItems() {
        return items;
    }

    public String getVar() {
        return var;
    }

    public Class getVarType() {
        return varType;
    }

}
