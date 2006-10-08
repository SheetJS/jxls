package net.sf.jxls.tag;

import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.parser.Expression;
import net.sf.jxls.transformation.ResultTransformation;
import net.sf.jxls.transformer.Configuration;
import net.sf.jxls.transformer.SheetTransformer;
import net.sf.jxls.util.GroupData;
import net.sf.jxls.util.ReportUtil;
import net.sf.jxls.util.Util;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import java.lang.reflect.InvocationTargetException;
import java.util.Collection;
import java.util.Iterator;
import java.util.Map;

/**
 * jx:forEach tag implementation
 * @author Leonid Vysochyn
 */
public class ForEachTag extends BaseTag {
    protected final Log log = LogFactory.getLog(getClass());

    public static final String TAG_NAME = "forEach";
    Configuration configuration = new Configuration();
    static final String GROUP_DATA_KEY = "group";

    public ForEachTag() {
        name = TAG_NAME;
    }

    private String items;
    private String var;

    private String itemsKey;
    private String collectionPropertyName;
    private String groupBy;

    private String groupOrder;

    private Collection itemsCollection;


    public String getGroupOrder() {
        return groupOrder;
    }

    public void setGroupOrder(String groupOrder) {
        this.groupOrder = groupOrder;
    }

    public String getItems() {
        return items;
    }

    public void setItems(String items) {
        this.items = items;
    }

    public String getVar() {
        return var;
    }

    public void setVar(String var) {
        this.var = var;
    }

    public String getGroupBy() {
        return groupBy;
    }

    public void setGroupBy(String groupBy) {
        this.groupBy = groupBy;
    }


    public void init(TagContext tagContext) {
        super.init(tagContext);
        configuration = tagContext.getSheet().getConfiguration();
        parseItemsProperty();
        if( tagContext.getBeans().containsKey( itemsKey ) ){
            Object itemsObject = tagContext.getBeans().get( itemsKey );
            if( collectionPropertyName != null){
                itemsCollection = (Collection) Util.getProperty( itemsObject, collectionPropertyName );
            }else{
                itemsCollection = (Collection) itemsObject;
            }
        }
    }


    private void parseItemsProperty() {
        if( items != null ){
            if( items.startsWith( configuration.getStartExpressionToken() ) && items.endsWith( configuration.getEndExpressionToken())){
                items = items.substring( 2, items.length() - 1 );

                try {
                    Expression expr = new Expression(items, tagContext.getBeans(), configuration);
                    Object obj = expr.evaluate();
                    if( obj instanceof Collection ){
                        itemsCollection = (Collection) obj;
                    }else{
                        throw new RuntimeException("items property in forEach tag must be a collection. " + items + " is not ");
                    }
                } catch (Exception e) {
                    throw new RuntimeException("Can't parse an expression " + items, e);
                }
            }else{
                log.error("items attribute should start from " + configuration.getStartExpressionToken() + " and end " +
                        "with " + configuration.getEndExpressionToken());
            }
        }else{
            log.error( "Collection key is null" );
        }
    }


    public ResultTransformation process(SheetTransformer sheetTransformer) {
        if( log.isDebugEnabled() ){
            log.debug("forEach tag processing. Attributes: var = " + var + ", items=" + items);
            log.debug("Current tagContext: " + tagContext);
            log.debug("Items Collection: " + itemsCollection);
        }
        Block body = tagContext.getTagBody();
        if( body.getNumberOfRows()==1 ){
            return processOneRowTag(sheetTransformer);
        }
        int shiftNumber = 0;
        if( itemsCollection!=null && !itemsCollection.isEmpty()){
            tagContext.getSheetTransformationController().removeBorders(body);
            shiftNumber += -2; // due to the borders
            ResultTransformation shift = new ResultTransformation(0);
            Map beans = tagContext.getBeans();
            int k = 0;
            ResultTransformation processResult;
            int startRowNum, endRowNum;
//            shift.setLastProcessedRow(body.getStartRowNum() + body.getNumberOfRows() * itemsCollection.size());
            if( groupBy == null || groupBy.length() == 0 ){
                shiftNumber += tagContext.getSheetTransformationController().duplicateDown( body, itemsCollection.size() - 1 );
                for (Iterator iterator = itemsCollection.iterator(); iterator.hasNext();) {
                    Object o = iterator.next();
                    beans.put( var, o );
                    try {
                        startRowNum = body.getStartRowNum() + shift.getLastRowShift() + body.getNumberOfRows() * k++;
                        endRowNum = startRowNum + body.getNumberOfRows() - 1;
                        processResult = sheetTransformer.processRows(tagContext.getSheetTransformationController(), tagContext.getSheet(), startRowNum, endRowNum, beans, null);
                        shift.add( processResult );
                    } catch (ParsePropertyException e) {
                        log.error("Can't parse property ", e);
                    }
                }
            }else{
                try {
                    Collection groupedData = ReportUtil.groupCollectionData( itemsCollection, groupBy, groupOrder );
                    shiftNumber += tagContext.getSheetTransformationController().duplicateDown( body, groupedData.size() - 1 );
                    Object savedGroupData = null;
                    if( beans.containsKey( GROUP_DATA_KEY ) ){
                        savedGroupData = beans.get( GROUP_DATA_KEY );
                    }
                    for (Iterator iterator = groupedData.iterator(); iterator.hasNext();) {
                        GroupData groupData = (GroupData) iterator.next();
                        beans.put(GROUP_DATA_KEY, groupData );
                        try {
                            startRowNum = body.getStartRowNum() + shift.getLastRowShift() + body.getNumberOfRows() * k++;
                            endRowNum = startRowNum + body.getNumberOfRows() - 1;
                            processResult = sheetTransformer.processRows(tagContext.getSheetTransformationController(), tagContext.getSheet(), startRowNum, endRowNum, beans, null);
                            shift.add( processResult );
                        } catch (ParsePropertyException e) {
                            log.error("Can't parse property ", e);
                        }
                    }
                    beans.remove( GROUP_DATA_KEY );
                    if( savedGroupData!=null){
                        beans.put( GROUP_DATA_KEY, savedGroupData );
                    }
                } catch (NoSuchMethodException e) {
                    log.error(e, new Exception("Can't group collection data by " + groupBy, e));
                } catch (IllegalAccessException e) {
                    log.error(e, new Exception("Can't group collection data by " + groupBy, e));
                } catch (InvocationTargetException e) {
                    log.error(e, new Exception("Can't group collection data by " + groupBy, e));
                }
            }
            shift.add( new ResultTransformation(shiftNumber, shiftNumber));
//            shift.setLastProcessedRow( processResult.getLastProcessedRow() );
            shift.setTagProcessResult( true );
            return shift;
        }else{
            log.warn("Collection " + items + " is empty");
            tagContext.getSheetTransformationController().removeBodyRows( body );
            ResultTransformation shift = new ResultTransformation(0);
            shift.add( new ResultTransformation(-1, -body.getNumberOfRows() ));
            shift.setLastProcessedRow( -1 );
            shift.setTagProcessResult( true );
            return shift;
        }
    }

    private ResultTransformation processOneRowTag(SheetTransformer sheetTransformer) {
        Block body = tagContext.getTagBody();
        int shiftNumber = 0;
        tagContext.getSheetTransformationController().removeLeftRightBorders(body);
        shiftNumber += -2;
        shiftNumber += tagContext.getSheetTransformationController().duplicateRight( body, itemsCollection.size() - 1 );
        int k = 0;
        Map beans = tagContext.getBeans();
        ResultTransformation shift = new ResultTransformation();
        shift.setLastProcessedRow( -1 );
        for (Iterator iterator = itemsCollection.iterator(); iterator.hasNext();) {
            Object o = iterator.next();
            beans.put( var, o );
            try {
                short startCellNum = (short) (body.getStartCellNum() + shift.getLastCellShift() + body.getNumberOfColumns() * k++);
                short endCellNum = (short) (startCellNum + body.getNumberOfColumns() - 1);
                ResultTransformation processResult = sheetTransformer.processRow(tagContext.getSheetTransformationController(), tagContext.getSheet(),
                        tagContext.getSheet().getHssfSheet().getRow( body.getStartRowNum() ),
                        startCellNum, endCellNum, beans, null);
                shift.add( processResult );
            } catch (Exception e) {
                log.error("Can't parse property ", e);
            }
        }
        shift.addRightShift( (short) shiftNumber );
        shift.setTagProcessResult( true );
        return shift;
    }


}
