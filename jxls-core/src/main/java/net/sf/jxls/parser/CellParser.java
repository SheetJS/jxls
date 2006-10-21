package net.sf.jxls.parser;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.commons.digester.Digester;
import org.xml.sax.SAXException;
import net.sf.jxls.util.Util;
import net.sf.jxls.tag.Tag;
import net.sf.jxls.tag.TagContext;
import net.sf.jxls.tag.Taglib;
import net.sf.jxls.tag.Block;
import net.sf.jxls.formula.Formula;
import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.transformer.Configuration;
import net.sf.jxls.transformer.Row;

import java.util.Map;
import java.util.Set;
import java.util.Iterator;
import java.io.StringReader;
import java.io.IOException;

/**
 * Class for parsing excel cell
 * @author Leonid Vysochyn
 */
public class CellParser {
    protected final Log log = LogFactory.getLog(getClass());

    private final Cell cell;

    private Configuration configuration;

    public CellParser(HSSFCell hssfCell, Row row, Configuration configuration) {
        this.cell = new Cell( hssfCell, row );
        if( configuration!=null ){
            this.configuration = configuration;
        }else{
            this.configuration = new Configuration();
        }
    }



    public CellParser(Cell cell) {
        this.cell = cell;
    }

    public Cell getCell() {
        return cell;
    }

    public Cell parseCell(Map beans){
        if (cell.getHssfCell() != null) {
            try {
                if( cell.getHssfCell().getCellType() == HSSFCell.CELL_TYPE_STRING ){
                    cell.setHssfCellValue(cell.getHssfCell().getStringCellValue());
                    parseCellValue( beans);
                }
            } catch (ParsePropertyException e) {
                log.error("Can't get value for property=" + cell.getCollectionProperty().getProperty(), e);
                throw new RuntimeException(e);
            }
            updateMergedRegions();
        }
        return cell;
    }

    public Formula parseCellFormula(){
        if( cell.getHssfCell() != null && (cell.getHssfCell().getCellType() == HSSFCell.CELL_TYPE_STRING) && cell.getHssfCell().getStringCellValue()!=null){
            cell.setHssfCellValue( cell.getHssfCell().getStringCellValue() );
            if( cell.getHssfCellValue().startsWith(configuration.getStartFormulaToken()) && cell.getHssfCellValue().lastIndexOf(configuration.getEndFormulaToken()) > 0 ){
                parseFormula();
            }
        }
        return cell.getFormula();
    }

    private void parseFormula() {
        // process formula cell
        int i = cell.getHssfCellValue().lastIndexOf(configuration.getEndFormulaToken());
        String expr = cell.getHssfCellValue().substring(2, i);
        cell.setFormula(new Formula(expr));
        cell.getFormula().setRowNum(new Integer(cell.getRow().getHssfRow().getRowNum()));
        cell.getFormula().setCellNum(new Integer(cell.getHssfCell().getCellNum()));
        if (i + 1 < cell.getHssfCellValue().length()) {
            String tail = cell.getHssfCellValue().substring(i+1);
            int j = tail.indexOf(configuration.getMetaInfoToken());
            if( j >= 0 ){
                cell.setMetaInfo(tail.substring(j));
                if( j > 0 ){
                    cell.setLabel(tail.substring(0, j));
                }
                cell.setCollectionName(tail.substring(j + 2));
            }else{
                cell.setLabel(tail);
            }
        }
        cell.setStringCellValue(cell.getHssfCellValue().substring(0, i+1));
    }

    private void parseCellExpression(Map beans) {
        cell.setCollectionProperty(null);
        String curValue = cell.getHssfCellValue();
        int depRowNum = 0;
        int j = curValue.lastIndexOf(configuration.getMetaInfoToken());
        if( j>=0 ){
            cell.setStringCellValue(cell.getHssfCellValue().substring(0, j));
            cell.setMetaInfo(cell.getHssfCellValue().substring(j + 2));
            String tail = curValue.substring(j + 2);
            // processing additional parameters
                // check if there is collection property name specified
                int k = tail.indexOf(":");
                if( k >= 0 ){
                    try {
                        depRowNum = Integer.parseInt( tail.substring(k+1) );
                    } catch (NumberFormatException e) {
                        // ignore it if not an integer
                    }
                    cell.setCollectionName(tail.substring(0, k));
                }else{
                    cell.setCollectionName(tail);
                }
                curValue = curValue.substring(0, j);
        }else{
            cell.setStringCellValue(cell.getHssfCellValue());
        }

        try {
            while( curValue.length()>0 ){
                int i = curValue.indexOf(configuration.getStartExpressionToken());
                if( i>=0 ) {
                    int k = curValue.indexOf(configuration.getEndExpressionToken(), i+2);
                    if( k>=0 ){
                        // new bean property found
                        String expr = curValue.substring(i+2, k);
                        if( i>0 ){
                            String before = curValue.substring(0, i);
                            cell.getExpressions().add( new Expression( before, configuration ) );
                        }
                        Expression expression = new Expression(expr, beans, configuration);
                        if( expression.getCollectionProperty() != null ){
                            if( cell.getCollectionProperty() == null ){
                                cell.setCollectionName(expression.getCollectionProperty().getFullCollectionName());
                                cell.setCollectionProperty(expression.getCollectionProperty());
                                cell.setDependentRowNumber(depRowNum);
                            }else{
                                if( log.isInfoEnabled() ){
                                    log.info("Only the same collection property in a cell is allowed.");
                                }
                            }
                        }
                        cell.getExpressions().add( expression );
                        curValue = curValue.substring(k+1, curValue.length());
                    }else{
                        cell.getExpressions().add( new Expression(curValue, configuration) );
                    }
                }else{
                    if( curValue.length()!=cell.getHssfCellValue().length() ){
                        cell.getExpressions().add( new Expression( curValue, configuration ));
                    }
                    curValue = "";
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            log.error("Can't parse expression", e);
        }
    }

    private void parseCellValue(Map beans) throws ParsePropertyException {
        if( cell.getHssfCellValue() !=null ){
            if( cell.getHssfCellValue().startsWith(configuration.getStartFormulaToken()) && cell.getHssfCellValue().lastIndexOf(configuration.getEndFormulaToken()) > 0 ){
                parseFormula();
            }else if(cell.getHssfCellValue().startsWith( "<" + configuration.getTagPrefix() )){
//                String tagName = cell.getHssfCellValue().split("(?<=<" + configuration.getTagPrefix() + ")\\w+", 2)[0];
                String tagName = getTagName( cell.getHssfCellValue() );
                if( tagName!=null ){
                    HSSFCell hssfCell = findMatchingPairInRow( cell.getRow().getHssfRow(), tagName );
                    if( hssfCell!=null ){
                        // closing tag is in the same row
                        Block tagBody = new Block(cell.getRow().getHssfRow().getRowNum(), cell.getHssfCell().getCellNum(),
                                cell.getRow().getHssfRow().getRowNum(), hssfCell.getCellNum());
                        parseTag( tagName, tagBody, beans);
                    }else{
                        HSSFRow hssfRow = findMatchingPair( tagName );
                        if( hssfRow!=null ){
                            // closing tag is in hssfRow
                            int lastTagBodyRowNum = hssfRow.getRowNum() ;
                            Block tagBody = new Block(null, cell.getRow().getHssfRow().getRowNum(), lastTagBodyRowNum);
                            parseTag( tagName, tagBody, beans );
                        }else{
                            log.warn("Can't find matching tag pair for " + cell.getHssfCellValue());
                        }
                    }
                }
            }else{
                parseCellExpression(beans);
            }
        }
    }

    private HSSFCell findMatchingPairInRow(HSSFRow hssfRow, String tagName) {
        int count = 0;
        if( hssfRow!=null ){
            for(short j = (short) (cell.getHssfCell().getCellNum() + 1); j <= hssfRow.getLastCellNum(); j++){
                HSSFCell hssfCell = hssfRow.getCell( j );
                if( hssfCell != null && hssfCell.getCellType() == HSSFCell.CELL_TYPE_STRING ){
                    String cellValue = hssfCell.getStringCellValue();
                    if( cellValue.matches("<" + configuration.getTagPrefix() + tagName + "\\b.*")){
                        count++;
                    }else{
                        if( cellValue.matches("</" + configuration.getTagPrefix() + tagName + ">" )){
                            if( count == 0 ){
                                return hssfCell;
                            }else{
                                count--;
                            }
                        }
                    }
                }
            }
        }
        return null;
    }

    private String getTagName(String xmlTag){
        int i = configuration.getTagPrefix().length() + 1;
        int j = i;
        while( j < xmlTag.length() && Character.isLetterOrDigit( xmlTag.charAt( j ) ) ){
            j++;
        }
        if( j == xmlTag.length() ){
            log.warn("can't determine tag name");
            return null;
        }else{
            return xmlTag.substring(i, j);
        }
    }

    private HSSFRow findMatchingPair(String tagName) {
        HSSFSheet hssfSheet = cell.getRow().getSheet().getHssfSheet();
        int count = 0;

        for( int i = cell.getRow().getHssfRow().getRowNum() + 1; i <= hssfSheet.getLastRowNum(); i++ ){
            HSSFRow hssfRow = hssfSheet.getRow( i );
            if( hssfRow!=null ){
                for(short j = hssfRow.getFirstCellNum(); j <= hssfRow.getLastCellNum(); j++){
                    HSSFCell hssfCell = hssfRow.getCell( j );
                    if( hssfCell != null && hssfCell.getCellType() == HSSFCell.CELL_TYPE_STRING ){
                        String cellValue = hssfCell.getStringCellValue();
                        if( cellValue.matches("<" + configuration.getTagPrefix() + tagName + "\\b.*")){
                            count++;
                        }else{
                            if( cellValue.matches("</" + configuration.getTagPrefix() + tagName + ">" )){
                                if( count == 0 ){
                                    return hssfRow;
                                }else{
                                    count--;
                                }
                            }
                        }
                    }
                }
            }

        }

        return null;
    }

    private void parseTag(String tagName, Block tagBody, Map beans){
        Digester digester = new Digester();
        digester.setNamespaceAware(true);
        digester.setRuleNamespaceURI( Configuration.NAMESPACE_URI );

        digester.setValidating( false );
        Set tagKeys = Taglib.getTagMap().keySet();
        for (Iterator iterator = tagKeys.iterator(); iterator.hasNext();) {
            String tagKey = (String) iterator.next();
            digester.addObjectCreate( Configuration.JXLS_ROOT_TAG + "/" + tagKey, (String) Taglib.getTagMap().get( tagKey ) );
            digester.addSetProperties( Configuration.JXLS_ROOT_TAG + "/" + tagKey );
        }
        try {
            String xml = Configuration.JXLS_ROOT_START + cell.getHssfCellValue() + "</" +
                    configuration.getTagPrefix() + tagName + ">" + Configuration.JXLS_ROOT_END;
            String escapedXml = Util.escapeAttributes( xml );
            Tag tag = (Tag) digester.parse(new StringReader( escapedXml ) );
            cell.setTag( tag );
            TagContext tagContext = new TagContext( cell.getRow().getSheet(), tagBody, beans );
            tag.init( tagContext );
        } catch (IOException e) {
            log.warn( "Can't parse cell tag " + cell.getHssfCellValue(), e);
        } catch (SAXException e) {
            log.warn( "Can't parse cell tag " + cell.getHssfCellValue(), e);
        }
    }

    private void updateMergedRegions() {
        cell.setMergedRegion(Util.getMergedRegion( cell.getRow().getSheet().getHssfSheet(), cell.getRow().getHssfRow().getRowNum(), cell.getHssfCell().getCellNum() ));
    }
}
