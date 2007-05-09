package net.sf.jxls.transformer;

import net.sf.jxls.controller.WorkbookTransformationController;
import net.sf.jxls.controller.WorkbookTransformationControllerImpl;
import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.formula.CommonFormulaResolver;
import net.sf.jxls.formula.FormulaController;
import net.sf.jxls.formula.FormulaResolver;
import net.sf.jxls.processor.CellProcessor;
import net.sf.jxls.processor.PropertyPreprocessor;
import net.sf.jxls.processor.RowProcessor;
import net.sf.jxls.util.Util;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.util.*;

/**
 * <p> This class uses excel template to generate excel file filled with required objects and collections.
 * <p/>
 * @author Leonid Vysochyn
 * @author Vincent Dutat
 */
public class XLSTransformer {
    protected final Log log = LogFactory.getLog(getClass());
    /**
     * property preprocessors will be applied before main transformation starts
     */
    private List propertyPreprocessors = new ArrayList();

    private List rowProcessors = new ArrayList();

    private List cellProcessors = new ArrayList();

//    private Map taglibs = new HashMap();

    private final String TAGLIB_DEFINITION_FILE = "jxls-core-taglib.xml";


    /**
     * Registers Property Preprocessor that will be applied before main template transformation
     * it is possible to have many Property Preprocessors
     * @param propPreprocessor - {@link PropertyPreprocessor} interface implementation
     */
    public void registerPropertyPreprocessor(PropertyPreprocessor propPreprocessor) {
        if (propPreprocessor != null) {
            propertyPreprocessors.add(propPreprocessor);
        }
    }

    /**
     * Registers {@link net.sf.jxls.processor.RowProcessor} object
     * @param rowProcessor {@link net.sf.jxls.processor.RowProcessor} to register
     */
    public void registerRowProcessor(RowProcessor rowProcessor) {
        if (rowProcessor != null) {
            rowProcessors.add(rowProcessor);
        }
    }

    /**
     * Registers {@link net.sf.jxls.processor.CellProcessor} object
     * @param cellProcessor {@link net.sf.jxls.processor.CellProcessor to register}
     */
    public void registerCellProcessor(CellProcessor cellProcessor) {
        if (cellProcessor != null) {
            cellProcessors.add(cellProcessor);
        }
    }

    /**
     * Mark a collection as static collection.
     * All static collection rows are presented in Excel template and will not be expanded
     * @param collectionName - Collection name to mark as fixed size collection
     */
    public void markAsFixedSizeCollection(String collectionName) {
        fixedSizeCollections.add(collectionName);
    }

//    public void registerTaglib(String prefix, Taglib taglib){
//        if( taglibs.containsKey( prefix ) ){
//            throw new TaglibRegistrationException( "Tag library with prefix '" + prefix + "' already registered");
//        }else{
//            taglibs.put( prefix, taglib );
//        }
//    }

    /**
     * Column numbers to hide
     */
    private short[] columnsToHide;

    private Set spreadsheetsToHide = new HashSet();

    private Map spreadsheetsToRename = new HashMap();    // hash map 'spdsheet tpl name' => 'new name'

    private String[] columnPropertyNamesToHide;

    Map customTags = new HashMap();

    /**
     * Stores the names of all 'fixed size' collections.
     * 'Fixed size' collection is a collection with fixed number of items which do not require to create new rows in excel file
     * because all rows for them are already presented in template file.
     */
    private Set fixedSizeCollections = new HashSet();

    /**
     * {@link Set} of all collections to outline
     */
    private Set groupedCollections = new HashSet();

    /**
     * {@link net.sf.jxls.transformer.Configuration} class
     */
    private Configuration configuration;

    public Configuration getConfiguration() {
        return configuration;
    }

    public void setConfiguration(Configuration configuration) {
        this.configuration = configuration;
    }

    public XLSTransformer() {
        this(new Configuration());
    }

    public XLSTransformer(Configuration configuration) {
        if( configuration!=null ){
            this.configuration = configuration;
        }else{
            this.configuration = new Configuration();
        }
        //todo
//        registerTaglib( TAGLIB_DEFINITION_FILE );
    }

//    public void registerTaglib(String taglibFileName){
//        TaglibXMLParser parser = new TaglibXMLParser();
//        Taglib taglib = parser.parseTaglibXMLFile( taglibFileName );
//    }

    private WorkbookTransformationController workbookTransformationController;

    private FormulaResolver formulaResolver;

    /**
     * @return {@link net.sf.jxls.formula.FormulaResolver} used to resolve coded formulas
     */
    public FormulaResolver getFormulaResolver() {
        return formulaResolver;
    }

    /**
     * Sets {@link FormulaResolver} to be used in resolving formula
     * @param formulaResolver {@link FormulaResolver} implementation to set
     */
    public void setFormulaResolver(FormulaResolver formulaResolver) {
        this.formulaResolver = formulaResolver;
    }

    public boolean isJexlInnerCollectionsAccess() {
        return configuration.isJexlInnerCollectionsAccess();
    }

    public void setJexlInnerCollectionsAccess(boolean jexlInnerCollectionsAccess) {
        configuration.setJexlInnerCollectionsAccess( jexlInnerCollectionsAccess );
    }


    /**
     * Set this collection to be grouped (outlined).
     * @param collectionName - Collection name to use for grouping
     */
    public void groupCollection(String collectionName) {
        groupedCollections.add(collectionName);
    }


    /**
     * Creates new .xls file at a given path using specified excel template file and a number of beans.
     * This method invokes {@link #transformXLS(InputStream is, Map beanParams)} passing input stream from a template file
     * and then writing resulted HSSFWorkbook into required output file.
     *
     * @param srcFilePath  Path to source .xls template file
     * @param beanParams   Map of beans to be applied to .xls template file with keys corresponding to bean aliases in template
     * @param destFilePath Path to result .xls file
     * @throws ParsePropertyException if there were any problems in evaluating specified property value from a bean
     * @throws IOException            if there were any access or input/output problems with source or destination file
     */
    public void transformXLS(String srcFilePath, Map beanParams, String destFilePath) throws ParsePropertyException, IOException {
        InputStream is = new BufferedInputStream(new FileInputStream(srcFilePath));
        HSSFWorkbook workbook = transformXLS(is, beanParams);
        OutputStream os = new BufferedOutputStream(new FileOutputStream(destFilePath));
        workbook.write(os);
        is.close();
        os.flush();
        os.close();
    }

    /**
     * Creates HSSFWorkbook instance based on .xls template from a given InputStream and a number of beans
     *
     * @param is         xls InputStream with required
     * @param beanParams Beans in a map under keys used in .xls template to access to the beans
     * @return new {@link HSSFWorkbook} generated by inserting beans into corresponding excel template
     * @throws net.sf.jxls.exception.ParsePropertyException if there were any problems in evaluating specified property value from a bean
     */
    public HSSFWorkbook transformXLS(InputStream is, Map beanParams) throws ParsePropertyException {
        HSSFWorkbook hssfWorkbook = null;
        try {
            POIFSFileSystem fs = new POIFSFileSystem(is);
            hssfWorkbook = new HSSFWorkbook(fs);
            Workbook workbook = createWorkbook( hssfWorkbook );
            exposePOIObjects(workbook, beanParams);
            workbookTransformationController = new WorkbookTransformationControllerImpl( workbook );
            preprocess(hssfWorkbook);
            SheetTransformer sheetTransformer = new SheetTransformer( fixedSizeCollections, groupedCollections, rowProcessors, cellProcessors, configuration) ;
            for (int sheetNo = 0; sheetNo < hssfWorkbook.getNumberOfSheets(); sheetNo++) {
                final String spreadsheetName = hssfWorkbook.getSheetName(sheetNo);
                if( spreadsheetName != null && !spreadsheetName.startsWith( configuration.getExcludeSheetProcessingMark() )){
                    if (!isSpreadsheetToHide(spreadsheetName)) {
                        if (isSpreadsheetToRename(spreadsheetName)) {
                            hssfWorkbook.setSheetName(sheetNo, getSpreadsheetToReName(spreadsheetName));
                        }
                        Sheet sheet = workbook.getSheetAt( sheetNo );
                        sheetTransformer.transformSheet( workbookTransformationController, sheet, beanParams );
                    } else {
                        // let's remove spreadsheet
                        workbook.removeSheetAt( sheetNo );
                        sheetNo--;
                    }
                }
            }
            updateFormulas();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return hssfWorkbook;
    }

    private void exposePOIObjects(Workbook workbook, Map beanParams) {
        beanParams.put( configuration.getWorkbookKeyName(), workbook.getHssfWorkbook() );
    }

    /**
     * This method transforms given XLS input stream template into multiple sheets workbook
     * creating separate Excel worksheet for every object in the list
     * @param is        - {@link InputStream} for source XLS template file
     * @param objects   - List of beans where each list item should be exported into a separated worksheet
     * @param newSheetNames - Sheet names to be used for newly created worksheets
     * @param beanName - Bean name to be used for a list item when processing sheet
     * @param beanParams - Common bean map containing all other objects to be used in the workbook
     * @param startSheetNum - Worksheet number (zero based) of the worksheet that needs to be used to create multiple worksheets
     * @return new {@link HSSFWorkbook} object containing the result of transformation
     * @throws net.sf.jxls.exception.ParsePropertyException - {@link ParsePropertyException} is thrown when some property can't be parsed
     */
    public HSSFWorkbook transformMultipleSheetsList(InputStream is, List objects, List newSheetNames, String beanName, Map beanParams, int startSheetNum) throws ParsePropertyException {
        HSSFWorkbook hssfWorkbook = null;
        try {
            if( beanParams!=null && beanParams.containsKey( beanName )){
                throw new IllegalArgumentException("Selected bean name '" + beanName + "' already exists in the bean map");
            }
            if( beanName==null ){
                throw new IllegalArgumentException(("Bean name must not be null" ) );
            }
            if( beanParams == null ){
                beanParams = new HashMap();
            }
            POIFSFileSystem fs = new POIFSFileSystem(is);
            hssfWorkbook = new HSSFWorkbook(fs);

            preprocess(hssfWorkbook);

            Workbook workbook = createWorkbook( hssfWorkbook );
            exposePOIObjects( workbook,  beanParams );
            workbookTransformationController = new WorkbookTransformationControllerImpl( workbook );

            SheetTransformer sheetTransformer = new SheetTransformer( fixedSizeCollections, groupedCollections, rowProcessors, cellProcessors, configuration) ;
            final String templateSheetName = "InternalTemplateSheetName";
            // todo refactoring required
            for (int sheetNo = 0; sheetNo < hssfWorkbook.getNumberOfSheets(); sheetNo++) {
                final String spreadsheetName = hssfWorkbook.getSheetName(sheetNo);
                if (!isSpreadsheetToHide(spreadsheetName)) {
                    if (isSpreadsheetToRename(spreadsheetName)) {
                        hssfWorkbook.setSheetName(sheetNo, getSpreadsheetToReName(spreadsheetName));
                    }
                    HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(sheetNo);
                    if( startSheetNum == sheetNo && objects != null && !objects.isEmpty()){
                        Object firstBean = objects.get(0);
                        beanParams.put( beanName, firstBean );
                        hssfWorkbook.setSheetName( sheetNo, (String) newSheetNames.get(0), HSSFWorkbook.ENCODING_UTF_16);
                        HSSFSheet templateSheet = hssfWorkbook.createSheet(templateSheetName );
                        Util.copySheets( templateSheet, hssfSheet );
                        Sheet sheet = workbook.getSheetAt( sheetNo );
                        sheetTransformer.transformSheet(workbookTransformationController, sheet, beanParams );
                        for (int i = 1; i < objects.size(); i++) {
                            Object bean = objects.get(i);
                            beanParams.put( beanName, bean );
                            HSSFSheet newSheet = hssfWorkbook.createSheet( (String) newSheetNames.get(i) );
                            Util.copySheets(newSheet, templateSheet);
                            Util.copyPageSetup(newSheet, hssfSheet);
                            Util.copyPrintSetup(newSheet, hssfSheet);
                            sheet = new Sheet(hssfWorkbook, newSheet, configuration);
                            // todo: implement update of the FormulaController instance when adding new sheet to workbook
                            workbook.addSheet( sheet );
                            workbook.initSheetNames();
                            sheetTransformer.transformSheet(workbookTransformationController, sheet, beanParams );
                        }
                        hssfWorkbook.removeSheetAt( hssfWorkbook.getSheetIndex( templateSheetName ) );
                        beanParams.remove( beanName );
                    }else{
                          Sheet sheet = workbook.getSheetAt( sheetNo );
                        sheetTransformer.transformSheet(workbookTransformationController, sheet, beanParams );
                    }
                } else {
                    // let's remove spreadsheet
                    workbook.removeSheetAt( sheetNo );
                    sheetNo--;
                }
            }
            updateFormulas();
        } catch (IOException e) {
            e.printStackTrace();
        }
        if( hssfWorkbook != null ){
            for(int i = 0;i < hssfWorkbook.getNumberOfSheets();i++)
            {
                Util.setPrintArea(hssfWorkbook,i);
            }
        }
        return hssfWorkbook;
    }

    /**
     * Multiple sheet template multiple transform.
     * It can be used to generate a workbook with N (N=N1+N2+...+Nn) sheets based on :
     * - N1 transformations of the sheet template T1
     * - N2 transformations of the sheet template T2
     * ...
     * - Nn transformations of the sheet template Tn
     * @param is  the {@link InputStream} of the workbook template containing the n template sheets
     * @param templateSheetNameList  the ordered list of the template sheet name used in the transformation.
     * @param sheetNameList  the ordered list of the resulting sheet name after transformation
     * @param beanParamsList  the ordered list of beanParams used in the transformation
     * @return - {@link HSSFWorkbook} representing transformation result
     * @throws ParsePropertyException  in case property parsing failure
     */
    public HSSFWorkbook transformXLS(InputStream is, List
            templateSheetNameList, List sheetNameList, List beanParamsList)
            throws ParsePropertyException {
        HSSFWorkbook hssfWorkbook = null;
        try {
            POIFSFileSystem fs = new POIFSFileSystem(is);
            hssfWorkbook = new HSSFWorkbook(fs);
            int numberOfSheets = hssfWorkbook.getNumberOfSheets();
            for (int templateSheetIndex = 0; templateSheetIndex < templateSheetNameList.size(); templateSheetIndex++) {
                String templateSheetName = (String)templateSheetNameList.get(templateSheetIndex);
                String sheetName = (String)sheetNameList.get(templateSheetIndex);
                for(int workbookSheetIndex = 0; workbookSheetIndex < numberOfSheets; workbookSheetIndex++) {
                    if (templateSheetName.equals(hssfWorkbook.getSheetName(workbookSheetIndex))) {
                        cloneSheet(hssfWorkbook, workbookSheetIndex, sheetName);
                        break;
                    }
                }
            }
            for (int i = 0; i < numberOfSheets; i++) {
                hssfWorkbook.removeSheetAt(0);
            }
            Workbook workbook = createWorkbook(hssfWorkbook);
            workbookTransformationController = new WorkbookTransformationControllerImpl(workbook);
            preprocess(hssfWorkbook);
            SheetTransformer sheetTransformer = new SheetTransformer(fixedSizeCollections, groupedCollections,
                    rowProcessors, cellProcessors, configuration);
            for (int sheetNo = 0; sheetNo < workbook.getNumberOfSheets(); sheetNo++) {
                final String spreadsheetName = hssfWorkbook.getSheetName(sheetNo);
                if (!isSpreadsheetToHide(spreadsheetName)) {
                    if (isSpreadsheetToRename(spreadsheetName)) {
                        hssfWorkbook.setSheetName(sheetNo, getSpreadsheetToReName(spreadsheetName));
                    }
                    Sheet sheet = workbook.getSheetAt(sheetNo);

                    Map beanParams = (Map) beanParamsList.get(sheetNo);
                    beanParams.put("index", String.valueOf(sheetNo));
                    exposePOIObjects( workbook,  beanParams );
                    sheetTransformer.transformSheet(workbookTransformationController, sheet, beanParams);
                } else {
                    // let's remove spreadsheet
                    workbook.removeSheetAt( sheetNo );
                    sheetNo--;
                }
            }
            updateFormulas();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return hssfWorkbook;
    }

    private void cloneSheet(HSSFWorkbook hssfWorkbook, int index, String name) {
        HSSFSheet hssfSheet = hssfWorkbook.cloneSheet(index);
        for (int i = 0; i < hssfWorkbook.getNumberOfSheets(); i++) {
            if(hssfSheet.equals(hssfWorkbook.getSheetAt(i))) {
                hssfWorkbook.setSheetName(i, name);
                break;
            }
        }
    }

    private Workbook createWorkbook(HSSFWorkbook hssfWorkbook) {
        Workbook workbook = new Workbook(hssfWorkbook);
        for (int sheetNo = 0; sheetNo < hssfWorkbook.getNumberOfSheets(); sheetNo++) {
            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(sheetNo);
            workbook.addSheet( new Sheet(hssfWorkbook, hssfSheet, configuration));
        }
        workbook.initSheetNames();
        workbook.createFormulaController();
        return workbook;
    }

    private void updateFormulas() {
        if( formulaResolver == null ){
            formulaResolver = new CommonFormulaResolver();
        }
        FormulaController formulaController = workbookTransformationController.getWorkbook().getFormulaController();
        formulaController.writeFormulas( formulaResolver );
    }


    private void preprocess(HSSFWorkbook workbook) {
        hideColumns(workbook);
        hideColumnsByPropertyName(workbook);
        for (int sheet_no = 0; sheet_no < workbook.getNumberOfSheets(); sheet_no++) {
            HSSFSheet sheet = workbook.getSheetAt(sheet_no);
            for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
                HSSFRow hssfRow = sheet.getRow(i);
                if (hssfRow != null) {
                    for (short j = hssfRow.getFirstCellNum(); j <= hssfRow.getLastCellNum(); j++) {
                        HSSFCell cell = hssfRow.getCell(j);
                        if (cell != null && cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                            String value = cell.getStringCellValue();
                            for (int k = 0; k < propertyPreprocessors.size(); k++) {
                                PropertyPreprocessor propertyPreprocessor = (PropertyPreprocessor) propertyPreprocessors.get(k);
                                String newValue = propertyPreprocessor.processProperty(value);
                                if (newValue != null) {
                                    cell.setCellValue(newValue);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    private void hideColumns(HSSFWorkbook workbook) {
        if (columnsToHide != null) {
            for (int i = 0; i < columnsToHide.length; i++) {
                short column = columnsToHide[i];
                for (int sheet_no = 0; sheet_no < workbook.getNumberOfSheets(); sheet_no++) {
                    HSSFSheet sheet = workbook.getSheetAt(sheet_no);
                    sheet.setColumnWidth(column, (short) 0);
                }
            }
        }
    }

    /**
     * Set column width = 0 for column if any it cell value contains any of {@link this#columnPropertyNamesToHide} string.
     * @param workbook - {@link HSSFWorkbook} to hide columns in
     */
    private void hideColumnsByPropertyName(HSSFWorkbook workbook) {
        if (columnPropertyNamesToHide == null) return;
        for (int sheet_no = 0; sheet_no < workbook.getNumberOfSheets(); sheet_no++) {
            HSSFSheet sheet = workbook.getSheetAt(sheet_no);
            //for all rows
            for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
                HSSFRow hssfRow = sheet.getRow(i);
                if (hssfRow != null) {
                    //for all cells
                    for (short j = hssfRow.getFirstCellNum(); j <= hssfRow.getLastCellNum(); j++) {
                        HSSFCell cell = hssfRow.getCell(j);
                        if (cell != null && cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                            String value = cell.getStringCellValue();
                            //if any from columnPropertyNamesToHide is substring of cell value, than hide column
                            for (int prptIndx = 0; prptIndx < columnPropertyNamesToHide.length; prptIndx++) {
                                if (value != null && value.indexOf(columnPropertyNamesToHide[prptIndx]) != -1) {
                                    sheet.setColumnWidth(j, (short) 0);
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * @return The column numbers to hide in result XLS
     */
    public short[] getColumnsToHide() {
        return columnsToHide;
    }

    /**
     * Sets the columns to hide in result XLS
     * @param columnsToHide - Column numbers to hide
     */
    public void setColumnsToHide(short[] columnsToHide) {
        this.columnsToHide = columnsToHide;
    }

    /**
     * @return The property names for which all columns containing them should be hidden
     */
    public String[] getColumnPropertyNamesToHide() {
        return columnPropertyNamesToHide;
    }

    /**
     * Set the columns to hide in result XLS
     * @param columnPropertyNamesToHide - The names of bean properties for which all columns
     * containing this properties should be hidden
     */
    public void setColumnPropertyNamesToHide(String[] columnPropertyNamesToHide) {
        this.columnPropertyNamesToHide = columnPropertyNamesToHide;
    }

    /**
     * Set spreadsheets with given names to be hidden
     * @param names - Names of the worksheets to hide
     */
    public void setSpreadsheetsToHide(String[] names) {
        spreadsheetsToHide.clear();
        for (int i = 0; i < names.length; i++) {
            spreadsheetsToHide.add(names[i]);
        }
    }

    public void setSpreadsheetToRename(String name, String newName) {
        spreadsheetsToRename.put(name, newName);
    }

    protected boolean isSpreadsheetToHide(String name) {
        return spreadsheetsToHide.contains(name);
    }

    protected boolean isSpreadsheetToRename(String name) {
        return spreadsheetsToRename.containsKey(name);
    }

    protected String getSpreadsheetToReName(String name) {
        final String newName = (String) spreadsheetsToRename.get(name);
        if (newName != null)
            return newName;
        return name;
    }

}
