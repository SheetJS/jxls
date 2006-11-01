package net.sf.jxls.transformer;

/**
 * Defines configuration properties for XLS transformation
 * @author Leonid Vysochyn
 */
public class Configuration {
    private String metaInfoToken = "//";
    private String startExpressionToken = "${";
    private String endExpressionToken = "}";
    private String startFormulaToken = "$[";
    private String endFormulaToken = "]";

    String tagPrefix = "jx";
    String forTagName = "forEach";
    String forTagItems = "items";
    String forTagVar = "var";
    boolean isUTF16 = false;

    private String excludeSheetProcessingMark = "#Exclude";

    public Configuration() {
    }

    public Configuration(String startExpressionToken, String endExpressionToken, String startFormulaToken, String endFormulaToken, String metaInfoToken) {
        this.startExpressionToken = startExpressionToken;
        this.endExpressionToken = endExpressionToken;
        this.startFormulaToken = startFormulaToken;
        this.endFormulaToken = endFormulaToken;
        this.metaInfoToken = metaInfoToken;
    }

    public Configuration(String startExpressionToken, String endExpressionToken, String startFormulaToken, String endFormulaToken, String metaInfoToken, boolean isUTF16) {
        this.startExpressionToken = startExpressionToken;
        this.endExpressionToken = endExpressionToken;
        this.startFormulaToken = startFormulaToken;
        this.endFormulaToken = endFormulaToken;
        this.metaInfoToken = metaInfoToken;
        this.isUTF16 = isUTF16;
    }

    public static final String NAMESPACE_URI = "http://jxls.sourceforge.net/jxls";
    public static final String JXLS_ROOT_TAG = "jxls";
    public static final String JXLS_ROOT_START = "<jx:jxls xmlns:jx=\"" + NAMESPACE_URI + "\">";
    public static final String JXLS_ROOT_END = "</jx:jxls>";


    private boolean jexlInnerCollectionsAccess;

    public boolean isJexlInnerCollectionsAccess() {
        return jexlInnerCollectionsAccess;
    }

    public void setJexlInnerCollectionsAccess(boolean jexlInnerCollectionsAccess) {
        this.jexlInnerCollectionsAccess = jexlInnerCollectionsAccess;
    }
    
    public boolean isUTF16() {
        return isUTF16;
    }

    public void setUTF16(boolean UTF16) {
        this.isUTF16 = UTF16;
    }

    public String getTagPrefix() {
        if( tagPrefix.length()>0 ){
            return tagPrefix + ":";
        }else{
            return tagPrefix;
        }
    }

    public String getForTagName() {
        return forTagName;
    }

    public String getForTagItems() {
        return forTagItems;
    }

    public String getForTagVar() {
        return forTagVar;
    }

    public String getMetaInfoToken() {
        return metaInfoToken;
    }

    public void setMetaInfoToken(String metaInfoToken) {
        this.metaInfoToken = metaInfoToken;
    }

    public String getStartExpressionToken() {
        return startExpressionToken;
    }

    public void setStartExpressionToken(String startExpressionToken) {
        this.startExpressionToken = startExpressionToken;
    }

    public String getEndExpressionToken() {
        return endExpressionToken;
    }

    public void setEndExpressionToken(String endExpressionToken) {
        this.endExpressionToken = endExpressionToken;
    }

    public String getStartFormulaToken() {
        return startFormulaToken;
    }

    public void setStartFormulaToken(String startFormulaToken) {
        this.startFormulaToken = startFormulaToken;
    }

    public String getEndFormulaToken() {
        return endFormulaToken;
    }

    public void setEndFormulaToken(String endFormulaToken) {
        this.endFormulaToken = endFormulaToken;
    }

    public String getExcludeSheetProcessingMark() {
        return excludeSheetProcessingMark;
    }

    public void setExcludeSheetProcessingMark(String excludeSheetProcessingMark) {
        this.excludeSheetProcessingMark = excludeSheetProcessingMark;
    }
}
