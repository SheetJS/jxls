package net.sf.jxls.reader;

import org.apache.commons.beanutils.ConversionException;
import org.apache.commons.beanutils.Converter;
import org.apache.poi.ss.usermodel.DateUtil;

/**
 * @author Leonid Vysochyn
 * @version 1.0 Jul 29, 2007
 */
public class DateConverter implements Converter {
    public Object convert(Class type, Object value) {
        if( value == null ) {
            throw new ConversionException("No value specified");
        }
        double date;
        if( value instanceof Double ){
            date = ((Double)value).doubleValue();
        }else if( value instanceof Number){
            date = ((Number)value).doubleValue();
        }else if( value instanceof String ){
            try {
                date = Double.parseDouble( (String)value );
            } catch (NumberFormatException e) {
                throw new ConversionException(e);
            }
        }else if(value instanceof java.util.Date) {
            return value;
        }else{
            throw new ConversionException("No value specified");
        }
        return DateUtil.getJavaDate( date );
    }
}
