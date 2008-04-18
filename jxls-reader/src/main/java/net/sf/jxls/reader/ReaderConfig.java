/**
 * @version 1.0 28.07.2007
 * @author Leonid Vysochyn
 */
package net.sf.jxls.reader;

import org.apache.commons.beanutils.ConvertUtils;
import org.apache.commons.beanutils.Converter;
import org.apache.commons.beanutils.converters.BooleanConverter;
import org.apache.commons.beanutils.converters.ByteConverter;
import org.apache.commons.beanutils.converters.CharacterConverter;
import org.apache.commons.beanutils.converters.DoubleConverter;
import org.apache.commons.beanutils.converters.FloatConverter;
import org.apache.commons.beanutils.converters.IntegerConverter;
import org.apache.commons.beanutils.converters.LongConverter;
import org.apache.commons.beanutils.converters.ShortConverter;

public class ReaderConfig {
    private static ReaderConfig ourInstance = new ReaderConfig();

    private boolean skipErrors = false;
    private boolean useDefaultValuesForPrimitiveTypes = true;

    /**
     * The default value for Character conversions.
     */
    private Character defaultCharacter = new Character(' ');
    /**
     * The default value for Byte conversions.
     */
    private Byte defaultByte = new Byte((byte) 0);
    /**
     * The default value for Boolean conversions.
     */
    private Boolean defaultBoolean = Boolean.FALSE;
    /**
     * The default value for Double conversions.
     */
    private Double defaultDouble = new Double((double) 0.0);
    /**
     * The default value for Float conversions.
     */
    private Float defaultFloat = new Float((float) 0.0);
    /**
     * The default value for Integer conversions.
     */
    private Integer defaultInteger = new Integer(0);
    /**
     * The default value for Long conversions.
     */
    private Long defaultLong = new Long((long) 0);
    /**
     * The default value for Short conversions.
     */
    private static Short defaultShort = new Short((short) 0);



    public static ReaderConfig getInstance() {
        return ourInstance;
    }

    private ReaderConfig() {
        setUseDefaultValuesForPrimitiveTypes( false );
        ConvertUtils.register( new DateConverter(), java.util.Date.class);
    }

    public boolean isSkipErrors() {
        return skipErrors;
    }

    public void setSkipErrors(boolean skipErrors) {
        this.skipErrors = skipErrors;
    }


    public boolean isUseDefaultValuesForPrimitiveTypes() {
        return useDefaultValuesForPrimitiveTypes;
    }

    public void setUseDefaultValuesForPrimitiveTypes(boolean useDefaultValuesForPrimitiveTypes) {
        this.useDefaultValuesForPrimitiveTypes = useDefaultValuesForPrimitiveTypes;
        Converter integerConverter;
        Converter doubleConverter;
        Converter longConverter;
        Converter shortConverter;
        Converter booleanConverter;
        Converter floatConverter;
        Converter characterConverter;
        Converter byteConverter;
        if( useDefaultValuesForPrimitiveTypes ){
            integerConverter = new IntegerConverter( defaultInteger );
            byteConverter = new ByteConverter( defaultByte );
            doubleConverter = new DoubleConverter( defaultDouble);
            longConverter = new LongConverter( defaultLong );
            shortConverter = new ShortConverter( defaultShort );
            booleanConverter = new BooleanConverter( defaultBoolean );
            floatConverter = new FloatConverter( defaultFloat );
            characterConverter = new CharacterConverter( defaultCharacter );
        }else{
            integerConverter = new IntegerConverter();
            byteConverter = new ByteConverter(  );
            doubleConverter = new DoubleConverter();
            longConverter = new LongConverter();
            shortConverter = new ShortConverter();
            booleanConverter = new BooleanConverter();
            floatConverter = new FloatConverter();
            characterConverter = new CharacterConverter();
        }
        ConvertUtils.register( integerConverter, Integer.TYPE);
        ConvertUtils.register( integerConverter, Integer.class);
        ConvertUtils.register( byteConverter, Byte.TYPE);
        ConvertUtils.register( byteConverter, Byte.class);
        ConvertUtils.register( doubleConverter, Double.TYPE);
        ConvertUtils.register( doubleConverter, Double.class);
        ConvertUtils.register( longConverter, Long.TYPE);
        ConvertUtils.register( longConverter, Long.class);
        ConvertUtils.register( shortConverter, Short.TYPE);
        ConvertUtils.register( shortConverter, Short.class);
        ConvertUtils.register( booleanConverter, Boolean.TYPE);
        ConvertUtils.register( booleanConverter, Boolean.class);
        ConvertUtils.register( floatConverter, Float.TYPE);
        ConvertUtils.register( floatConverter, Float.class);
        ConvertUtils.register( characterConverter, Character.TYPE);
        ConvertUtils.register( characterConverter, Character.class);
    }
}
