package net.sf.jxls;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import junit.framework.TestCase;
import net.sf.jxls.bean.SimpleBean;
import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.transformer.XLSTransformer;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * @author 
 */
public class XLSFormatterBeanTest extends TestCase {
    protected final Log log = LogFactory.getLog(getClass());

    public static final String formatterBeanXLS = "/templates/formatterbean.xls";
    public static final String formatterBeanDestXLS = "target/formatterbean_output.xls";
    
    SimpleBean simpleBean1;
    SimpleBean simpleBean2;
    SimpleBean simpleBean3;
    List<SimpleBean> beanList = new ArrayList<SimpleBean>();
    List<Integer> ii = new ArrayList<Integer>();

    public XLSFormatterBeanTest() {
    }

    public XLSFormatterBeanTest(String s) {
        super(s);
    }

    protected void setUp() throws Exception {
        super.setUp();
        simpleBean1 = new SimpleBean("Bean 1", new Double(100.34567), new Integer(10), (Date) new Date());
        simpleBean2 = new SimpleBean("Bean 2", new Double(555.3), new Integer(123), null);
        simpleBean3 = new SimpleBean("Bean 3", new Double(777.569), new Integer(10234), new Date());

        beanList.add(simpleBean1);
        beanList.add(simpleBean2);
        beanList.add(simpleBean3);

        for (int i = 0; i < 256; ++i) {
        	ii.add( new Integer( i ) );
        }
    }

    public void testFormatting() throws IOException, ParsePropertyException, InvalidFormatException {
    	Map<String, Object> beans = new HashMap<String, Object>();
    	beans.put("beans", beanList);
    	beans.put("ii", ii);
    	beans.put("fmt", new Formatter());

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(formatterBeanXLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(formatterBeanXLS));
        Workbook sourceWorkbook = WorkbookFactory.create(is);

        Sheet sourceSheet = sourceWorkbook.getSheetAt(0);
        Sheet resultSheet = resultWorkbook.getSheetAt(0);

        //assert and formatting and styles
        
        is.close();
        saveWorkbook(resultWorkbook, formatterBeanDestXLS);
    }
    
    private void saveWorkbook(Workbook resultWorkbook, String fileName) throws IOException {
            if (log.isInfoEnabled()) {
                log.info("Saving " + fileName);
            }
            OutputStream os = new BufferedOutputStream(new FileOutputStream(fileName));
            resultWorkbook.write(os);
            os.flush();
            os.close();
            log.info("Output Excel saved to " + fileName);
    }
    
	public static class FontVO {
		private static String DELIM = "-";

		public String fontName;
		public short fontHeight;
		public short boldweight;
		public boolean italic;
		public boolean strikeout;
		public short typeOffset;
		public byte underline;
		public short color;

		public FontVO( HSSFFont font ) {
			fontName = font.getFontName();
			fontHeight = font.getFontHeight();
			boldweight = font.getBoldweight();
			italic = font.getItalic();
			strikeout = font.getStrikeout();
			typeOffset = font.getTypeOffset();
			underline = font.getUnderline();
			color = font.getColor();
		}

		@Override
		public String toString() {
			StringBuilder builder = new StringBuilder();
			builder.append( fontName );
			builder.append( DELIM );
			builder.append( fontHeight );
			builder.append( DELIM );
			builder.append( boldweight );
			builder.append( DELIM );
			builder.append( italic );

			builder.append( DELIM );
			builder.append( strikeout );
			builder.append( DELIM );
			builder.append( typeOffset );
			builder.append( DELIM );
			builder.append( underline );
			builder.append( DELIM );
			builder.append( color );

			return builder.toString();
		}

		public void applyTo( HSSFFont font ) {
			font.setFontName( fontName );
			font.setFontHeight( fontHeight );
			font.setBoldweight( boldweight );
			font.setItalic( italic );
			font.setStrikeout( strikeout );
			font.setTypeOffset( typeOffset );
			font.setUnderline( underline );
			font.setColor( color );
		}

	}

	public static class StyleVO {

		private static String DELIM = "-";

		public boolean setNext = true;

		public short alignment;
		public short borderBottom;
		public short borderLeft;
		public short borderRight;
		public short borderTop;
		public short bottomBorderColor;
		public short dataFormat;
		public short fillBackgroundColor;
		public short fillForegroundColor;
		public short fillPattern;
		public short fontIndex;
		public boolean hidden;
		public short indention;
		public short leftBorderColor;
		public boolean locked;
		public short rightBorderColor;
		public short rotation;
		public short topBorderColor;
		public short verticalAlignment;
		public boolean wrapText;

		public FontVO fontVal;

		public StyleVO() {
		}

		public StyleVO( HSSFCellStyle style, Workbook workbook ) {
			alignment = style.getAlignment();
			borderBottom = style.getBorderBottom();
			borderLeft = style.getBorderLeft();
			borderRight = style.getBorderRight();
			borderTop = style.getBorderTop();
			bottomBorderColor = style.getBottomBorderColor();
			dataFormat = style.getDataFormat();
			fillBackgroundColor = style.getFillBackgroundColor();
			fillForegroundColor = style.getFillForegroundColor();
			fillPattern = style.getFillPattern();
			fontIndex = style.getFontIndex();

			hidden = style.getHidden();
			indention = style.getIndention();
			leftBorderColor = style.getLeftBorderColor();
			locked = style.getLocked();
			rightBorderColor = style.getRightBorderColor();
			rotation = style.getRotation();
			topBorderColor = style.getTopBorderColor();
			verticalAlignment = style.getVerticalAlignment();
			wrapText = style.getWrapText();

			fontVal = new FontVO( style.getFont( workbook ) );
		}

		@Override
		public String toString() {
			StringBuilder builder = new StringBuilder();
			builder.append( alignment );
			builder.append( DELIM );
			builder.append( borderBottom );	
			builder.append( DELIM );
			builder.append( borderLeft );
			builder.append( DELIM );
			builder.append( borderRight );
			builder.append( DELIM );
			builder.append( borderTop );
			builder.append( DELIM );
			builder.append( bottomBorderColor );
			builder.append( DELIM );
			builder.append( dataFormat );
			builder.append( DELIM );
			builder.append( fillBackgroundColor );
			builder.append( DELIM );
			builder.append( fillForegroundColor );
			builder.append( DELIM );
			builder.append( fillPattern );
			builder.append( DELIM );
			builder.append( fontIndex );
			builder.append( DELIM );

			builder.append( hidden );
			builder.append( DELIM );
			builder.append( indention );
			builder.append( DELIM );
			builder.append( leftBorderColor );
			builder.append( DELIM );
			builder.append( locked );
			builder.append( DELIM );
			builder.append( rightBorderColor );
			builder.append( DELIM );
			builder.append( rotation );
			builder.append( DELIM );
			builder.append( topBorderColor );
			builder.append( DELIM );
			builder.append( verticalAlignment );
			builder.append( DELIM );
			builder.append( wrapText );
			return ( builder.toString() );
		}

		public void applyFillTo( HSSFCellStyle style ) {
			style.setFillPattern( this.fillPattern );
			style.setFillForegroundColor( this.fillForegroundColor );
			style.setFillBackgroundColor( this.fillBackgroundColor );
		}

		public void applyTo( HSSFCellStyle style, HSSFWorkbook workbook ) {

			style.setAlignment( this.alignment );
			style.setBorderBottom( this.borderBottom );
			style.setBorderLeft( this.borderLeft );
			style.setBorderRight( this.borderRight );
			style.setBorderTop( this.borderTop );
			style.setBottomBorderColor( this.bottomBorderColor );
			style.setDataFormat( this.dataFormat );

			style.setFillPattern( this.fillPattern );
			style.setFillForegroundColor( this.fillForegroundColor );
			style.setFillBackgroundColor( this.fillBackgroundColor );

			style.setFont( workbook.getFontAt( this.fontIndex ) );

			style.setHidden( this.hidden );
			style.setIndention( this.indention );
			style.setLeftBorderColor( this.leftBorderColor );
			style.setLocked( this.locked );
			style.setRightBorderColor( this.rightBorderColor );
			style.setRotation( this.rotation );
			style.setTopBorderColor( this.topBorderColor );
			style.setVerticalAlignment( this.verticalAlignment );
			style.setWrapText( this.wrapText );
		}

		/** not all JEXL (e.g. ternary conditional) works with JXLS, this DSL allows conditional alternative. */
		public StyleVO setNext( boolean setNext ) {
			this.setNext = setNext;
			return this;
		}

		public StyleVO setAlignment( short alignment ) {
			if ( !this.setNext ) {
				return this;
			}
			this.alignment = alignment;
			return this;
		}

		public StyleVO setBorderBottom( short borderBottom ) {
			if ( !this.setNext ) {
				return this;
			}
			this.borderBottom = borderBottom;
			return this;
		}

		public StyleVO setBorderLeft( short borderLeft ) {
			if ( !this.setNext ) {
				return this;
			}
			this.borderLeft = borderLeft;
			return this;
		}

		public StyleVO setBorderRight( short borderRight ) {
			if ( !this.setNext ) {
				return this;
			}
			this.borderRight = borderRight;
			return this;
		}

		public StyleVO setBorderTop( short borderTop ) {
			if ( !this.setNext ) {
				return this;
			}
			this.borderTop = borderTop;
			return this;
		}

		public StyleVO setBottomBorderColor( short bottomBorderColor ) {
			if ( !this.setNext ) {
				return this;
			}
			this.bottomBorderColor = bottomBorderColor;
			return this;
		}

		public StyleVO setDataFormat( short dataFormat ) {
			if ( !this.setNext ) {
				return this;
			}
			this.dataFormat = dataFormat;
			return this;
		}

		public StyleVO setDataFormat( HSSFWorkbook workbook, String dataFormatString ) {
			int builtinFormat = BuiltinFormats.getBuiltinFormat( dataFormatString );
			if ( !this.setNext ) {
				return this;
			}
			if ( builtinFormat == -1 ) {
				this.dataFormat = workbook.createDataFormat().getFormat( dataFormatString );
			} else {
				this.dataFormat = (short) builtinFormat;
			}
			return this;
		}

		public StyleVO setFillBackgroundColor( short fillBackgroundColor ) {
			if ( !this.setNext ) {
				return this;
			}
			this.fillBackgroundColor = fillBackgroundColor;
			return this;
		}

		public StyleVO setFillForegroundColor( short fillForegroundColor ) {
			if ( !this.setNext ) {
				return this;
			}
			this.fillForegroundColor = fillForegroundColor;
			return this;
		}

		public StyleVO setFillPattern( short fillPattern ) {
			if ( !this.setNext ) {
				return this;
			}
			this.fillPattern = fillPattern;
			return this;
		}

		public StyleVO setFontIndex( short fontIndex ) {
			if ( !this.setNext ) {
				return this;
			}
			this.fontIndex = fontIndex;
			return this;
		}

		public StyleVO setHidden( boolean hidden ) {
			if ( !this.setNext ) {
				return this;
			}
			this.hidden = hidden;
			return this;
		}

		public StyleVO setIndention( short indention ) {
			if ( !this.setNext ) {
				return this;
			}
			this.indention = indention;
			return this;
		}

		public StyleVO setLeftBorderColor( short leftBorderColor ) {
			if ( !this.setNext ) {
				return this;
			}
			this.leftBorderColor = leftBorderColor;
			return this;
		}

		public StyleVO setLocked( boolean locked ) {
			if ( !this.setNext ) {
				return this;
			}
			this.locked = locked;
			return this;
		}

		public StyleVO setRightBorderColor( short rightBorderColor ) {
			if ( !this.setNext ) {
				return this;
			}
			this.rightBorderColor = rightBorderColor;
			return this;
		}

		public StyleVO setRotation( short rotation ) {
			if ( !this.setNext ) {
				return this;
			}
			this.rotation = rotation;
			return this;
		}

		public StyleVO setTopBorderColor( short topBorderColor ) {
			if ( !this.setNext ) {
				return this;
			}
			this.topBorderColor = topBorderColor;
			return this;
		}

		public StyleVO setVerticalAlignment( short verticalAlignment ) {
			if ( !this.setNext ) {
				return this;
			}
			this.verticalAlignment = verticalAlignment;
			return this;
		}

		public StyleVO setWrapText( boolean wrapText ) {
			if ( !this.setNext ) {
				return this;
			}
			this.wrapText = wrapText;
			return this;
		}

		public StyleVO setFontName( String fontName ) {
			if ( !this.setNext ) {
				return this;
			}
			this.fontVal.fontName = fontName;
			return this;
		}

		public StyleVO setFontHeight( short fontHeight ) {
			if ( !this.setNext ) {
				return this;
			}
			this.fontVal.fontHeight = fontHeight;
			return this;
		}

		public StyleVO setFontBoldweight( short boldweight ) {
			if ( !this.setNext ) {
				return this;
			}
			this.fontVal.boldweight = boldweight;
			return this;
		}

		public StyleVO setFontItalic( boolean italic ) {
			if ( !this.setNext ) {
				return this;
			}
			this.fontVal.italic = italic;
			return this;
		}

		public StyleVO setFontStrikeout( boolean strikeout ) {
			if ( !this.setNext ) {
				return this;
			}
			this.fontVal.strikeout = strikeout;
			return this;
		}

		public StyleVO setFontTypeOffset( short typeOffset ) {
			if ( !this.setNext ) {
				return this;
			}
			this.fontVal.typeOffset = typeOffset;
			return this;
		}

		public StyleVO setFontUnderline( byte underline ) {
			if ( !this.setNext ) {
				return this;
			}
			this.fontVal.underline = underline;
			return this;
		}

		public StyleVO setFontColor( short color ) {
			if ( !this.setNext ) {
				return this;
			}
			this.fontVal.color = color;
			return this;
		}
	}

	public static class ReusableFonts {
		private Map<String, HSSFFont> reusableFonts = new HashMap<String, HSSFFont>();
		
		private HSSFFont getOrCreateFont( HSSFWorkbook workbook, FontVO fontVO ) {		
			String keyString = fontVO.toString();
			HSSFFont reusableFont = this.reusableFonts.get( keyString );
			if( reusableFont == null ) {
				reusableFont = workbook.createFont();
				fontVO.applyTo( reusableFont );
				reusableFonts.put( keyString, reusableFont );
			}
			return reusableFont;
		}
	}
	
	public static class ReusableStyles {
		private Map<String, HSSFCellStyle> reusableStyles = new HashMap<String, HSSFCellStyle>();
		
		public HSSFCellStyle getOrCreateStyle( HSSFWorkbook workbook, StyleVO styleVO ) {
			String keyString = styleVO.toString(); 
			HSSFCellStyle style = reusableStyles.get( keyString );

			if( style == null ) {
				style = workbook.createCellStyle();
				styleVO.applyTo( style, workbook );
				reusableStyles.put( keyString, style );
			}
			return style;
		}
	}
	
	public static class Formatter {
		
		private ReusableFonts fonts = new ReusableFonts();
		private ReusableStyles styles = new ReusableStyles();
		
		public StyleVO getStyle( HSSFCell cell ) {
			HSSFWorkbook workbook = ((HSSFCell)cell).getRow().getSheet().getWorkbook(); 
			return new StyleVO( cell.getCellStyle(), workbook );
		}
		
		public Object setStyle( Object cellVal, HSSFCell cell, StyleVO styleVal ) {
			HSSFWorkbook workbook = cell.getRow().getSheet().getWorkbook(); 
			
			if ( cellVal instanceof BigDecimal ) {
				cellVal = ((BigDecimal)cellVal).doubleValue();
			} else if ( cellVal instanceof Float ) {
				cellVal = new Double( cellVal.toString() );
			}
					
			HSSFFont font = fonts.getOrCreateFont( workbook, styleVal.fontVal );
			styleVal.fontIndex = font.getIndex();
			cell.setCellStyle( styles.getOrCreateStyle( workbook, styleVal ) );
			return cellVal;
		}
	}
    
}
