package net.sf.jxls.util;

import net.sf.jxls.parser.Cell;
import net.sf.jxls.transformer.Row;
import net.sf.jxls.transformer.RowCollection;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.hssf.util.CellReference;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.*;
import java.util.regex.Pattern;

/**
 * This class contains many utility methods used by jXLS framework
 * 
 * @author Leonid Vysochyn
 * @author Vincent Dutat
 */
public final class Util {
	protected static Log log = LogFactory.getLog(Util.class);

	private static final String[][] ENTITY_ARRAY = { { "quot", "34" }, // " -
																		// double
																		// -
																		// quote
			{ "amp", "38" }, // & - ampersand
			{ "lt", "60" }, // < - less-than
			{ "gt", "62" }, // > - greater-than
			{ "apos", "39" } // XML apostrophe
	};

	private static Map xmlEntities = new HashMap();

	static {
		for (int i = 0; i < ENTITY_ARRAY.length; i++) {
			xmlEntities.put(ENTITY_ARRAY[i][1], ENTITY_ARRAY[i][0]);
		}
	}

	public static void removeRowCollectionPropertiesFromRow(
			RowCollection rowCollection) {
		int startRow = rowCollection.getParentRow().getHssfRow().getRowNum();
		HSSFSheet sheet = rowCollection.getParentRow().getSheet()
				.getHssfSheet();
		for (int i = 0; i <= rowCollection.getDependentRowNumber(); i++) {
			HSSFRow hssfRow = sheet.getRow(startRow + i);
			for (int j = hssfRow.getFirstCellNum(); j <= hssfRow
					.getLastCellNum(); j++) {
				HSSFCell cell = hssfRow.getCell(j);
				removeRowCollectionPropertyFromCell(cell, rowCollection
						.getCollectionProperty().getFullCollectionName());
			}
		}
	}

	private static void removeRowCollectionPropertyFromCell(HSSFCell cell,
			String collectionName) {
		String regex = "[-+*/().A-Za-z_0-9\\s]*";
		if (cell != null && cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
			String cellValue = cell.getRichStringCellValue().getString();
			String strToReplace = "\\$\\{" + regex
					+ collectionName.replaceAll("\\.", "\\\\.") + "\\." + regex
					+ "\\}";
			cell.setCellValue(new HSSFRichTextString(cellValue.replaceAll(
					strToReplace, "")));
		}
	}

	/**
	 * Removes merged region from sheet
	 * 
	 * @param sheet
	 * @param region
	 */
	private static void removeMergedRegion(HSSFSheet sheet,
			CellRangeAddress region) {
		int index = getMergedRegionIndex(sheet, region);
		if (index >= 0) {
			sheet.removeMergedRegion(index);
		}
	}

	/**
	 * returns merged region index
	 * 
	 * @param sheet
	 * @param mergedRegion
	 * @return index of mergedRegion or -1 if the region not found
	 */
	private static int getMergedRegionIndex(HSSFSheet sheet,
			CellRangeAddress mergedRegion) {
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
			CellRangeAddress region = getMergedRegion(sheet, i);
			if (areRegionsEqual(region, mergedRegion)) {
				return i;
			}
		}
		return -1;
	}

	public static boolean areRegionsEqual(CellRangeAddress region1,
			CellRangeAddress region2) {
		if ((region1 == null && region2 != null)
				|| (region1 != null && region2 == null)) {
			return false;
		}
		if (region1 == null && region2 == null) {
			return true;
		}
		return (region1.getFirstColumn() == region2.getFirstColumn()
				&& region1.getLastColumn() == region2.getLastColumn()
				&& region1.getFirstRow() == region2.getFirstRow() && region2
				.getLastRow() == region2.getLastRow());
	}

	private static CellRangeAddress getMergedRegion(HSSFSheet sheet, int i) {
		CellRangeAddress region = sheet.getMergedRegion(i);
		return region;
	}

	private static boolean isNewMergedRegion(CellRangeAddress region,
			Collection mergedRegions) {
		for (Iterator iterator = mergedRegions.iterator(); iterator.hasNext();) {
			CellRangeAddress cellRangeAddress = (CellRangeAddress) iterator
					.next();
			if (areRegionsEqual(cellRangeAddress, region)) {
				return false;
			}
		}
		return true;
	}

	public static CellRangeAddress getMergedRegion(HSSFSheet sheet, int rowNum,
			int cellNum) {
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
			CellRangeAddress merged = getMergedRegion(sheet, i);
			if (isRangeContainsCell(merged, rowNum, cellNum)) {
				return merged;
			}
		}
		return null;
	}

	public static boolean isRangeContainsCell(CellRangeAddress range, int row,
			int col) {
		if ((range.getFirstRow() <= row) && (range.getLastRow() >= row)
				&& (range.getFirstColumn() <= col)
				&& (range.getLastColumn() >= col)) {
			return true;
		}
		return false;
	}

	public static boolean removeMergedRegion(HSSFSheet sheet, int rowNum,
			int cellNum) {
		Set mergedRegionNumbersToRemove = new TreeSet();
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
			CellRangeAddress merged = getMergedRegion(sheet, i);
			if (isRangeContainsCell(merged, rowNum, cellNum)) {
				mergedRegionNumbersToRemove.add(new Integer(i));
			}
		}
		for (Iterator iterator = mergedRegionNumbersToRemove.iterator(); iterator
				.hasNext();) {
			Integer regionNumber = (Integer) iterator.next();
			sheet.removeMergedRegion(regionNumber.intValue());
		}
		return !mergedRegionNumbersToRemove.isEmpty();
	}

	public static void prepareCollectionPropertyInRowForDuplication(
			RowCollection rowCollection, String collectionItemName) {
		int startRow = rowCollection.getParentRow().getHssfRow().getRowNum();
		HSSFSheet sheet = rowCollection.getParentRow().getSheet()
				.getHssfSheet();
		for (int i = 0; i <= rowCollection.getDependentRowNumber(); i++) {
			HSSFRow hssfRow = sheet.getRow(startRow + i);
			for (int j = hssfRow.getFirstCellNum(); j <= hssfRow
					.getLastCellNum(); j++) {
				HSSFCell cell = hssfRow.getCell(j);
				prepareCollectionPropertyInCellForDuplication(cell,
						rowCollection.getCollectionProperty()
								.getFullCollectionName(), collectionItemName);
			}
		}
	}

	private static void prepareCollectionPropertyInCellForDuplication(
			HSSFCell cell, String collectionName, String collectionItemName) {
		if (cell != null && cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
			String cellValue = cell.getRichStringCellValue().getString();
			String newValue = replaceCollectionProperty(cellValue,
					collectionName, collectionItemName);
			// String newValue = cellValue.replaceFirst(collectionName,
			// collectionItemName);
			cell.setCellValue(new HSSFRichTextString(newValue));
		}
	}

	private static String replaceCollectionProperty(String property,
			String collectionName, String newValue) {
		return property.replaceAll(collectionName, newValue);
	}

	public static void prepareCollectionPropertyInRowForContentDuplication(
			RowCollection rowCollection) {
		for (int i = 0; i < rowCollection.getCells().size(); i++) {
			Cell cell = (Cell) rowCollection.getCells().get(i);
			prepareCollectionPropertyInCellForDuplication(cell.getHssfCell(),
					rowCollection.getCollectionProperty()
							.getFullCollectionName(), rowCollection
							.getCollectionItemName());
		}
	}

	public static void duplicateRowCollectionProperty(
			RowCollection rowCollection) {
		Collection collection = rowCollection.getCollectionProperty()
				.getCollection();
		int rowNum = rowCollection.getParentRow().getHssfRow().getRowNum();
		HSSFRow srcRow = rowCollection.getParentRow().getHssfRow();
		HSSFSheet sheet = rowCollection.getParentRow().getSheet()
				.getHssfSheet();
		if (collection.size() > 1) {
			for (int i = 1; i < collection.size(); i++) {
				HSSFRow destRow = sheet.getRow(rowNum + i);
				for (int j = 0; j < rowCollection.getCells().size(); j++) {
					Cell cell = (Cell) rowCollection.getCells().get(j);
					if (!cell.isEmpty()) {
						HSSFCell destCell = destRow.getCell(cell.getHssfCell()
								.getColumnIndex());
						if (destCell == null) {
							destCell = destRow.createCell(cell.getHssfCell()
									.getColumnIndex());
						}
						copyCell(srcRow.getCell(cell.getHssfCell()
								.getColumnIndex()), destCell, false);
					}
				}
			}
		}
	}

	public static int duplicateRow(RowCollection rowCollection) {
		Collection collection = rowCollection.getCollectionProperty()
				.getCollection();
		int row = rowCollection.getParentRow().getHssfRow().getRowNum();
		HSSFSheet sheet = rowCollection.getParentRow().getSheet()
				.getHssfSheet();
		if (collection.size() > 1) {
			if (rowCollection.getDependentRowNumber() == 0) {
				shiftRows(sheet, row + 1, sheet.getLastRowNum(), collection
						.size() - 1);
				duplicateStyle(rowCollection, row, row + 1,
						collection.size() - 1);
				shiftUncoupledCellsUp(rowCollection);
			} else {
				for (int i = 0; i < collection.size() - 1; i++) {
					shiftCopyRowCollection(rowCollection);
				}
				shiftUncoupledCellsUp(rowCollection);
			}
		}
		return (collection.size() - 1)
				* (rowCollection.getDependentRowNumber() + 1);
	}

	private static void shiftCopyRowCollection(RowCollection rowCollection) {
		HSSFSheet hssfSheet = rowCollection.getParentRow().getSheet()
				.getHssfSheet();
		int startRow = rowCollection.getParentRow().getHssfRow().getRowNum();
		int num = rowCollection.getDependentRowNumber();
		shiftRows(hssfSheet, startRow + num + 1, hssfSheet.getLastRowNum(),
				num + 1);
		copyRowCollection(rowCollection);
	}

	private static void copyRowCollection(RowCollection rowCollection) {
		HSSFSheet sheet = rowCollection.getParentRow().getSheet()
				.getHssfSheet();
		int from = rowCollection.getParentRow().getHssfRow().getRowNum();
		int num = rowCollection.getDependentRowNumber() + 1;
		int to = from + num;
		Set mergedRegions = new HashSet();
		for (int i = from; i < from + num; i++) {
			HSSFRow srcRow = sheet.getRow(i);
			HSSFRow destRow = sheet.getRow(to + i - from);
			if (destRow == null) {
				destRow = sheet.createRow(to + i - from);
			}
			if (srcRow.getHeight() >= 0) {
				destRow.setHeight(srcRow.getHeight());
			}
			for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
				HSSFCell srcCell = srcRow.getCell(j);
				if (srcCell != null) {
					HSSFCell destCell = destRow.createCell(j);
					copyCell(srcCell, destCell, true);
					CellRangeAddress mergedRegion = getMergedRegion(sheet, i, j);
					if (mergedRegion != null) {
						CellRangeAddress newMergedRegion = new CellRangeAddress(
								to - from + mergedRegion.getFirstRow(), to
										- from + mergedRegion.getLastRow(),
								mergedRegion.getFirstColumn(), mergedRegion
										.getLastColumn());
						if (isNewMergedRegion(newMergedRegion, mergedRegions)) {
							mergedRegions.add(newMergedRegion);
						}
					}
				}
			}
		}
		// set merged regions
		for (Iterator iterator = mergedRegions.iterator(); iterator.hasNext();) {
			CellRangeAddress region = (CellRangeAddress) iterator.next();
			sheet.addMergedRegion(region);
		}
	}

	private static void shiftUncoupledCellsUp(RowCollection rowCollection) {
		Row row = rowCollection.getParentRow();
		if (row.getCells().size() > rowCollection.getCells().size()) {
			for (int i = 0; i < row.getCells().size(); i++) {
				Cell cell = (Cell) row.getCells().get(i);
				if (!rowCollection.containsCell(cell)) {
					shiftColumnUp(cell, row.getHssfRow().getRowNum()
							+ rowCollection.getCollectionProperty()
									.getCollection().size(), rowCollection
							.getCollectionProperty().getCollection().size() - 1);
				}
			}
		}
	}

	private static void shiftColumnUp(Cell cell, int startRow, int shiftNumber) {
		HSSFSheet sheet = cell.getRow().getSheet().getHssfSheet();
		int cellNum = cell.getHssfCell().getColumnIndex();
		List hssfMergedRegions = new ArrayList();
		// find all merged regions in this area
		for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
			CellRangeAddress region = getMergedRegion(sheet, i, cellNum);
			if (region != null && isNewMergedRegion(region, hssfMergedRegions)) {
				hssfMergedRegions.add(region);
			}
		}
		// move all related cells up
		for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
			if (sheet.getRow(i).getCell(cellNum) != null) {
				HSSFCell destCell = sheet.getRow(i - shiftNumber).getCell(
						cellNum);
				if (destCell == null) {
					destCell = sheet.getRow(i - shiftNumber)
							.createCell(cellNum);
				}
				moveCell(sheet.getRow(i).getCell(cellNum), destCell);
			}
		}
		// remove previously shifted merged regions in this area
		for (Iterator iterator = hssfMergedRegions.iterator(); iterator
				.hasNext();) {
			removeMergedRegion(sheet, (CellRangeAddress) iterator.next());
		}
		// set merged regions for shifted cells
		for (Iterator iterator = hssfMergedRegions.iterator(); iterator
				.hasNext();) {
			CellRangeAddress region = (CellRangeAddress) iterator.next();
			CellRangeAddress newRegion = new CellRangeAddress(region
					.getFirstRow()
					- shiftNumber, region.getLastRow() - shiftNumber, region
					.getFirstColumn(), region.getLastColumn());
			sheet.addMergedRegion(newRegion);
		}
		// remove moved cells
		int i = sheet.getLastRowNum();
		while (sheet.getRow(i).getCell(cellNum) == null && i >= startRow) {
			i--;
		}
		for (int j = 0; j < shiftNumber && i >= startRow; j++, i--) {
			if (sheet.getRow(i).getCell(cellNum) != null) {
				sheet.getRow(i).removeCell(sheet.getRow(i).getCell(cellNum));
			}
		}
	}

	private static void moveCell(HSSFCell srcCell, HSSFCell destCell) {
		destCell.setCellStyle(srcCell.getCellStyle());
		switch (srcCell.getCellType()) {
		case HSSFCell.CELL_TYPE_STRING:
			destCell.setCellValue(srcCell.getRichStringCellValue());
			break;
		case HSSFCell.CELL_TYPE_NUMERIC:
			destCell.setCellValue(srcCell.getNumericCellValue());
			break;
		case HSSFCell.CELL_TYPE_BLANK:
			destCell.setCellType(HSSFCell.CELL_TYPE_BLANK);
			break;
		case HSSFCell.CELL_TYPE_BOOLEAN:
			destCell.setCellValue(srcCell.getBooleanCellValue());
			break;
		case HSSFCell.CELL_TYPE_ERROR:
			destCell.setCellErrorValue(srcCell.getErrorCellValue());
			break;
		case HSSFCell.CELL_TYPE_FORMULA:
			break;
		default:
			break;
		}
		srcCell.setCellType(HSSFCell.CELL_TYPE_BLANK);
	}

	private static void duplicateStyle(RowCollection rowCollection,
			int rowToCopy, int startRow, int num) {
		HSSFSheet sheet = rowCollection.getParentRow().getSheet()
				.getHssfSheet();
		Set mergedRegions = new HashSet();
		HSSFRow srcRow = sheet.getRow(rowToCopy);
		for (int i = startRow; i < startRow + num; i++) {
			HSSFRow destRow = sheet.getRow(i);
			if (destRow == null) {
				destRow = sheet.createRow(i);
			}
			if (srcRow.getHeight() >= 0) {
				destRow.setHeight(srcRow.getHeight());
			}
			for (int j = 0; j < rowCollection.getCells().size(); j++) {
				Cell cell = (Cell) rowCollection.getCells().get(j);
				HSSFCell hssfCell = cell.getHssfCell();
				if (hssfCell != null) {
					HSSFCell newCell = destRow.createCell(hssfCell
							.getColumnIndex());
					copyCell(hssfCell, newCell, true);
					CellRangeAddress mergedRegion = getMergedRegion(sheet,
							rowToCopy, hssfCell.getColumnIndex());
					if (mergedRegion != null) {
						CellRangeAddress newMergedRegion = new CellRangeAddress(
								i, i + mergedRegion.getLastRow()
										- mergedRegion.getFirstRow(),
								mergedRegion.getFirstColumn(), mergedRegion
										.getLastColumn());
						if (isNewMergedRegion(newMergedRegion, mergedRegions)) {
							mergedRegions.add(newMergedRegion);
							sheet.addMergedRegion(newMergedRegion);
						}
					}
				}
			}
		}
	}

	public static void copyRow(HSSFSheet sheet, HSSFRow oldRow, HSSFRow newRow) {
		Set mergedRegions = new HashSet();
		if (oldRow.getHeight() >= 0) {
			newRow.setHeight(oldRow.getHeight());
		}
		for (int j = oldRow.getFirstCellNum(); j <= oldRow.getLastCellNum(); j++) {
			HSSFCell oldCell = oldRow.getCell(j);
			HSSFCell newCell = newRow.getCell(j);
			if (oldCell != null) {
				if (newCell == null) {
					newCell = newRow.createCell(j);
				}
				copyCell(oldCell, newCell, true);
				CellRangeAddress mergedRegion = getMergedRegion(sheet, oldRow
						.getRowNum(), oldCell.getColumnIndex());
				if (mergedRegion != null) {
					CellRangeAddress newMergedRegion = new CellRangeAddress(
							newRow.getRowNum(), newRow.getRowNum()
									+ mergedRegion.getLastRow()
									- mergedRegion.getFirstRow(), mergedRegion
									.getFirstColumn(), mergedRegion
									.getLastColumn());
					if (isNewMergedRegion(newMergedRegion, mergedRegions)) {
						mergedRegions.add(newMergedRegion);
						sheet.addMergedRegion(newMergedRegion);
					}
				}
			}
		}
	}

	public static void copyRow(HSSFSheet srcSheet, HSSFSheet destSheet,
			HSSFRow srcRow, HSSFRow destRow) {
		Set mergedRegions = new TreeSet();
		if (srcRow.getHeight() >= 0) {
			destRow.setHeight(srcRow.getHeight());
		}
		for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
			HSSFCell oldCell = srcRow.getCell(j);
			HSSFCell newCell = destRow.getCell(j);
			if (oldCell != null) {
				if (newCell == null) {
					newCell = destRow.createCell(j);
				}
				copyCell(oldCell, newCell, true);
				CellRangeAddress mergedRegion = getMergedRegion(srcSheet,
						srcRow.getRowNum(), oldCell.getColumnIndex());
				if (mergedRegion != null) {
					// Region newMergedRegion = new Region( destRow.getRowNum(),
					// mergedRegion.getColumnFrom(),
					// destRow.getRowNum() + mergedRegion.getRowTo() -
					// mergedRegion.getRowFrom(), mergedRegion.getColumnTo() );
					CellRangeAddress newMergedRegion = new CellRangeAddress(
							mergedRegion.getFirstRow(), mergedRegion
									.getLastRow(), mergedRegion
									.getFirstColumn(), mergedRegion
									.getLastColumn());
					if (isNewMergedRegion(newMergedRegion, mergedRegions)) {
						mergedRegions.add(newMergedRegion);
						destSheet.addMergedRegion(newMergedRegion);
					}
				}
			}
		}
	}

	public static void copyRow(HSSFSheet srcSheet, HSSFSheet destSheet,
			HSSFRow srcRow, HSSFRow destRow, String expressionToReplace,
			String expressionReplacement) {
		Set mergedRegions = new HashSet();
		if (srcRow.getHeight() >= 0) {
			destRow.setHeight(srcRow.getHeight());
		}
		for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
			HSSFCell oldCell = srcRow.getCell(j);
			HSSFCell newCell = destRow.getCell(j);
			if (oldCell != null) {
				if (newCell == null) {
					newCell = destRow.createCell(j);
				}
				copyCell(oldCell, newCell, true, expressionToReplace,
						expressionReplacement);
				CellRangeAddress mergedRegion = getMergedRegion(srcSheet,
						srcRow.getRowNum(), oldCell.getColumnIndex());
				if (mergedRegion != null) {
					// Region newMergedRegion = new Region( destRow.getRowNum(),
					// mergedRegion.getColumnFrom(),
					// destRow.getRowNum() + mergedRegion.getRowTo() -
					// mergedRegion.getRowFrom(), mergedRegion.getColumnTo() );
					CellRangeAddress newMergedRegion = new CellRangeAddress(
							mergedRegion.getFirstRow(), mergedRegion
									.getLastRow(), mergedRegion
									.getFirstColumn(), mergedRegion
									.getLastColumn());
					if (isNewMergedRegion(newMergedRegion, mergedRegions)) {
						mergedRegions.add(newMergedRegion);
						destSheet.addMergedRegion(newMergedRegion);
					}
				}
			}
		}
	}

	public static void copySheets(HSSFSheet newSheet, HSSFSheet sheet) {
		int maxColumnNum = 0;
		for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
			HSSFRow srcRow = sheet.getRow(i);
			HSSFRow destRow = newSheet.createRow(i);
			if (srcRow != null) {
				Util.copyRow(sheet, newSheet, srcRow, destRow);
				if (srcRow.getLastCellNum() > maxColumnNum) {
					maxColumnNum = srcRow.getLastCellNum();
				}
			}
		}
		for (int i = 0; i <= maxColumnNum; i++) {
			newSheet.setColumnWidth(i, sheet.getColumnWidth(i));
		}
	}

	public static void copySheets(HSSFSheet newSheet, HSSFSheet sheet,
			String expressionToReplace, String expressionReplacement) {
		int maxColumnNum = 0;
		for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
			HSSFRow srcRow = sheet.getRow(i);
			HSSFRow destRow = newSheet.createRow(i);
			if (srcRow != null) {
				Util.copyRow(sheet, newSheet, srcRow, destRow,
						expressionToReplace, expressionReplacement);
				if (srcRow.getLastCellNum() > maxColumnNum) {
					maxColumnNum = srcRow.getLastCellNum();
				}
			}
		}
		for (int i = 0; i <= maxColumnNum; i++) {
			newSheet.setColumnWidth(i, sheet.getColumnWidth(i));
		}
	}

	public static void copyCell(HSSFCell oldCell, HSSFCell newCell,
			boolean copyStyle) {
		if (copyStyle) {
			newCell.setCellStyle(oldCell.getCellStyle());
		}
		switch (oldCell.getCellType()) {
		case HSSFCell.CELL_TYPE_STRING:
			newCell.setCellValue(oldCell.getRichStringCellValue());
			break;
		case HSSFCell.CELL_TYPE_NUMERIC:
			newCell.setCellValue(oldCell.getNumericCellValue());
			break;
		case HSSFCell.CELL_TYPE_BLANK:
			newCell.setCellType(HSSFCell.CELL_TYPE_BLANK);
			break;
		case HSSFCell.CELL_TYPE_BOOLEAN:
			newCell.setCellValue(oldCell.getBooleanCellValue());
			break;
		case HSSFCell.CELL_TYPE_ERROR:
			newCell.setCellErrorValue(oldCell.getErrorCellValue());
			break;
		case HSSFCell.CELL_TYPE_FORMULA:
			newCell.setCellFormula(oldCell.getCellFormula());
			break;
		default:
			break;
		}
	}

	public static void copyCell(HSSFCell oldCell, HSSFCell newCell,
			boolean copyStyle, String expressionToReplace,
			String expressionReplacement) {
		if (copyStyle) {
			newCell.setCellStyle(oldCell.getCellStyle());
		}
		switch (oldCell.getCellType()) {
		case HSSFCell.CELL_TYPE_STRING:
			String oldValue = oldCell.getRichStringCellValue().getString();
			newCell.setCellValue(new HSSFRichTextString(oldValue.replaceAll(
					expressionToReplace, expressionReplacement)));
			break;
		case HSSFCell.CELL_TYPE_NUMERIC:
			newCell.setCellValue(oldCell.getNumericCellValue());
			break;
		case HSSFCell.CELL_TYPE_BLANK:
			newCell.setCellType(HSSFCell.CELL_TYPE_BLANK);
			break;
		case HSSFCell.CELL_TYPE_BOOLEAN:
			newCell.setCellValue(oldCell.getBooleanCellValue());
			break;
		case HSSFCell.CELL_TYPE_ERROR:
			newCell.setCellErrorValue(oldCell.getErrorCellValue());
			break;
		case HSSFCell.CELL_TYPE_FORMULA:
			newCell.setCellFormula(oldCell.getCellFormula());
			break;
		default:
			break;
		}
	}

	public static Object getProperty(Object bean, String propertyName) {
		Object value = null;
		try {
			if (log.isDebugEnabled()) {
				log.debug("getting property=" + propertyName + " for bean="
						+ bean.getClass().getName());
			}
			value = PropertyUtils.getProperty(bean, propertyName);
		} catch (IllegalAccessException e) {
			log.warn("Can't get property " + propertyName + " in the bean "
					+ bean, e);
		} catch (InvocationTargetException e) {
			log.warn("Can't get property " + propertyName + " in the bean "
					+ bean, e);
		} catch (NoSuchMethodException e) {
			log.warn("Can't get property " + propertyName + " in the bean "
					+ bean, e);
		}
		return value;

	}

	/**
	 * Saves workbook to file
	 * 
	 * @param fileName
	 *            - File name to save workbook
	 * @param workbook
	 *            - Workbook to save
	 */
	public static void writeToFile(String fileName, HSSFWorkbook workbook) {
		OutputStream os;
		try {
			os = new BufferedOutputStream(new FileOutputStream(fileName));
			workbook.write(os);
			os.flush();
			os.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * Duplicates given HSSFCellStyle object
	 * 
	 * @param workbook
	 *            - source HSSFWorkbook object
	 * @param style
	 *            - HSSFCellStyle object to duplicate
	 * @return HSSFCellStyle
	 */
	public static HSSFCellStyle duplicateStyle(HSSFWorkbook workbook,
			HSSFCellStyle style) {
		HSSFCellStyle newStyle = workbook.createCellStyle();
		newStyle.setAlignment(style.getAlignment());
		newStyle.setBorderBottom(style.getBorderBottom());
		newStyle.setBorderLeft(style.getBorderLeft());
		newStyle.setBorderRight(style.getBorderRight());
		newStyle.setBorderTop(style.getBorderTop());
		newStyle.setBottomBorderColor(style.getBottomBorderColor());
		newStyle.setDataFormat(style.getDataFormat());
		newStyle.setFillBackgroundColor(style.getFillBackgroundColor());
		newStyle.setFillForegroundColor(style.getFillForegroundColor());
		newStyle.setFillPattern(style.getFillPattern());
		newStyle.setFont(workbook.getFontAt(style.getFontIndex()));
		newStyle.setHidden(style.getHidden());
		newStyle.setIndention(style.getIndention());
		newStyle.setLeftBorderColor(style.getLeftBorderColor());
		newStyle.setLocked(style.getLocked());
		newStyle.setRightBorderColor(style.getRightBorderColor());
		newStyle.setTopBorderColor(style.getTopBorderColor());
		newStyle.setVerticalAlignment(style.getVerticalAlignment());
		newStyle.setWrapText(style.getWrapText());
		return newStyle;
	}

	public static String escapeAttributes(String tag) {
		if (tag == null) {
			return tag;
		}
		int i = 0;
		StringBuffer sb = new StringBuffer("");
		StringBuffer attrValue = new StringBuffer("");
		final char expressionClosingSymbol = '}';
		final char expressionStartSymbol = '{';
		boolean isAttrValue = false;
		int exprCount = 0;
		while (i < tag.length()) {
			if (!isAttrValue) {
				sb.append(tag.charAt(i));
				if (tag.charAt(i) == '\"') {
					isAttrValue = true;
					attrValue = new StringBuffer("");
				}
			} else {
				if (tag.charAt(i) == '\"') {
					if (exprCount != 0) {
						attrValue.append(tag.charAt(i));
					} else {
						sb.append(escapeXml(attrValue.toString()));
						sb.append(tag.charAt(i));
						isAttrValue = false;
					}
				} else {
					attrValue.append(tag.charAt(i));
					if (tag.charAt(i) == expressionClosingSymbol) {
						exprCount--;
					} else if (tag.charAt(i) == expressionStartSymbol) {
						exprCount++;
					}
				}
			}
			i++;
		}
		if (isAttrValue) {
			log.warn("Can't parse ambiguous quot in " + tag);
		}
		return sb.toString();
	}

	/**
	 * <p>
	 * Escapes XML entities in a <code>String</code>.
	 * </p>
	 * 
	 * @param str
	 *            The <code>String</code> to escape.
	 * @return A new escaped <code>String</code>.
	 */
	private static String escapeXml(String str) {
		if (str == null) {
			return str;
		}
		StringBuffer buf = new StringBuffer(str.length() * 2);
		int i;
		for (i = 0; i < str.length(); ++i) {
			char ch = str.charAt(i);
			String entityName = getEntityName(ch);
			if (entityName == null) {
				if (ch > 0x7F) {
					buf.append("&#");
					buf.append((int) ch);
					buf.append(';');
				} else {
					buf.append(ch);
				}
			} else {
				buf.append('&');
				buf.append(entityName);
				buf.append(';');
			}
		}
		return buf.toString();
	}

	private static String getEntityName(char ch) {
		return (String) xmlEntities.get(Integer.toString(ch));
	}

	public static void shiftCellsLeft(HSSFSheet sheet, int startRow,
			int startCol, int endRow, int endCol, int shiftNumber) {
		for (int i = startRow; i <= endRow; i++) {
			boolean doSetWidth = true;
			HSSFRow row = sheet.getRow(i);
			if (row != null) {
				for (int j = startCol; j <= endCol; j++) {
					HSSFCell cell = row.getCell(j);
					if (cell == null) {
						cell = row.createCell(j);
						doSetWidth = false;
					}
					HSSFCell destCell = row.getCell((int) (j - shiftNumber));
					if (destCell == null) {
						destCell = row.createCell((int) (j - shiftNumber));
					}
					copyCell(cell, destCell, true);
					if (doSetWidth) {
						sheet.setColumnWidth(destCell.getColumnIndex(),
								getWidth(sheet, cell.getColumnIndex()));
					}
				}
			}
		}
	}

	static int getWidth(HSSFSheet sheet, int col) {
		int width = sheet.getColumnWidth(col);
		if (width == sheet.getDefaultColumnWidth()) {
			width = (int) (width * 256);
		}
		return width;
	}

	public static void shiftCellsRight(HSSFSheet sheet, int startRow,
			int endRow, int startCol, int shiftNumber) {
		for (int i = startRow; i <= endRow; i++) {
			HSSFRow row = sheet.getRow(i);
			if (row != null) {
				int lastCellNum = row.getLastCellNum();
				for (int j = lastCellNum; j >= startCol; j--) {
					HSSFCell destCell = row.getCell((int) (j + shiftNumber));
					if (destCell == null) {
						destCell = row.createCell((int) (j + shiftNumber));
					}
					HSSFCell cell = row.getCell(j);
					if (cell == null) {
						cell = row.createCell(j);
					}
					copyCell(cell, destCell, true);
				}
			}
		}
	}

	public static void updateCellValue(HSSFSheet sheet, int rowNum, int colNum,
			String cellValue) {
		HSSFRow hssfRow = sheet.getRow(rowNum);
		HSSFCell hssfCell = hssfRow.getCell(colNum);
		hssfCell.setCellValue(new HSSFRichTextString(cellValue));
	}

	public static void copyPageSetup(HSSFSheet destSheet, HSSFSheet srcSheet) {
		HSSFHeader header = srcSheet.getHeader();
		HSSFFooter footer = srcSheet.getFooter();
		if (footer != null) {
			destSheet.getFooter().setLeft(footer.getLeft());
			destSheet.getFooter().setCenter(footer.getCenter());
			destSheet.getFooter().setRight(footer.getRight());
		}
		if (header != null) {
			destSheet.getHeader().setLeft(header.getLeft());
			destSheet.getHeader().setCenter(header.getCenter());
			destSheet.getHeader().setRight(header.getRight());
		}
	}

	public static void copyPrintSetup(HSSFSheet destSheet, HSSFSheet srcSheet) {
		HSSFPrintSetup setup = srcSheet.getPrintSetup();
		if (setup != null) {
			destSheet.getPrintSetup().setLandscape(setup.getLandscape());
			destSheet.getPrintSetup().setPaperSize(setup.getPaperSize());
			destSheet.getPrintSetup().setScale(setup.getScale());
			destSheet.getPrintSetup().setFitWidth(setup.getFitWidth());
			destSheet.getPrintSetup().setFitHeight(setup.getFitHeight());
			destSheet.getPrintSetup().setFooterMargin(setup.getFooterMargin());
			destSheet.getPrintSetup().setHeaderMargin(setup.getHeaderMargin());
			destSheet.getPrintSetup().setPaperSize(setup.getPaperSize());
			destSheet.getPrintSetup().setPageStart(setup.getPageStart());
		}
	}

	public static void setPrintArea(HSSFWorkbook resultWorkbook, int sheetNum) {
		int maxColumnNum = 0;
		for (int j = resultWorkbook.getSheetAt(sheetNum).getFirstRowNum(); j <= resultWorkbook
				.getSheetAt(sheetNum).getLastRowNum(); j++) {
			HSSFRow row = resultWorkbook.getSheetAt(sheetNum).getRow(j);
			if (row != null) {
				maxColumnNum = row.getLastCellNum();
			}
		}
		resultWorkbook.setPrintArea(sheetNum, 0, maxColumnNum, 0,
				resultWorkbook.getSheetAt(sheetNum).getLastRowNum());
	}

	protected static final String regexCellRef = "[a-zA-Z]+[0-9]+";
	protected static final Pattern regexCellRefPattern = Pattern
			.compile(regexCellRef);
	protected static final String regexCellCharPart = "[0-9]+";
	protected static final String regexCellDigitPart = "[a-zA-Z]+";
	protected static String cellRangeSeparator = ":";

	public static boolean isColumnRange(List cells) {
		String firstCell = (String) cells.get(0);
		boolean isColumnRange = true;
		if (firstCell != null && firstCell.length() > 0) {
			String firstCellCharPart = firstCell.split(regexCellCharPart)[0];
			String firstCellDigitPart = firstCell.split(regexCellDigitPart)[1];
			int cellNumber = Integer.parseInt(firstCellDigitPart);
			String nextCell, cellCharPart, cellDigitPart;
			for (int i = 1; i < cells.size() && isColumnRange; i++) {
				nextCell = (String) cells.get(i);
				cellCharPart = nextCell.split(regexCellCharPart)[0];
				cellDigitPart = nextCell.split(regexCellDigitPart)[1];
				if (!firstCellCharPart.equalsIgnoreCase(cellCharPart)
						|| Integer.parseInt(cellDigitPart) != ++cellNumber) {
					isColumnRange = false;
				}
			}
		}
		return isColumnRange;
	}

	public static boolean isRowRange(List cells) {
		String firstCell = (String) cells.get(0);
		boolean isRowRange = true;
		if (firstCell != null && firstCell.length() > 0) {
			String firstCellDigitPart = firstCell.split(regexCellDigitPart)[1];
			String nextCell, cellDigitPart;
			CellReference cellRef = new CellReference(firstCell);
			int cellNumber = cellRef.getCol();
			for (int i = 1; i < cells.size() && isRowRange; i++) {
				nextCell = (String) cells.get(i);
				cellDigitPart = nextCell.split(regexCellDigitPart)[1];
				cellRef = new CellReference(nextCell);
				if (!firstCellDigitPart.equalsIgnoreCase(cellDigitPart)
						|| cellRef.getCol() != ++cellNumber) {
					isRowRange = false;
				}
			}
		}
		return isRowRange;
	}

	public static String buildCommaSeparatedListOfCells(String refSheetName,
			List cells) {
		String listOfCells = "";
		for (int i = 0; i < cells.size() - 1; i++) {
			String cell = (String) cells.get(i);
			listOfCells += getRefCellName(refSheetName, cell) + ",";
		}
		listOfCells += getRefCellName(refSheetName, (String) cells.get(cells
				.size() - 1));
		return listOfCells;
	}

	public static String detectCellRange(String refSheetName, List cells) {
		if (cells == null || cells.isEmpty()) {
			return "";
		}
		String firstCell = (String) cells.get(0);
		String range = firstCell;
		if (firstCell != null && firstCell.length() > 0) {
			if (isRowRange(cells) || isColumnRange(cells)) {
				String lastCell = (String) cells.get(cells.size() - 1);
				range = getRefCellName(refSheetName, firstCell)
						+ cellRangeSeparator + lastCell.toUpperCase();
			} else {
				range = buildCommaSeparatedListOfCells(refSheetName, cells);
			}
		}
		return range;
	}

	public static String getRefCellName(String refSheetName, String cellName) {
		if (refSheetName == null) {
			return cellName.toUpperCase();
		}
		return refSheetName + "!" + cellName.toUpperCase();
	}

	public static void shiftRows(HSSFSheet sheet, int startRow, int endRow,
			int shiftNum) {
		if (startRow <= endRow) {
			short[] rowHeights = getRowHeights(sheet, startRow, endRow);
			sheet.shiftRows(startRow, endRow, shiftNum, false, false);
			copyPositiveRowHeight(sheet, startRow, endRow, shiftNum, rowHeights);
		}
	}

	private static short[] getRowHeights(HSSFSheet sheet, int startRow,
			int endRow) {
		if (endRow - startRow + 1 < 0) {
			return new short[0];
		}
		short[] rowHeights = new short[endRow - startRow + 1];
		for (int i = startRow; i <= endRow; i++) {
			HSSFRow row = sheet.getRow(i);
			if (row != null) {
				rowHeights[i - startRow] = row.getHeight();
			} else {
				rowHeights[i - startRow] = -1;
			}
		}
		return rowHeights;
	}

	static void copyPositiveRowHeight(HSSFSheet sheet, int startRow,
			int endRow, int shiftNum, short[] rowHeights) {
		for (int i = startRow; i <= endRow; i++) {
			HSSFRow destRow = sheet.getRow(i + shiftNum);
			if (destRow != null && rowHeights[i - startRow] >= 0) {
				destRow.setHeight(rowHeights[i - startRow]);
			}
			// HSSFRow srcRow = sheet.getRow(i);
			// if( srcRow != null && destRow != null ){
			// if( srcRow.getHeight() >= 0 ){
			// destRow.setHeight( srcRow.getHeight() );
			// }
			// }
		}
	}
}
