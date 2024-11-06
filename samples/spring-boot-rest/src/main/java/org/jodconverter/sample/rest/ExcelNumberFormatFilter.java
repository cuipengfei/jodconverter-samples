package org.jodconverter.sample.rest;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.lang.Locale;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.sheet.XCellRangeAddressable;
import com.sun.star.sheet.XSheetCellCursor;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XUsedAreaCursor;
import com.sun.star.table.CellContentType;
import com.sun.star.table.CellRangeAddress;
import com.sun.star.table.XCell;
import com.sun.star.uno.AnyConverter;
import com.sun.star.util.MalformedNumberFormatException;
import com.sun.star.util.XNumberFormats;
import com.sun.star.util.XNumberFormatsSupplier;
import org.jodconverter.core.office.OfficeContext;
import org.jodconverter.local.filter.Filter;
import org.jodconverter.local.filter.FilterChain;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.math.BigDecimal;

import static com.sun.star.table.CellContentType.FORMULA;
import static com.sun.star.table.CellContentType.VALUE;
import static com.sun.star.uno.UnoRuntime.queryInterface;
import static java.math.BigDecimal.valueOf;

public class ExcelNumberFormatFilter implements Filter {
    private static final Logger log = LoggerFactory.getLogger(ExcelNumberFormatFilter.class);

    @Override
    public void doFilter(OfficeContext context, XComponent document, FilterChain chain) throws Exception {
        XSpreadsheetDocument xSpreadsheetDocument = queryInterface(XSpreadsheetDocument.class, document);
        if (xSpreadsheetDocument == null) {
            chain.doFilter(context, document);
            return;
        }

        String[] sheetNames = xSpreadsheetDocument.getSheets().getElementNames();
        XNumberFormatsSupplier xNumberFormatsSupplier = queryInterface(XNumberFormatsSupplier.class, xSpreadsheetDocument);
        XNumberFormats xNumberFormats = xNumberFormatsSupplier.getNumberFormats();
        for (String sheetName : sheetNames) {
            XSpreadsheet sheet = queryInterface(XSpreadsheet.class, xSpreadsheetDocument.getSheets().getByName(sheetName));
            if (isSheetVisible(sheet)) {
                processSheet(sheet, xNumberFormats);
            }
        }

        chain.doFilter(context, document);
    }

    private void processSheet(XSpreadsheet sheet, XNumberFormats xNumberFormats) throws Exception {
        XSheetCellCursor cursor = sheet.createCursor();
        XUsedAreaCursor usedAreaCursor = queryInterface(XUsedAreaCursor.class, cursor);
        usedAreaCursor.gotoEndOfUsedArea(true);

        CellRangeAddress rangeAddress = getCellRangeAddress(usedAreaCursor);

        for (int col = 0; col <= rangeAddress.EndColumn; col++) {
            for (int row = 0; row <= rangeAddress.EndRow; row++) {
                XCell cell = sheet.getCellByPosition(col, row);
                processCell(cell, xNumberFormats);
            }
        }
    }

    private void processCell(XCell cell, XNumberFormats xNumberFormats) {
        try {
            CellContentType cellType = cell.getType();
            if (cellType == VALUE || cellType == FORMULA) {
                XPropertySet cellProps = queryInterface(XPropertySet.class, cell);
                int key = AnyConverter.toInt(cellProps.getPropertyValue("NumberFormat"));
                XPropertySet numberFormat = xNumberFormats.getByKey(key);
                Locale locale = (Locale) numberFormat.getPropertyValue("Locale");
                String formatString = numberFormat.getPropertyValue("FormatString").toString();

                if (formatString.equals("General")) {
                    BigDecimal cellValue = new BigDecimal(valueOf(cell.getValue()).stripTrailingZeros().toPlainString());
                    handleGeneralFormat(cellProps, xNumberFormats, locale, cellValue);
                }
            }
        } catch (Exception e) {
            log.error("Error processing cell", e);
        }
    }

    /**
     * if 整数或者小数，整数部分>=12位
     * then 则使用科学计数法，科学计数法中E前面的数字最多保留五位小数，如果五位小数都是0，则忽略0
     * <p>
     * else if 绝对值小于0.0001(10的负四次方)，且总位数（不算末尾的额外的0）>=11
     * then 则使用科学计数法，科学计数法中E前面的数字最多保留五位小数，如果五位小数都是0，则忽略0
     * <p>
     * else if 有小数点的，并且总位数>=11，则四舍五入
     * then 如果四舍五入后最后一位是0，则忽略0
     */
    private void handleGeneralFormat(XPropertySet cellProps, XNumberFormats xNumberFormats, Locale locale, BigDecimal cellValue)
            throws PropertyVetoException, WrappedTargetException, MalformedNumberFormatException, UnknownPropertyException {
        boolean isInteger = isInteger(cellValue);
        int totalDigits = getTotalDigits(cellValue);
        int digitsBeforeDecimal = getDigitsBeforeDecimal(cellValue);
        int digitsAfterDecimal = getDigitsAfterDecimal(cellValue);

        if (digitsBeforeDecimal >= 12) {
            applyScientificNotationFormat(cellProps, xNumberFormats, locale, cellValue);
        } else if (cellValue.abs().compareTo(new BigDecimal("0.0001")) < 0 && totalDigits >= 11) {
            applyScientificNotationFormat(cellProps, xNumberFormats, locale, cellValue);
        } else if (!isInteger && totalDigits >= 11) {
            applyDecimalFormat(cellProps, xNumberFormats, locale, cellValue, digitsAfterDecimal, totalDigits);
        } else {
            log.info("Not going to change format for: value={}, isInteger={}, total digits={}, digitsBeforeDecimal={}, digitsAfterDecimal={}",
                    cellValue, isInteger, totalDigits, digitsBeforeDecimal, digitsAfterDecimal);
        }
    }

    private void applyScientificNotationFormat(XPropertySet cellProps, XNumberFormats xNumberFormats, Locale locale,
                                               BigDecimal cellValue)
            throws PropertyVetoException, WrappedTargetException, UnknownPropertyException, MalformedNumberFormatException {
        String newFormat = "0.#####E+00";
        int newFormatID = addOrQueryFormat(xNumberFormats, newFormat, locale);
        changeNumberFormat(cellProps, newFormatID);
        log.info("Integer value with total digits >= 12. Changed format to {} for value={}", newFormat, cellValue);
    }

    private void applyDecimalFormat(XPropertySet cellProps, XNumberFormats xNumberFormats, Locale locale,
                                    BigDecimal cellValue, int digitsAfterDecimal, int totalDigits)
            throws MalformedNumberFormatException, PropertyVetoException, WrappedTargetException, UnknownPropertyException {
        int zerosAfterDecimal = digitsAfterDecimal - (totalDigits - 10);
        String newFormat = "0." + "#".repeat(Math.max(0, zerosAfterDecimal));
        int newFormatID = addOrQueryFormat(xNumberFormats, newFormat, locale);
        changeNumberFormat(cellProps, newFormatID);
        log.info("Decimal value with total digits >= 11. Changed format to {} for value={}", newFormat, cellValue);
    }

    private int addOrQueryFormat(XNumberFormats xNumberFormats, String format, Locale locale)
            throws MalformedNumberFormatException {
        int formatID = xNumberFormats.queryKey(format, locale, false);
        if (formatID == -1) {
            formatID = xNumberFormats.addNew(format, locale);
        }
        return formatID;
    }

    private void changeNumberFormat(XPropertySet cellProps, int newNumFormat) throws
            WrappedTargetException, UnknownPropertyException, PropertyVetoException {
        log.info("before set:{}, new id: {}", cellProps.getPropertyValue("NumberFormat"), newNumFormat);
        cellProps.setPropertyValue("NumberFormat", Integer.valueOf(newNumFormat));
        log.info("after set format id:{}", cellProps.getPropertyValue("NumberFormat"));
    }

    private CellRangeAddress getCellRangeAddress(XUsedAreaCursor xUsedAreaCursor) {
        XCellRangeAddressable rangeAddressable = queryInterface(XCellRangeAddressable.class, xUsedAreaCursor);
        return rangeAddressable.getRangeAddress();
    }

    private boolean isSheetVisible(XSpreadsheet sheet)
            throws com.sun.star.uno.Exception {
        XPropertySet xSheetProps = queryInterface(XPropertySet.class, sheet);
        return (boolean) xSheetProps.getPropertyValue("IsVisible");
    }

    public static boolean isInteger(BigDecimal bd) {
        return bd.scale() == 0;
    }

    public static int getTotalDigits(BigDecimal bd) {
        String plainString = bd.stripTrailingZeros().toPlainString();
        int totalDigits = 0;
        for (char c : plainString.toCharArray()) {
            if (Character.isDigit(c)) {
                totalDigits++;
            }
        }
        return totalDigits;
    }

    public static int getDigitsBeforeDecimal(BigDecimal bd) {
        return getTotalDigits(bd) - bd.scale();
    }

    public static int getDigitsAfterDecimal(BigDecimal bd) {
        return Math.max(bd.scale(), 0);
    }
}
