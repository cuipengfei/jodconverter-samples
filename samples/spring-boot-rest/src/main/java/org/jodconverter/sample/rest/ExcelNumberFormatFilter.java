package org.jodconverter.sample.rest;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.lang.Locale;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.sheet.*;
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

import static com.sun.star.table.CellContentType.VALUE;
import static com.sun.star.uno.UnoRuntime.queryInterface;

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
            if (cell.getType() == VALUE) {
                XPropertySet cellProps = queryInterface(XPropertySet.class, cell);
                int key = AnyConverter.toInt(cellProps.getPropertyValue("NumberFormat"));
                XPropertySet numberFormat = xNumberFormats.getByKey(key);
                Locale locale = (Locale) numberFormat.getPropertyValue("Locale");
                String formatString = numberFormat.getPropertyValue("FormatString").toString();

                if (formatString.equals("General")) {
                    BigDecimal cellValue = BigDecimal.valueOf(cell.getValue());
                    handleGeneralFormat(cellProps, xNumberFormats, locale, cellValue);
                }
            }
        } catch (Exception e) {
            log.error("Error processing cell", e);
        }
    }

    private void handleGeneralFormat(XPropertySet cellProps, XNumberFormats xNumberFormats, Locale locale, BigDecimal cellValue)
            throws PropertyVetoException, WrappedTargetException, MalformedNumberFormatException, UnknownPropertyException {
        boolean isInteger = isInteger(cellValue);
        int totalDigits = getTotalDigits(cellValue);
        int digitsBeforeDecimal = getDigitsBeforeDecimal(cellValue);
        int digitsAfterDecimal = getDigitsAfterDecimal(cellValue);

        if (isInteger && totalDigits >= 12) {
            applyScientificNotationFormat(cellProps, xNumberFormats, locale, cellValue);
        } else if (!isInteger && totalDigits >= 11) {
            applyDecimalFormat(cellProps, xNumberFormats, locale, cellValue, digitsAfterDecimal);
        } else {
            log.info("Not going to change format for: value={}, isInteger={}, total digits={}, digitsBeforeDecimal={}, digitsAfterDecimal={}",
                    cellValue, isInteger, totalDigits, digitsBeforeDecimal, digitsAfterDecimal);
        }
    }

    private void applyScientificNotationFormat(XPropertySet cellProps, XNumberFormats xNumberFormats, Locale locale, BigDecimal cellValue)
            throws PropertyVetoException, WrappedTargetException, UnknownPropertyException, MalformedNumberFormatException {
        String newFormat = "0.00000E+00";
        int newFormatID = addOrQueryFormat(xNumberFormats, newFormat, locale);
        changeNumberFormat(cellProps, newFormatID);
        log.info("Integer value with total digits >= 12. Changed format to {} for value={}", newFormat, cellValue);
    }

    private void applyDecimalFormat(XPropertySet cellProps, XNumberFormats xNumberFormats, Locale locale, BigDecimal cellValue, int digitsAfterDecimal)
            throws MalformedNumberFormatException, PropertyVetoException, WrappedTargetException, UnknownPropertyException {
        int zerosAfterDecimal = digitsAfterDecimal - (getTotalDigits(cellValue) - 10);
        String newFormat = "0." + "0".repeat(Math.max(0, zerosAfterDecimal));
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
        return bd.precision();
    }

    public static int getDigitsBeforeDecimal(BigDecimal bd) {
        return bd.precision() - bd.scale();
    }

    public static int getDigitsAfterDecimal(BigDecimal bd) {
        return Math.max(bd.scale(), 0);
    }
}
