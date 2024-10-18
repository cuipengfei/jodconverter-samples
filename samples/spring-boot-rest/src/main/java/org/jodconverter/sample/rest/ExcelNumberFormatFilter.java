package org.jodconverter.sample.rest;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.Locale;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.sheet.*;
import com.sun.star.table.CellRangeAddress;
import com.sun.star.table.XCell;
import com.sun.star.uno.AnyConverter;
import com.sun.star.util.XNumberFormats;
import com.sun.star.util.XNumberFormatsSupplier;
import org.jodconverter.core.office.OfficeContext;
import org.jodconverter.local.filter.Filter;
import org.jodconverter.local.filter.FilterChain;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.Date;

import static com.sun.star.table.CellContentType.VALUE;
import static com.sun.star.uno.UnoRuntime.queryInterface;
import static java.util.stream.Collectors.joining;

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
//
//        // Save the processed result into a file
//        XStorable xStorable = queryInterface(XStorable.class, document);
//        if (xStorable != null) {
//            String outputPath = "file:///" + new File("output_" + Math.abs(new Date().hashCode()) + ".ods").getAbsolutePath().replace("\\", "/");
//            xStorable.storeAsURL(outputPath, new com.sun.star.beans.PropertyValue[0]);
//        }

        chain.doFilter(context, document);
    }

    private void processSheet(XSpreadsheet sheet, XNumberFormats xNumberFormats) throws Exception {
        XSheetCellCursor cursor = sheet.createCursor();
        XUsedAreaCursor usedAreaCursor = queryInterface(XUsedAreaCursor.class, cursor);
        usedAreaCursor.gotoEndOfUsedArea(true);

        CellRangeAddress rangeAddress = getCellRangeAddress(usedAreaCursor);

        int endColumn = rangeAddress.EndColumn;
        int endRow = rangeAddress.EndRow;

        for (int col = 0; col <= endColumn; col++) {
            for (int row = 0; row <= endRow; row++) {
                XCell cell = sheet.getCellByPosition(col, row);
                processCell(cell, xNumberFormats);
            }
        }
    }

    private void processCell(XCell cell, XNumberFormats xNumberFormats) {
        try {
            if (cell.getType() == VALUE) {
                XPropertySet cellProps = queryInterface(XPropertySet.class, cell);
                int formatID = AnyConverter.toInt(cellProps.getPropertyValue("NumberFormat"));
                XPropertySet numberFormat = xNumberFormats.getByKey(formatID);
                Locale locale = (Locale) numberFormat.getPropertyValue("Locale");
                String formatString = numberFormat.getPropertyValue("FormatString").toString();

                if (formatString.equals("General")) {
                    double value = cell.getValue();
                    BigDecimal cellValue = BigDecimal.valueOf(value);
                    boolean isInteger = isInteger(cellValue);
                    int totalDigits = getTotalDigits(cellValue);
                    int digitsBeforeDecimal = getDigitsBeforeDecimal(cellValue);
                    int digitsAfterDecimal = getDigitsAfterDecimal(cellValue);

                    if (isInteger && totalDigits >= 12) {
                        String newFormat = "0.00000E+00";
                        int newFormatID = xNumberFormats.queryKey(newFormat, locale, false);
                        if (newFormatID == -1) {
                            newFormatID = xNumberFormats.addNew(newFormat, locale);
                        }
                        changeNumberFormat(cell, cellProps, newFormatID, value);
                        log.info("Integer value with total digits >= 12. Changed format to {} for value={}", newFormat, cellValue);
                    } else if (!isInteger && totalDigits >= 11) {
                        int zerosAfterDecimal = digitsAfterDecimal - (totalDigits - 10);
                        StringBuilder formatBuilder = new StringBuilder("0.");
                        for (int i = 0; i < zerosAfterDecimal; i++) {
                            formatBuilder.append('0');
                        }
                        String newFormat = formatBuilder.toString();
                        int newFormatID = xNumberFormats.queryKey(newFormat, locale, false);
                        if (newFormatID == -1) {
                            newFormatID = xNumberFormats.addNew(newFormat, locale);
                        }
                        changeNumberFormat(cell, cellProps, newFormatID, value);
                        log.info("Decimal value with total digits >= 11. Changed format to {} for value={}", newFormat, cellValue);
                    } else {
                        log.info("Not going to change format for: value={}, isInteger={}, total digits={}, digitsBeforeDecimal={}, digitsAfterDecimal={}",
                                cellValue, isInteger, totalDigits, digitsBeforeDecimal, digitsAfterDecimal);
                    }
                }
            }
        } catch (Exception e) {
            log.error("Error processing cell", e);
        }
    }

    private static void changeNumberFormat(XCell cell, XPropertySet cellProps, int newValue, double value)
            throws UnknownPropertyException, PropertyVetoException, WrappedTargetException {
        log.info("before set:{}, new id: {}", cellProps.getPropertyValue("NumberFormat"), newValue);

        cellProps.setPropertyValue("NumberFormat", newValue);
        log.info("after set:{}", cellProps.getPropertyValue("NumberFormat"));

        cellProps.setPropertyValue("NumberFormat", Integer.valueOf(newValue));
        log.info("after set Integer:{}", cellProps.getPropertyValue("NumberFormat"));

        cellProps.setPropertyValue("NumberFormat", AnyConverter.toInt(newValue));
        log.info("after set AnyConverter:{}", cellProps.getPropertyValue("NumberFormat"));

        // try to refresh, but seem not helping
        cell.setFormula(cell.getFormula());  // Refresh the cell
        log.info("after formula:{}", cellProps.getPropertyValue("NumberFormat"));

        cell.setValue(Double.NaN);
        log.info("after NaN:{}", cellProps.getPropertyValue("NumberFormat"));

        cell.setValue(value);
        log.info("after value:{}", cellProps.getPropertyValue("NumberFormat"));
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

    private String printProps(XPropertySet xPageStyleProps) {
        String info = Arrays.stream(xPageStyleProps.getPropertySetInfo().getProperties())
                .filter(x -> true)
                .map(x -> {
                    try {
                        return x.Name + " is " + xPageStyleProps.getPropertyValue(x.Name);
                    } catch (UnknownPropertyException | WrappedTargetException e) {
                        return "failed";
                    }
                }).collect(joining("\n"));
        log.info(info);
        return info;
    }
}