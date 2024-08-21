package org.jodconverter.sample.rest;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPageSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.sheet.*;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.table.*;
import org.jodconverter.core.office.OfficeContext;
import org.jodconverter.local.filter.Filter;
import org.jodconverter.local.filter.FilterChain;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Arrays;
import java.util.concurrent.CompletableFuture;

import static com.sun.star.uno.UnoRuntime.queryInterface;

public class ExcelSinglePageFilter implements Filter {
    private static final Logger log = LoggerFactory.getLogger(ExcelSinglePageFilter.class);

    @Override
    public void doFilter(OfficeContext context, XComponent document, FilterChain chain) throws Exception {
        XSpreadsheetDocument xSpreadsheetDocument = queryInterface(XSpreadsheetDocument.class, document);
        if (xSpreadsheetDocument == null) {
            chain.doFilter(context, document);
            return;
        }

        log.info("Adjusting Excel document to fit on single page.");
        XNameAccess xPageStyles = getPageStyles(xSpreadsheetDocument);

        String[] sheetNames = xSpreadsheetDocument.getSheets().getElementNames();
        CompletableFuture[] futures = Arrays.stream(sheetNames).map(sheetName -> CompletableFuture.runAsync(() -> {
            try {
                log.info("Processing sheet: {}", sheetName);

                XSpreadsheet sheet = queryInterface(XSpreadsheet.class, xSpreadsheetDocument.getSheets().getByName(sheetName));

                if (isSheetVisible(sheet)) {
                    adjustSheetForSinglePage(sheetName, sheet, xPageStyles);
                } else {
                    log.info("Clearing print area of hidden sheet: {}", sheetName);
                    clearPrintArea(sheet);
                    log.info("Skipping other processing of hidden sheet: {}", sheetName);
                }
            } catch (Exception e) {
                log.error("Error processing sheet: {}", sheetName, e);
            }
        })).toList().toArray(new CompletableFuture[0]);

        CompletableFuture.allOf(futures).join();

        chain.doFilter(context, document);
    }

    private static boolean isSheetVisible(XSpreadsheet sheet) throws UnknownPropertyException, WrappedTargetException {
        XPropertySet xSheetProps = queryInterface(XPropertySet.class, sheet);
        return (boolean) xSheetProps.getPropertyValue("IsVisible");
    }

    private void adjustSheetForSinglePage(String sheetName, XSpreadsheet sheet, XNameAccess xPageStyles)
            throws Exception {
        XUsedAreaCursor xUsedAreaCursor = goToEnd(sheet);
        clearPrintArea(sheet);

        CellRangeAddress rangeAddress = getCellRangeAddress(xUsedAreaCursor);
        XColumnRowRange columnRowRange = getxColumnRowRange(sheet);
        XPropertySet xPageStyleProps = getPageStyleProps(sheet, xPageStyles);

        enableFooter(xPageStyleProps);
        setFooterText(xPageStyleProps, sheetName, "RightPageFooterContent");

        log.info("Sheet: {} used area column: {}, row: {}", sheetName, rangeAddress.EndColumn, rangeAddress.EndRow);

        XTableColumns columns = columnRowRange.getColumns();
        XTableRows rows = columnRowRange.getRows();

        int totalWidth = getTotalWidth(rangeAddress.EndColumn, columns);
        int totalHeight = getTotalHeight(rangeAddress.EndRow, rows);
        log.info("Sheet: {} used area total width: {}, total height: {}", sheetName, totalWidth, totalHeight);

        Size graphicalSize = getGraphicalObjectsSize(sheet);

        // Adjust totalWidth and totalHeight to accommodate graphical objects
        while (totalWidth < graphicalSize.Width || totalHeight < graphicalSize.Height) {
            if (totalWidth < graphicalSize.Width) {
                int newColumnWidth = getColumnWidth(rangeAddress.EndColumn + 1, columns);
                totalWidth += newColumnWidth;
                rangeAddress.EndColumn++;
            }
            if (totalHeight < graphicalSize.Height) {
                int newRowHeight = getRowHeight(rangeAddress.EndRow + 1, rows);
                totalHeight += newRowHeight;
                rangeAddress.EndRow++;
            }
        }
        log.info("Sheet: {} adjusted total width: {}, adjusted total height: {}", sheetName, totalWidth, totalHeight);

        totalHeight += calculateFooterHeaderHeight(xPageStyleProps);
        totalWidth += calculateMargins(xPageStyleProps);

        setPaperSizeAndPosition(xPageStyleProps, totalWidth, totalHeight);
    }

    private int getColumnWidth(int columnIndex, XTableColumns columns)
            throws com.sun.star.lang.IndexOutOfBoundsException, WrappedTargetException, UnknownPropertyException {
        Object column = columns.getByIndex(columnIndex);
        XPropertySet columnProps = queryInterface(XPropertySet.class, column);
        return (int) columnProps.getPropertyValue("Width");
    }

    private int getRowHeight(int rowIndex, XTableRows rows)
            throws com.sun.star.lang.IndexOutOfBoundsException, WrappedTargetException, UnknownPropertyException {
        Object row = rows.getByIndex(rowIndex);
        XPropertySet rowProps = queryInterface(XPropertySet.class, row);
        return (int) rowProps.getPropertyValue("Height");
    }

    private static XUsedAreaCursor goToEnd(XSpreadsheet sheet) {
        XSheetCellCursor xSheetCellCursor = sheet.createCursor();
        XUsedAreaCursor xUsedAreaCursor = queryInterface(XUsedAreaCursor.class, xSheetCellCursor);
        xUsedAreaCursor.gotoEndOfUsedArea(true); // 定位到使用过的区域
        return xUsedAreaCursor;
    }

    private static CellRangeAddress getCellRangeAddress(XUsedAreaCursor xUsedAreaCursor) {
        XCellRangeAddressable rangeAddressable = queryInterface(XCellRangeAddressable.class, xUsedAreaCursor);
        return rangeAddressable.getRangeAddress();
    }

    private static XColumnRowRange getxColumnRowRange(XSpreadsheet sheet) {
        XCellRange cellRange = queryInterface(XCellRange.class, sheet);
        return queryInterface(XColumnRowRange.class, cellRange);
    }

    private int getTotalWidth(int endColumn, XTableColumns columns)
            throws com.sun.star.lang.IndexOutOfBoundsException, WrappedTargetException, UnknownPropertyException {
        int totalWidth = 0;
        for (int j = 0; j <= endColumn; j++) {
            Object column = columns.getByIndex(j);
            XPropertySet columnProps = queryInterface(XPropertySet.class, column);
            totalWidth += (int) columnProps.getPropertyValue("Width");
        }
        return totalWidth;
    }

    private int getTotalHeight(int endRow, XTableRows rows)
            throws com.sun.star.lang.IndexOutOfBoundsException, WrappedTargetException, UnknownPropertyException {
        int totalHeight = 0;
        for (int i = 0; i <= endRow; i++) {
            Object row = rows.getByIndex(i);
            XPropertySet rowProps = queryInterface(XPropertySet.class, row);
            totalHeight += (int) rowProps.getPropertyValue("Height");
        }
        return totalHeight;
    }

    private static void clearPrintArea(XSpreadsheet sheet) {
        // If none of the sheets in a document have print areas, the whole sheets are printed.
        // If any sheet contains print areas, other sheets without print areas are not printed.
        XPrintAreas xPrintAreas = queryInterface(XPrintAreas.class, sheet);
        if (xPrintAreas != null) {
            xPrintAreas.setPrintAreas(new CellRangeAddress[]{});
        }
    }

    private static XPropertySet getPageStyleProps(XSpreadsheet sheet, XNameAccess xPageStyles)
            throws UnknownPropertyException, WrappedTargetException, NoSuchElementException {
        String pageStyleName = queryInterface(XPropertySet.class, sheet).getPropertyValue("PageStyle").toString();
        log.info("page style name is: {}", pageStyleName);
        return queryInterface(XPropertySet.class, xPageStyles.getByName(pageStyleName));
    }

    private static XNameAccess getPageStyles(XSpreadsheetDocument xSpreadsheetDocument)
            throws NoSuchElementException, WrappedTargetException {
        XStyleFamiliesSupplier xStyleFamiliesSupplier = queryInterface(XStyleFamiliesSupplier.class, xSpreadsheetDocument);
        XNameAccess xStyleFamilies = xStyleFamiliesSupplier.getStyleFamilies();
        return queryInterface(XNameAccess.class, xStyleFamilies.getByName("PageStyles"));
    }

    private void enableFooter(XPropertySet xPageStyleProps)
            throws UnknownPropertyException, PropertyVetoException, WrappedTargetException {
        xPageStyleProps.setPropertyValue("FooterShared", true);
        xPageStyleProps.setPropertyValue("FooterIsShared", true);
        xPageStyleProps.setPropertyValue("FirstPageFooterIsShared", true);

        xPageStyleProps.setPropertyValue("FooterIsOn", true);
        xPageStyleProps.setPropertyValue("FooterOn", true);
    }

    private void setFooterText(XPropertySet xPageStyleProps, String sheetName, String pageFooterContent)
            throws UnknownPropertyException, WrappedTargetException, PropertyVetoException {
        XHeaderFooterContent footerContent = queryInterface(XHeaderFooterContent.class, xPageStyleProps.getPropertyValue(pageFooterContent));
        if (footerContent != null) {
            log.info("Sheet {} {} has footer: {}, will change it to sheet name", sheetName, pageFooterContent, footerContent.getLeftText().getString());
            footerContent.getLeftText().setString(sheetName);
            xPageStyleProps.setPropertyValue(pageFooterContent, footerContent);
        }
    }

    private int calculateFooterHeaderHeight(XPropertySet xPageStyleProps)
            throws UnknownPropertyException, PropertyVetoException, WrappedTargetException {
        int footerHeight = Math.max((int) xPageStyleProps.getPropertyValue("FooterHeight"), minFooterHeader());
        int bottomMargin = Math.max((int) xPageStyleProps.getPropertyValue("BottomMargin"), minMargin());

        if (isHeaderEnabled(xPageStyleProps)) {
            int headerHeight = Math.max((int) xPageStyleProps.getPropertyValue("HeaderHeight"), minFooterHeader());
            int topMargin = Math.max((int) xPageStyleProps.getPropertyValue("TopMargin"), minMargin());
            return footerHeight + bottomMargin + headerHeight + topMargin;
        } else {
            return footerHeight + bottomMargin;
        }
    }

    private boolean isHeaderEnabled(XPropertySet xPageStyleProps)
            throws UnknownPropertyException, PropertyVetoException, WrappedTargetException {
        return (boolean) xPageStyleProps.getPropertyValue("HeaderIsOn") || (boolean) xPageStyleProps.getPropertyValue("HeaderOn");
    }

    private int calculateMargins(XPropertySet xPageStyleProps)
            throws UnknownPropertyException, PropertyVetoException, WrappedTargetException {
        int leftMargin = Math.max((int) xPageStyleProps.getPropertyValue("LeftMargin"), minMargin());
        int rightMargin = Math.max((int) xPageStyleProps.getPropertyValue("RightMargin"), minMargin());
        return leftMargin + rightMargin;
    }

    private Size getGraphicalObjectsSize(XSpreadsheet sheet)
            throws WrappedTargetException, IndexOutOfBoundsException {
        XDrawPageSupplier drawPageSupplier = queryInterface(XDrawPageSupplier.class, sheet);
        XDrawPage drawPage = drawPageSupplier.getDrawPage();
        int count = drawPage.getCount();

        int maxWidth = 0;
        int maxHeight = 0;

        for (int i = 0; i < count; i++) {
            XShape shape = queryInterface(XShape.class, drawPage.getByIndex(i));
            Point position = shape.getPosition();
            Size size = shape.getSize();

            maxWidth = Math.max(maxWidth, position.X + size.Width);
            maxHeight = Math.max(maxHeight, position.Y + size.Height);
        }

        return new Size(maxWidth, maxHeight);
    }

    private void setPaperSizeAndPosition(XPropertySet xPageStyleProps, int totalWidth, int totalHeight)
            throws UnknownPropertyException, PropertyVetoException, WrappedTargetException {
        xPageStyleProps.setPropertyValue("Size", new Size(totalWidth, totalHeight));
        xPageStyleProps.setPropertyValue("CenterVertically", true);
        xPageStyleProps.setPropertyValue("CenterHorizontally", true);
        xPageStyleProps.setPropertyValue("TopMargin", 0);
        xPageStyleProps.setPropertyValue("BottomMargin", 0);
        xPageStyleProps.setPropertyValue("RightMargin", 0);
        xPageStyleProps.setPropertyValue("LeftMargin", 0);
        // 设置缩放比例以适应一页
        // must be short
        xPageStyleProps.setPropertyValue("ScaleToPages", (short) 1);
    }

    private static int minMargin() {
        return 2100;
    }

    private static int minFooterHeader() {
        return 1200;
    }

//    private static void printHeaderFooterProps(XPropertySet xPageStyleProps) {
//        String info = Arrays.stream(xPageStyleProps.getPropertySetInfo().getProperties())
//                .filter(x -> {
//                    try {
//                        boolean isAboutHeaderFooter = x.Name.toLowerCase().contains("header") || x.Name.toLowerCase().contains("footer");
//                        boolean isBoolean = xPageStyleProps.getPropertyValue(x.Name) instanceof Boolean;
//                        return isBoolean;
//                    } catch (UnknownPropertyException | WrappedTargetException e) {
//                        return false;
//                    }
//                })
//                .map(x -> {
//                    try {
//                        return x.Name + " is " + xPageStyleProps.getPropertyValue(x.Name);
//                    } catch (UnknownPropertyException | WrappedTargetException e) {
//                        return "failed";
//                    }
//                }).collect(Collectors.joining("\n"));
//        log.info(info);
//    }
}
