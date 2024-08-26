package org.jodconverter.sample.rest;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameAccess;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPageSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.XComponent;
import com.sun.star.sheet.XCellRangeAddressable;
import com.sun.star.sheet.XHeaderFooterContent;
import com.sun.star.sheet.XPrintAreas;
import com.sun.star.sheet.XSheetCellCursor;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XUsedAreaCursor;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.table.CellRangeAddress;
import com.sun.star.table.XCellRange;
import com.sun.star.table.XColumnRowRange;
import com.sun.star.table.XTableColumns;
import com.sun.star.table.XTableRows;
import org.jodconverter.core.office.OfficeContext;
import org.jodconverter.local.filter.Filter;
import org.jodconverter.local.filter.FilterChain;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Arrays;
import java.util.Objects;
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

    private boolean isSheetVisible(XSpreadsheet sheet)
            throws com.sun.star.uno.Exception {
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
        triggerReLayout(rows, columns);

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

        totalHeight += 4000;
        totalWidth += 2000;

        setPaperSizeAndPosition(xPageStyleProps, totalWidth, totalHeight);
    }

    /**
     * The purpose of these operations is to trigger a re-layout of the spreadsheet.
     * By temporarily inserting and then removing a row, the method forces the spreadsheet to
     * recalculate the positions of all elements, including images.
     * This recalculation helps to ensure that images are properly positioned and do not overlap with the text.
     **/
    private static void triggerReLayout(XTableRows rows, XTableColumns columns) {
        columns.insertByIndex(0, 1);
        columns.removeByIndex(0, 1);
        rows.insertByIndex(0, 1);
        rows.removeByIndex(0, 1);
    }

    private int getColumnWidth(int columnIndex, XTableColumns columns)
            throws com.sun.star.uno.Exception {
        Object column = columns.getByIndex(columnIndex);
        XPropertySet columnProps = queryInterface(XPropertySet.class, column);
        return (int) columnProps.getPropertyValue("Width");
    }

    private int getRowHeight(int rowIndex, XTableRows rows)
            throws com.sun.star.uno.Exception {
        Object row = rows.getByIndex(rowIndex);
        XPropertySet rowProps = queryInterface(XPropertySet.class, row);
        return (int) rowProps.getPropertyValue("Height");
    }

    private XUsedAreaCursor goToEnd(XSpreadsheet sheet) {
        XSheetCellCursor xSheetCellCursor = sheet.createCursor();
        XUsedAreaCursor xUsedAreaCursor = queryInterface(XUsedAreaCursor.class, xSheetCellCursor);
        xUsedAreaCursor.gotoEndOfUsedArea(true); // 定位到使用过的区域
        return xUsedAreaCursor;
    }

    private CellRangeAddress getCellRangeAddress(XUsedAreaCursor xUsedAreaCursor) {
        XCellRangeAddressable rangeAddressable = queryInterface(XCellRangeAddressable.class, xUsedAreaCursor);
        return rangeAddressable.getRangeAddress();
    }

    private XColumnRowRange getxColumnRowRange(XSpreadsheet sheet) {
        XCellRange cellRange = queryInterface(XCellRange.class, sheet);
        return queryInterface(XColumnRowRange.class, cellRange);
    }

    private int getTotalWidth(int endColumn, XTableColumns columns)
            throws com.sun.star.uno.Exception {
        int totalWidth = 0;
        for (int j = 0; j <= endColumn; j++) {
            Object column = columns.getByIndex(j);
            XPropertySet columnProps = queryInterface(XPropertySet.class, column);
            totalWidth += (int) columnProps.getPropertyValue("Width");
        }
        return totalWidth;
    }

    private int getTotalHeight(int endRow, XTableRows rows)
            throws com.sun.star.uno.Exception {
        int totalHeight = 0;
        for (int i = 0; i <= endRow; i++) {
            Object row = rows.getByIndex(i);
            XPropertySet rowProps = queryInterface(XPropertySet.class, row);
            totalHeight += (int) rowProps.getPropertyValue("Height");
        }
        return totalHeight;
    }

    private void clearPrintArea(XSpreadsheet sheet) {
        // If none of the sheets in a document have print areas, the whole sheets are printed.
        // If any sheet contains print areas, other sheets without print areas are not printed.
        XPrintAreas xPrintAreas = queryInterface(XPrintAreas.class, sheet);
        if (xPrintAreas != null) {
            xPrintAreas.setPrintAreas(new CellRangeAddress[]{});
        }
    }

    private XPropertySet getPageStyleProps(XSpreadsheet sheet, XNameAccess xPageStyles)
            throws com.sun.star.uno.Exception {
        String pageStyleName = queryInterface(XPropertySet.class, sheet).getPropertyValue("PageStyle").toString();
        log.info("page style name is: {}", pageStyleName);
        return queryInterface(XPropertySet.class, xPageStyles.getByName(pageStyleName));
    }

    private XNameAccess getPageStyles(XSpreadsheetDocument xSpreadsheetDocument)
            throws com.sun.star.uno.Exception {
        XStyleFamiliesSupplier xStyleFamiliesSupplier = queryInterface(XStyleFamiliesSupplier.class, xSpreadsheetDocument);
        XNameAccess xStyleFamilies = xStyleFamiliesSupplier.getStyleFamilies();
        return queryInterface(XNameAccess.class, xStyleFamilies.getByName("PageStyles"));
    }

    private void enableFooter(XPropertySet xPageStyleProps)
            throws com.sun.star.uno.Exception {
        xPageStyleProps.setPropertyValue("FooterShared", true);
        xPageStyleProps.setPropertyValue("FooterIsShared", true);
        xPageStyleProps.setPropertyValue("FirstPageFooterIsShared", true);

        xPageStyleProps.setPropertyValue("FooterIsOn", true);
        xPageStyleProps.setPropertyValue("FooterOn", true);
    }

    private void setFooterText(XPropertySet xPageStyleProps, String sheetName, String pageFooterContent)
            throws com.sun.star.uno.Exception {
        XHeaderFooterContent footerContent = queryInterface(XHeaderFooterContent.class, xPageStyleProps.getPropertyValue(pageFooterContent));
        if (footerContent != null) {
            log.info("Sheet {} {} has footer: {}, will change it to sheet name", sheetName, pageFooterContent, footerContent.getLeftText().getString());
            footerContent.getLeftText().setString(sheetName);
            xPageStyleProps.setPropertyValue(pageFooterContent, footerContent);
        }
    }

    private Size getGraphicalObjectsSize(XSpreadsheet sheet)
            throws com.sun.star.uno.Exception {
        XDrawPageSupplier drawPageSupplier = queryInterface(XDrawPageSupplier.class, sheet);
        XDrawPage drawPage = drawPageSupplier.getDrawPage();
        int count = drawPage.getCount();

        int maxWidth = 0;
        int maxHeight = 0;

        for (int i = 0; i < count; i++) {
            XShape shape = queryInterface(XShape.class, drawPage.getByIndex(i));
            addGlowForTinyImage(shape);

            Point position = shape.getPosition();
            Size size = shape.getSize();

            maxWidth = Math.max(maxWidth, position.X + size.Width);
            maxHeight = Math.max(maxHeight, position.Y + size.Height);
        }

        return new Size(maxWidth, maxHeight);
    }

    /**
     * Add glow effect for tiny image. The image is considered tiny if its width or height is less than 5mm.
     * Adding glow will make it visible when rendered by pdf.js, otherwise you can not see it in pdf.js.
     */
    private void addGlowForTinyImage(XShape shape)
            throws com.sun.star.uno.Exception {
        if (Objects.equals(shape.getShapeType(), "com.sun.star.drawing.GraphicObjectShape")) {
            Size size = shape.getSize();
            boolean isSmallerThan5mm = size.Width <= 500 || size.Height <= 500;
            if (isSmallerThan5mm) {
                XPropertySet shapeProps = queryInterface(XPropertySet.class, shape);
                Object radius = shapeProps.getPropertyValue("GlowEffectRadius");
                if ((int) radius == 0) {
                    shapeProps.setPropertyValue("GlowEffectRadius", 1);
                    shapeProps.setPropertyValue("GlowEffectColor", 4485828);
                }
            }
        }
    }

    private void setPaperSizeAndPosition(XPropertySet xPageStyleProps, int totalWidth, int totalHeight)
            throws com.sun.star.uno.Exception {
        xPageStyleProps.setPropertyValue("CenterVertically", true);
        xPageStyleProps.setPropertyValue("CenterHorizontally", true);

        xPageStyleProps.setPropertyValue("TopMargin", 1000);
        xPageStyleProps.setPropertyValue("HeaderBodyDistance", 0);
        xPageStyleProps.setPropertyValue("HeaderHeight", 0);

        xPageStyleProps.setPropertyValue("BottomMargin", 1000);
        xPageStyleProps.setPropertyValue("FooterBodyDistance", 0);
        xPageStyleProps.setPropertyValue("FooterHeight", 0);

        xPageStyleProps.setPropertyValue("LeftMargin", 0);
        xPageStyleProps.setPropertyValue("RightMargin", 0);

        xPageStyleProps.setPropertyValue("Size", new Size(totalWidth, totalHeight));

        // 设置缩放比例以适应一页
        // must be short
        xPageStyleProps.setPropertyValue("ScaleToPages", (short) 1);
    }

    //
//    private int minMargin() {
//        return 1000;
//    }
//
//    private int minFooterHeader() {
//        return 3000;
//    }
//
//    private void printProps(XPropertySet xPageStyleProps) {
//        String info = Arrays.stream(xPageStyleProps.getPropertySetInfo().getProperties())
//                .filter(x -> true)
//                .map(x -> {
//                    try {
//                        return x.Name + " is " + xPageStyleProps.getPropertyValue(x.Name);
//                    } catch (UnknownPropertyException | WrappedTargetException e) {
//                        return "failed";
//                    }
//                }).collect(joining("\n"));
//        log.info(info);
//    }
//
//    private int calculateFooterHeaderHeight(XPropertySet xPageStyleProps)
//            throws UnknownPropertyException, PropertyVetoException, WrappedTargetException {
//        int footerHeight = Math.max((int) xPageStyleProps.getPropertyValue("FooterHeight"), minFooterHeader());
//        int bottomMargin = Math.max((int) xPageStyleProps.getPropertyValue("BottomMargin"), minMargin());
//
//        if (isHeaderEnabled(xPageStyleProps)) {
//            int headerHeight = Math.max((int) xPageStyleProps.getPropertyValue("HeaderHeight"), minFooterHeader());
//            int topMargin = Math.max((int) xPageStyleProps.getPropertyValue("TopMargin"), minMargin());
//            return footerHeight + bottomMargin + headerHeight + topMargin;
//        } else {
//            return footerHeight + bottomMargin;
//        }
//    }
//
//    private int calculateMargins(XPropertySet xPageStyleProps)
//            throws UnknownPropertyException, PropertyVetoException, WrappedTargetException {
//        int leftMargin = Math.max((int) xPageStyleProps.getPropertyValue("LeftMargin"), minMargin());
//        int rightMargin = Math.max((int) xPageStyleProps.getPropertyValue("RightMargin"), minMargin());
//        return leftMargin + rightMargin;
//    }

//    private boolean isHeaderEnabled(XPropertySet xPageStyleProps)
//            throws com.sun.star.uno.Exception {
//        return (boolean) xPageStyleProps.getPropertyValue("HeaderIsOn") || (boolean) xPageStyleProps.getPropertyValue("HeaderOn");
//    }
}
