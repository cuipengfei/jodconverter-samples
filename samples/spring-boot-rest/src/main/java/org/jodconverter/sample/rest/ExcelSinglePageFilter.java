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
import com.sun.star.table.CellRangeAddress;
import com.sun.star.table.XCellRange;
import com.sun.star.table.XColumnRowRange;
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
        // 检查是否是Excel文档
        XSpreadsheetDocument xSpreadsheetDocument = queryInterface(XSpreadsheetDocument.class, document);
        if (xSpreadsheetDocument == null) {
            chain.doFilter(context, document);
            return;
        }

        log.info("going to make excel single page");
        // 获取全局xPageStyles
        XNameAccess xPageStyles = getPageStyles(xSpreadsheetDocument);

        // 遍历每个工作表
        String[] sheetNames = xSpreadsheetDocument.getSheets().getElementNames();
        CompletableFuture[] futures = Arrays.stream(sheetNames).map(sheetName -> CompletableFuture.runAsync(() -> {
            try {
                log.info("going to process sheet: {}", sheetName);

                // 获取当前工作表
                XSpreadsheet sheet = queryInterface(XSpreadsheet.class, xSpreadsheetDocument.getSheets().getByName(sheetName));

                // 跳过隐藏的工作表
                if (isSheetVisible(sheet)) {
                    adjustOneSheet(sheetName, sheet, xPageStyles);
                } else {
                    log.info("skipping hidden sheet: {}", sheetName);
                }
            } catch (Exception e) {
                log.error("Error processing sheet: {}", sheetName, e);
            }
        })).toList().toArray(new CompletableFuture[0]);

        // Wait for all futures to complete
        CompletableFuture.allOf(futures).join();

        chain.doFilter(context, document);
    }

    private static boolean isSheetVisible(XSpreadsheet sheet) throws UnknownPropertyException, WrappedTargetException {
        XPropertySet xSheetProps = queryInterface(XPropertySet.class, sheet);
        return (boolean) xSheetProps.getPropertyValue("IsVisible");
    }

    private static void adjustOneSheet(String sheetName, XSpreadsheet sheet, XNameAccess xPageStyles)
            throws com.sun.star.lang.IndexOutOfBoundsException, WrappedTargetException, UnknownPropertyException, NoSuchElementException, PropertyVetoException {
        XUsedAreaCursor xUsedAreaCursor = goToEnd(sheet);

        // 使用XCellRangeAddressable接口来获取范围地址
        CellRangeAddress rangeAddress = getCellRangeAddress(xUsedAreaCursor);
        // 获取列和行
        XColumnRowRange columnRowRange = getxColumnRowRange(sheet);

        log.info("sheet: {} used area column: {}, row: {}", sheetName, rangeAddress.EndColumn, rangeAddress.EndRow);
        // 计算非空列宽度
        int totalWidth = getTotalWidth(columnRowRange, rangeAddress.EndColumn);
        // 计算非空行高度
        int totalHeight = getTotalHeight(columnRowRange, rangeAddress.EndRow);
        log.info("sheet: {} used area total width: {}, total height: {}", sheetName, totalWidth, totalHeight);

        // Include graphical objects in the total dimensions
        Size graphicalSize = getGraphicalObjectsSize(sheet);
        totalWidth = Math.max(totalWidth, graphicalSize.Width);
        totalHeight = Math.max(totalHeight, graphicalSize.Height);
        log.info("sheet: {} final total width: {}, final total height: {}", sheetName, totalWidth, totalHeight);

        // 获取PageStyle
        XPropertySet xPageStyleProps = getPageStyleProps(sheet, xPageStyles);

        // 设置纸张大小和方向
        xPageStyleProps.setPropertyValue("IsLandscape", true); // 设置为横向打印
        xPageStyleProps.setPropertyValue("Size", new Size(totalWidth, totalHeight));
        setMarginToZero(xPageStyleProps);

        // 设置缩放比例以适应一页
        // must be short
        xPageStyleProps.setPropertyValue("ScaleToPages", (short) 1);
    }

    private static Size getGraphicalObjectsSize(XSpreadsheet sheet)
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

    private static void setMarginToZero(XPropertySet xPageStyleProps)
            throws UnknownPropertyException, PropertyVetoException, WrappedTargetException {
        xPageStyleProps.setPropertyValue("TopMargin", 0);
        xPageStyleProps.setPropertyValue("BottomMargin", 0);
        xPageStyleProps.setPropertyValue("RightMargin", 0);
        xPageStyleProps.setPropertyValue("LeftMargin", 0);
    }

    private static XPropertySet getPageStyleProps(XSpreadsheet sheet, XNameAccess xPageStyles)
            throws UnknownPropertyException, WrappedTargetException, NoSuchElementException {
        String pageStyleName = queryInterface(XPropertySet.class, sheet).getPropertyValue("PageStyle").toString();
        log.info("page style name is: {}", pageStyleName);
        return queryInterface(XPropertySet.class, xPageStyles.getByName(pageStyleName));
    }

    private static int getTotalHeight(XColumnRowRange columnRowRange, int endRow)
            throws com.sun.star.lang.IndexOutOfBoundsException, WrappedTargetException, UnknownPropertyException {
        int totalHeight = 0;
        for (int i = 0; i <= endRow; i++) {
            Object row = columnRowRange.getRows().getByIndex(i);
            XPropertySet rowProps = queryInterface(XPropertySet.class, row);
            totalHeight += (int) rowProps.getPropertyValue("Height");
        }
        return totalHeight;
    }

    private static int getTotalWidth(XColumnRowRange columnRowRange, int endColumn)
            throws com.sun.star.lang.IndexOutOfBoundsException, WrappedTargetException, UnknownPropertyException {
        int totalWidth = 0;
        for (int j = 0; j <= endColumn; j++) {
            Object column = columnRowRange.getColumns().getByIndex(j);
            XPropertySet columnProps = queryInterface(XPropertySet.class, column);
            totalWidth += (int) columnProps.getPropertyValue("Width");
        }
        return totalWidth;
    }

    private static XColumnRowRange getxColumnRowRange(XSpreadsheet sheet) {
        XCellRange cellRange = queryInterface(XCellRange.class, sheet);
        return queryInterface(XColumnRowRange.class, cellRange);
    }

    private static CellRangeAddress getCellRangeAddress(XUsedAreaCursor xUsedAreaCursor) {
        XCellRangeAddressable rangeAddressable = queryInterface(XCellRangeAddressable.class, xUsedAreaCursor);
        return rangeAddressable.getRangeAddress();
    }

    private static XUsedAreaCursor goToEnd(XSpreadsheet sheet) {
        XSheetCellCursor xSheetCellCursor = sheet.createCursor();
        XUsedAreaCursor xUsedAreaCursor = queryInterface(XUsedAreaCursor.class, xSheetCellCursor);
        xUsedAreaCursor.gotoEndOfUsedArea(true); // 定位到使用过的区域
        return xUsedAreaCursor;
    }

    private static XNameAccess getPageStyles(XSpreadsheetDocument xSpreadsheetDocument)
            throws NoSuchElementException, WrappedTargetException {
        XStyleFamiliesSupplier xStyleFamiliesSupplier = queryInterface(XStyleFamiliesSupplier.class, xSpreadsheetDocument);
        XNameAccess xStyleFamilies = xStyleFamiliesSupplier.getStyleFamilies();
        return queryInterface(XNameAccess.class, xStyleFamilies.getByName("PageStyles"));
    }
}
