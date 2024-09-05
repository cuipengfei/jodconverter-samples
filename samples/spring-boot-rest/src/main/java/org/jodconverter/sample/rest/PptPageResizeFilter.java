package org.jodconverter.sample.rest;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.uno.UnoRuntime;
import org.jodconverter.core.office.OfficeContext;
import org.jodconverter.local.filter.Filter;
import org.jodconverter.local.filter.FilterChain;
import org.jodconverter.local.office.utils.Draw;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class PptPageResizeFilter implements Filter {
    private static final Logger log = LoggerFactory.getLogger(PptPageResizeFilter.class);

    @Override
    public void doFilter(OfficeContext context, XComponent document, FilterChain chain) throws Exception {
        boolean isImpress = Draw.isImpress(document);
        if (!isImpress) {
            chain.doFilter(context, document);
            return;
        }

        log.info("Adjusting PowerPoint document to fit images within slide bounds.");
        XDrawPagesSupplier xDrawPagesSupplier = UnoRuntime.queryInterface(XDrawPagesSupplier.class, document);

        int maxWidth = 0;
        int maxHeight = 0;
        Size originalSize = null;

        for (int i = 0; i < xDrawPagesSupplier.getDrawPages().getCount(); i++) {
            log.info("Processing slide {}", i + 1);
            XDrawPage drawPage = UnoRuntime.queryInterface(XDrawPage.class, xDrawPagesSupplier.getDrawPages().getByIndex(i));
            Size sizeIncludingImages = calculateSlideSizeIncludingImages(drawPage);
            log.info("Calculated size including images for slide {}: width={}, height={}", i + 1, sizeIncludingImages.Width, sizeIncludingImages.Height);

            if (originalSize == null) {
                originalSize = getOriginalSize(drawPage);
                log.info("Original size of slide {}: width={}, height={}", i + 1, originalSize.Width, originalSize.Height);
            }
            maxWidth = Math.max(maxWidth, sizeIncludingImages.Width);
            maxHeight = Math.max(maxHeight, sizeIncludingImages.Height);
        }

        if (originalSize != null && (maxWidth > originalSize.Width || maxHeight > originalSize.Height)) {
            int newWidth = Math.min((int) (originalSize.Width * 1.2), maxWidth);
            int newHeight = Math.min((int) (originalSize.Height * 1.2), maxHeight);
            log.info("Resizing slides to new dimensions: width={}, height={}", newWidth, newHeight);

            for (int i = 0; i < xDrawPagesSupplier.getDrawPages().getCount(); i++) {
                XDrawPage xDrawPage = UnoRuntime.queryInterface(XDrawPage.class, xDrawPagesSupplier.getDrawPages().getByIndex(i));
                XPropertySet slideProps = UnoRuntime.queryInterface(XPropertySet.class, xDrawPage);
                slideProps.setPropertyValue("Width", newWidth);
                slideProps.setPropertyValue("Height", newHeight);
                log.info("Slide {} resized to width={}, height={}", i + 1, newWidth, newHeight);
            }
        }

        log.info("Finished adjusting PowerPoint document.");
        chain.doFilter(context, document);
        log.info("Finished PptPageResizeFilter.doFilter");
    }

    private static Size getOriginalSize(XDrawPage drawPage) throws UnknownPropertyException, WrappedTargetException {
        XPropertySet slideProps = UnoRuntime.queryInterface(XPropertySet.class, drawPage);
        int w = (int) slideProps.getPropertyValue("Width");
        int h = (int) slideProps.getPropertyValue("Height");
        return new Size(w, h);
    }

    private Size calculateSlideSizeIncludingImages(XDrawPage drawPage) throws Exception {
        int maxWidth = 0;
        int maxHeight = 0;
        int shapeCount = drawPage.getCount();

        for (int i = 0; i < shapeCount; i++) {
            XShape shape = UnoRuntime.queryInterface(XShape.class, drawPage.getByIndex(i));
            Point position = shape.getPosition();
            Size size = shape.getSize();

            maxWidth = Math.max(maxWidth, position.X + size.Width);
            maxHeight = Math.max(maxHeight, position.Y + size.Height);
        }

        return new Size(maxWidth, maxHeight);
    }
}
