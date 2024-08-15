package org.jodconverter.sample.rest;

import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.text.*;
import org.jodconverter.core.office.OfficeContext;
import org.jodconverter.local.filter.Filter;
import org.jodconverter.local.filter.FilterChain;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import static com.sun.star.uno.UnoRuntime.queryInterface;

public class WordFrameFilter implements Filter {
    private static final Logger log = LoggerFactory.getLogger(WordFrameFilter.class);

    @Override
    public void doFilter(OfficeContext context, XComponent document, FilterChain chain) throws Exception {
        // Check if it is a word document
        XTextDocument xTextDocument = queryInterface(XTextDocument.class, document);
        if (xTextDocument == null) {
            // Not a word document, pass it down the filter chain.
            chain.doFilter(context, document);
            return;
        }

        overrideFrameColorWithParagraphColor(xTextDocument);

        // Continue with the filter chain
        chain.doFilter(context, document);
    }

    public void overrideFrameColorWithParagraphColor(XTextDocument xTextDocument) throws Exception {
        // 获取文档的文本内容
        XText xText = xTextDocument.getText();

        // 获取 XTextFramesSupplier 接口
        XTextFramesSupplier xTextFramesSupplier = queryInterface(XTextFramesSupplier.class, xTextDocument);

        // 获取所有的框架
        XNameAccess xNameAccess = xTextFramesSupplier.getTextFrames();
        String[] frameNames = xNameAccess.getElementNames();

        for (String frameName : frameNames) {
            try {
                // 获取每个frame
                XTextFrame xTextFrame = queryInterface(XTextFrame.class, xNameAccess.getByName(frameName));
                XPropertySet xProps = queryInterface(XPropertySet.class, xTextFrame);
                Object firstParagraphBackColor = getFirstParagraphBackColor(xTextFrame.getText());

                if ((int) xProps.getPropertyValue("BackColor") < 0) {
                    xProps.setPropertyValue("BackColor", firstParagraphBackColor);
                }
                if ((int) xProps.getPropertyValue("BackColorRGB") < 0) {
                    xProps.setPropertyValue("BackColorRGB", firstParagraphBackColor);
                }
                if ((Byte) xProps.getPropertyValue("BackColorTransparency") == 100) {
                    xProps.setPropertyValue("BackColorTransparency", 0);
                }
            } catch (Exception e) {
                log.error("error handing frame", e);
            }
        }
    }

    private static int getFirstParagraphBackColor(XText xText)
            throws NoSuchElementException, WrappedTargetException, UnknownPropertyException {
        XEnumerationAccess xEnumerationAccess = queryInterface(XEnumerationAccess.class, xText);
        XEnumeration enumeration = xEnumerationAccess.createEnumeration();

        while (enumeration.hasMoreElements()) {
            XTextContent xTextContent = queryInterface(XTextContent.class, enumeration.nextElement());
            if (xTextContent != null) {
                log.info("found paragraph inside: {}", xTextContent.getAnchor().getString());

                XPropertySet xParagraphProperties = queryInterface(XPropertySet.class, xTextContent);
                Object paraBackColor = xParagraphProperties.getPropertyValue("ParaBackColor");
                log.info("this paragraph's back color is {}", paraBackColor);

                return (int) paraBackColor;
            }
        }

        log.info("no paragraph found, will return null");
        return 0;
    }
}
