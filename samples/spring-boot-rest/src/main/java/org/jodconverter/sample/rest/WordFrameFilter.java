package org.jodconverter.sample.rest;

import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.style.XStyle;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.text.*;
import org.jodconverter.core.office.OfficeContext;
import org.jodconverter.local.filter.Filter;
import org.jodconverter.local.filter.FilterChain;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Arrays;
import java.util.stream.Collectors;

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

        extractParagraphColors(xTextDocument);

        // Continue with the filter chain
        chain.doFilter(context, document);
    }

    public void extractParagraphColors(XTextDocument xTextDocument) throws Exception {
        // 获取文档的文本内容
        XText xText = xTextDocument.getText();

        XEnumerationAccess xEnumerationAccess = queryInterface(XEnumerationAccess.class, xText);
        XEnumeration enumeration = xEnumerationAccess.createEnumeration();

        while (enumeration.hasMoreElements()) {
            XTextContent xTextContent = queryInterface(XTextContent.class, enumeration.nextElement());
            log.info("paragraph content: " + xTextContent.getAnchor().getString());

            XPropertySet xParagraphProperties = queryInterface(XPropertySet.class, xTextContent);
            log.info("para back color: " + xParagraphProperties.getPropertyValue("ParaBackColor"));

//            logProps(xParagraphProperties.getPropertySetInfo(), xParagraphProperties, "xTextContent");
        }
    }

    private static void logProps(XPropertySetInfo propInfo, XPropertySet textProps, String type) {
        log.info("{} {}", type, Arrays.stream(propInfo.getProperties()).map(p -> {
            try {
//                log.info(p.Name);
//                log.info(p.Name + " is " + textProps.getPropertyValue(p.Name));
                return p.Name + " is " + textProps.getPropertyValue(p.Name);
            } catch (UnknownPropertyException e) {
                log.info("faild {}", p.Name);
            } catch (WrappedTargetException e) {
                log.info("faild {}", p.Name);
            }
            return p.Name;
        }).collect(Collectors.joining("\n")));
    }
}
