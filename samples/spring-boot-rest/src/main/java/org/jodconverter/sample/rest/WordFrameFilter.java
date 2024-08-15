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

/**
 * 一个实现Filter接口的类，用于在JOD转换过程中调整Word文档中frame的背景色
 */
public class WordFrameFilter implements Filter {
    private static final Logger log = LoggerFactory.getLogger(WordFrameFilter.class);

    @Override
    public void doFilter(OfficeContext context, XComponent document, FilterChain chain) throws Exception {
        // 检查是否为Word文档
        XTextDocument xTextDocument = queryInterface(XTextDocument.class, document);
        if (xTextDocument == null) {
            // 非Word文档，传递给过滤链的下一个处理
            chain.doFilter(context, document);
            return;
        }

        // 调用方法用段落颜色覆盖frame颜色
        overrideFrameColorWithParagraphColor(xTextDocument);

        // 继续执行过滤链
        chain.doFilter(context, document);
    }

    /**
     * 用段落颜色覆盖frame的背景色
     */
    public void overrideFrameColorWithParagraphColor(XTextDocument xTextDocument) throws Exception {
        // 获取文档的文本内容
        XText xText = xTextDocument.getText();

        // 获取 XTextFramesSupplier 接口
        XTextFramesSupplier xTextFramesSupplier = queryInterface(XTextFramesSupplier.class, xTextDocument);

        // 获取所有的frame
        XNameAccess xNameAccess = xTextFramesSupplier.getTextFrames();
        String[] frameNames = xNameAccess.getElementNames();

        // 遍历所有frame名
        for (String frameName : frameNames) {
            try {
                // 获取每个frame
                XTextFrame xTextFrame = queryInterface(XTextFrame.class, xNameAccess.getByName(frameName));
                XPropertySet xProps = queryInterface(XPropertySet.class, xTextFrame);
                Object firstParagraphBackColor = getFirstParagraphBackColor(xTextFrame.getText());

                // 如果frame的背景色是负值，则用段落颜色覆盖
                if ((int) xProps.getPropertyValue("BackColor") < 0) {
                    xProps.setPropertyValue("BackColor", firstParagraphBackColor);
                }
                // 如果frame的背景色RGB是负值，则用段落颜色覆盖
                if ((int) xProps.getPropertyValue("BackColorRGB") < 0) {
                    xProps.setPropertyValue("BackColorRGB", firstParagraphBackColor);
                }
                // 如果frame的背景透明度为100%，则设置为0%
                if ((Byte) xProps.getPropertyValue("BackColorTransparency") == 100) {
                    xProps.setPropertyValue("BackColorTransparency", 0);
                }
            } catch (Exception e) {
                log.error("error handing frame", e);
            }
        }
    }

    /**
     * 获取段落的背景色
     */
    private static int getFirstParagraphBackColor(XText xText)
            throws WrappedTargetException, NoSuchElementException, UnknownPropertyException {
        XEnumerationAccess xEnumerationAccess = queryInterface(XEnumerationAccess.class, xText);
        XEnumeration enumeration = xEnumerationAccess.createEnumeration();

        // 遍历内容，寻找第一个段落
        while (enumeration.hasMoreElements()) {
            XTextContent xTextContent = queryInterface(XTextContent.class, enumeration.nextElement());
            if (xTextContent != null) {
                log.info("found paragraph inside: {}", xTextContent.getAnchor().getString());

                XPropertySet xParagraphProperties = queryInterface(XPropertySet.class, xTextContent);
                Object paraBackColor = xParagraphProperties.getPropertyValue("ParaBackColor");
                log.info("this paragraph's back color is {}", paraBackColor);

                // 返回找到的第一个段落的背景色
                return (int) paraBackColor;
            }
        }

        log.info("no paragraph found, will return 0");
        return 0;
    }
}
