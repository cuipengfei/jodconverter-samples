package org.jodconverter.sample.rest;

import com.sun.star.lang.XComponent;
import com.sun.star.text.XText;
import com.sun.star.text.XTextDocument;
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
        }

        // This is a word document, so we can apply specific filters or processing here.
        XText xText = xTextDocument.getText();
        String textContent = xText.getString();
        log.info("Word document text: {}", textContent);
    }
}
