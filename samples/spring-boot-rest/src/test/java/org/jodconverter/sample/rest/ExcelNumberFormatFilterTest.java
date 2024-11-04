package org.jodconverter.sample.rest;

import org.junit.jupiter.api.Test;

import java.math.BigDecimal;

import static org.jodconverter.sample.rest.ExcelNumberFormatFilter.*;
import static org.junit.jupiter.api.Assertions.*;

class ExcelNumberFormatFilterTest {
    @Test
    void testIsInteger() {
        assertTrue(isInteger(new BigDecimal("123")));
        assertFalse(isInteger(new BigDecimal("123.45")));
    }

    @Test
    void testGetTotalDigits() {
        assertEquals(3, getTotalDigits(new BigDecimal("123")));
        assertEquals(5, getTotalDigits(new BigDecimal("123.45")));
    }

    @Test
    void testGetDigitsBeforeDecimal() {
        assertEquals(3, getDigitsBeforeDecimal(new BigDecimal("123")));
        assertEquals(3, getDigitsBeforeDecimal(new BigDecimal("123.45")));
    }

    @Test
    void testGetDigitsAfterDecimal() {
        assertEquals(0, getDigitsAfterDecimal(new BigDecimal("123")));
        assertEquals(2, getDigitsAfterDecimal(new BigDecimal("123.45")));
    }

    @Test
    void testLongNumberWithMultipleZeros() {
        BigDecimal longNum = new BigDecimal("0.000112233456789");
        assertFalse(isInteger(longNum));
        assertEquals(16, getTotalDigits(longNum));
        assertEquals(1, getDigitsBeforeDecimal(longNum));
        assertEquals(15, getDigitsAfterDecimal(longNum));
    }

    @Test
    void testLongNumber() {
        BigDecimal longNum = new BigDecimal("0.123456789012345");
        assertFalse(isInteger(longNum));
        assertEquals(16, getTotalDigits(longNum));
        assertEquals(1, getDigitsBeforeDecimal(longNum));
        assertEquals(15, getDigitsAfterDecimal(longNum));
    }

    @Test
    void testLongInt() {
        BigDecimal longNum = new BigDecimal("1000000010000000");
        assertTrue(isInteger(longNum));
        assertEquals(16, getTotalDigits(longNum));
        assertEquals(16, getDigitsBeforeDecimal(longNum));
        assertEquals(0, getDigitsAfterDecimal(longNum));
    }
}