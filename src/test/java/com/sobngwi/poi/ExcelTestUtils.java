package com.sobngwi.poi;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.Date;

import static org.junit.Assert.*;

/**
 * Utility method for testing excel files.
 *
 * @author radek.hecl
 *
 */
public class ExcelTestUtils {

    /**
     * Prevent construction.
     */
    private ExcelTestUtils() {
    }

    /**
     * Asserts 2 excel files that they have equal fields in across spread sheets.
     * All row and column indices in assertions and exceptions are 0 based.
     *
     * @param expected file with expected sheet
     * @param actual file with actual sheet
     */
    public static void assertEqualsInFields(File expected, File actual) {
        Workbook actualWb = null;
        Workbook expectedWb = null;
        try {
            actualWb = WorkbookFactory.create(actual);
            expectedWb = WorkbookFactory.create(expected);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        assertNotNull(actualWb);
        assertNotNull(expectedWb);

        assertEquals("number of sheets doesn't match", expectedWb.getNumberOfSheets(), actualWb.getNumberOfSheets());

        for (int sh = 0; sh < expectedWb.getNumberOfSheets(); sh++) {
            Sheet exSh = expectedWb.getSheetAt(sh);
            Sheet acSh = actualWb.getSheetAt(sh);
            assertEquals("sheet names doesn't match", exSh.getSheetName(), acSh.getSheetName());
            int minRow = Math.min(exSh.getFirstRowNum(), acSh.getFirstRowNum());
            int maxRow = Math.max(exSh.getLastRowNum(), acSh.getLastRowNum());
            for (int r = minRow; r <= maxRow; ++r) {
                Row exRow = exSh.getRow(r);
                Row acRow = acSh.getRow(r);
                if (exRow != null || acRow != null) {
                    if (exRow == null) {
                        if (acRow.getLastCellNum() < 0) {
                            // means that actual row is also empty
                            continue;
                        }
                        fail("row in expected file only is null: sheet = " + exSh.getSheetName() + "; row = " + r);
                    }
                    if (acRow == null) {
                        if (exRow.getLastCellNum() < 0) {
                            // means that expected row is also empty
                            continue;
                        }
                        fail("row in actual file only is null: sheet = " + exSh.getSheetName() + "; row = " + r);
                    }
                    int minCell = Math.min(exRow.getFirstCellNum(), acRow.getFirstCellNum());
                    int maxCell = Math.max(exRow.getLastCellNum(), acRow.getLastCellNum());
                    if (minCell == -1 && minCell == maxCell) {
                        // means both rows have no cells
                        continue;
                    }
                    for (int c = Math.max(0, minCell); c <= maxCell; ++c) {
                        Cell exCell = exRow.getCell(c);
                        Cell acCell = acRow.getCell(c);
                        if (exCell != null || acCell != null) {

                            if (exCell == null && acCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                                continue;
                            }
                            if (acCell == null && exCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                                continue;
                            }
                            if (exCell == null) {
                                fail("row in expected file only is null: sheet = " + exSh.getSheetName() + "; " +
                                        "row = " + r + "; cell = " + c);
                            }
                            if (acCell == null) {
                                fail("row in actual file only is null: sheet = " + exSh.getSheetName() + "; " +
                                        "row = " + r + "; cell = " + c);
                            }

                            if (exCell.getCellType() == Cell.CELL_TYPE_BLANK && acCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                                continue;
                            }
                            else if (exCell.getCellType() == Cell.CELL_TYPE_BLANK && acCell.getCellType() == Cell.CELL_TYPE_STRING) {
                                String acStr = acCell.getStringCellValue();
                                assertEquals("sheet = " + exSh.getSheetName() + "; " + "row = " + r + "; colum = " + c,
                                        "", acStr);
                            }
                            else if (exCell.getCellType() == Cell.CELL_TYPE_STRING && acCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                                String exStr = exCell.getStringCellValue();
                                assertEquals("sheet = " + exSh.getSheetName() + "; " + "row = " + r + "; colum = " + c,
                                        "", exStr);
                            }
                            else if (exCell.getCellType() == Cell.CELL_TYPE_NUMERIC && acCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                if (HSSFDateUtil.isCellDateFormatted(exCell) && HSSFDateUtil.isCellDateFormatted(acCell)) {
                                    Date ex = exCell.getDateCellValue();
                                    Date ac = acCell.getDateCellValue();
                                    assertEquals("sheet = " + exSh.getSheetName() + "; " + "row = " + r + "; colum = ", ex, ac);
                                    continue;
                                }
                                Double ex = exCell.getNumericCellValue();
                                Double ac = acCell.getNumericCellValue();
                                assertEquals("sheet = " + exSh.getSheetName() + "; " + "row = " + r + "; colum = " + c,
                                        ex, ac);
                            }
                            else if (exCell.getCellType() == Cell.CELL_TYPE_STRING && acCell.getCellType() == Cell.CELL_TYPE_STRING) {
                                String exStr = exCell.getStringCellValue();
                                String acStr = acCell.getStringCellValue();
                                assertEquals("sheet = " + exSh.getSheetName() + "; " + "row = " + r + "; colum = " + c,
                                        exStr, acStr);
                            }
                            else {
                                throw new RuntimeException("unsupported cell type, implement me: " +
                                        "expected.type = " + exCell.getCellType() + "; actual.type = " + acCell.getCellType() + "; " +
                                        "sheet = " + exSh.getSheetName() + "; " + "row = " + r + "; colum = " + c);
                            }

                        }
                    }
                }
            }
        }

    }
}
