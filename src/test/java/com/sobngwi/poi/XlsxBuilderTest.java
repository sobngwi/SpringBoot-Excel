package com.sobngwi.poi;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.builder.ToStringBuilder;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.IOException;

/**
 *
 * @author radek.hecl
 */
public class XlsxBuilderTest {

    /**
     * Creates new instance.
     */
    public XlsxBuilderTest() {
    }

    /**
     * Sets up the test environment.
     *
     * @throws IOException in case there is a problem with test directory setup
     */
    @Before
    public void setUp() throws IOException {
        File baseDir = new File("tmp");
        if (!baseDir.isDirectory()) {
            baseDir.mkdirs();
            if (!baseDir.isDirectory()) {
                throw new RuntimeException("unable to create test directory");
            }
        }
    }

    /**
     * Tests the build process.
     *
     * @throws IOException in case of error during writing the out file
     */
    @Test
    public void testProcess() throws IOException {
        byte[] report = new XlsxBuilder().
                startSheet("Dream cars").
                startRow().
                setRowTitleHeight().
                addTitleTextColumn("Dream cars").
                startRow().
                setRowTitleHeight().
                setRowThinBottomBorder().
                addBoldTextLeftAlignedColumn("Dreamed By:").
                addTextLeftAlignedColumn("John Seaman").
                startRow().
                startRow().
                setRowTitleHeight().
                setRowThickTopBorder().
                setRowThickBottomBorder().
                addBoldTextCenterAlignedColumn("Type").
                addBoldTextCenterAlignedColumn("Color").
                addBoldTextCenterAlignedColumn("Reason").
                startRow().
                addTextLeftAlignedColumn("Ferrari").
                addTextLeftAlignedColumn("Green").
                addTextLeftAlignedColumn("It looks nice").
                startRow().
                addTextLeftAlignedColumn("Lamborgini").
                addTextLeftAlignedColumn("Yellow").
                addTextLeftAlignedColumn("It's fast enough").
                startRow().
                addTextLeftAlignedColumn("Bugatti").
                addTextLeftAlignedColumn("Blue").
                addTextLeftAlignedColumn("Price is awesome").
                startRow().
                setRowThinTopBorder().
                startRow().
                startRow().
                addTextLeftAlignedColumn("This is just a footer and I use it instead of 'Lorem ipsum dolor...'").
                setColumnSize(0, "XXXXXXXXXXXXX".length()).
                setAutoSizeColumn(1).
                setAutoSizeColumn(2).
                build();
        File resFile = new File("tmp/XlsxBuilder.xlsx");
        if (resFile.isFile()) {
            resFile.delete();
        }
        FileUtils.writeByteArrayToFile(resFile, report);
        ExcelTestUtils.assertEqualsInFields(new File("src/test/resources/XlsxBuilder-expected.xlsx"), resFile);
    }

    @Override
    public String toString() {
        return ToStringBuilder.reflectionToString(this);
    }

}
