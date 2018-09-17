package zss.tool.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import zss.tool.IOTool;
import zss.tool.LoggedException;
import zss.tool.Version;

@Version("2018.08.25")
public class POITool {
    private static final Logger LOGGER = LoggerFactory.getLogger(POITool.class);

    public static Workbook newWorkbook(final PushbackInputStream stream) {
        try {
            if (NPOIFSFileSystem.hasPOIFSHeader(stream)) {
                return new HSSFWorkbook(stream);
            }
            if (DocumentFactoryHelper.hasOOXMLHeader(stream)) {
                return new XSSFWorkbook(stream);
            }
        } catch (IOException e) {
            LOGGER.error(e.getMessage(), e);
            throw new LoggedException();
        }
        throw new LoggedException();
    }

    public static Workbook newWorkbook(final InputStream stream) {
        if (stream instanceof PushbackInputStream) {
            return newWorkbook((PushbackInputStream) stream);
        }
        return newWorkbook(new PushbackInputStream(stream, 8));
    }

    public static Workbook newWorkbook(final File file) {
        final FileInputStream stream = IOTool.newFileInputStream(file);
        try {
            return newWorkbook(stream);
        } finally {
            IOTool.close(stream);
        }
    }

    public static void setSheetName(final Sheet sheet, final String name) {
        final Workbook workbook = sheet.getWorkbook();
        final int index = workbook.getSheetIndex(sheet);
        workbook.setSheetName(index, name);
    }

    public static Sheet defaultSheet(final Workbook workbook, final String name) {
        final Sheet sheet = workbook.getSheet(name);
        if (sheet == null) {
            return workbook.createSheet(name);
        }
        return sheet;
    }

    public static Row defaultRow(final Sheet sheet, final int rowNumber) {
        final Row row = sheet.getRow(rowNumber);
        if (row == null) {
            return sheet.createRow(rowNumber);
        }
        return row;
    }

    public static Cell defaultCell(Row row, int columnIndex) {
        final Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return row.createCell(columnIndex);
        }
        return cell;
    }

    public static void setBorder(final CellStyle style, final BorderStyle borderStyle) {
        style.setBorderBottom(borderStyle);
        style.setBorderTop(borderStyle);
        style.setBorderLeft(borderStyle);
        style.setBorderRight(borderStyle);
    }
}
