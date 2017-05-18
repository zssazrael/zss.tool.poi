package zss.tool.poi;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import zss.tool.Version;

@Version("2017.05.09")
public class POITool {
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
