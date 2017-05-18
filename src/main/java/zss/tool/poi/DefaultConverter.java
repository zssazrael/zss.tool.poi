package zss.tool.poi;

import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;

import zss.tool.Version;

@Version("2017.05.11")
public class DefaultConverter implements IConverter {
    private static final DefaultConverter instance = new DefaultConverter();

    public static DefaultConverter getInstance() {
        return instance;
    }

    private DefaultConverter() {
    }

    @Override
    public void convert(Cell cell, Object value) {
        if (value instanceof CharSequence) {
            cell.setCellValue(value.toString());
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        }
    }
}
