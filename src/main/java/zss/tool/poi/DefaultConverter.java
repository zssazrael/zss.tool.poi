package zss.tool.poi;

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
        }
    }
}
