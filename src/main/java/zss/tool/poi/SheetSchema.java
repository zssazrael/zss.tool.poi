package zss.tool.poi;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import zss.tool.Version;

@Version("2017.05.11")
public class SheetSchema {
    private final Sheet sheet;
    private final List<CellSchema> cellSchemaList = new LinkedList<>();
    private int rowNumber = 1;
    private final MethodMap methodMap = new MethodMap();

    public CellSchema addCellSchema(final String name) {
        final CellSchema cellSchema = new CellSchema(name);
        cellSchemaList.add(cellSchema);
        return cellSchema;
    }

    public void fillTitle() {
        final Row row = POITool.defaultRow(sheet, 0);
        int columnIndex = 0;
        for (CellSchema cellSchema : cellSchemaList) {
            final Cell cell = POITool.defaultCell(row, columnIndex++);
            cellSchema.setTitle(cell);
        }
    }

    public SheetSchema(final Sheet sheet) {
        this.sheet = sheet;
    }

    public void fill(final Object value) {
        final Row row = POITool.defaultRow(sheet, rowNumber++);
        methodMap.setType(value.getClass());
        int columnIndex = 0;
        for (CellSchema cellSchema : cellSchemaList) {
            final Cell cell = POITool.defaultCell(row, columnIndex++);
            cellSchema.setStyle(cell);
            if (value instanceof Map) {
                cellSchema.setValue(cell, ((Map<?, ?>) value).get(cellSchema.getName()));
            } else {
                final Method method = methodMap.getMethod(cellSchema.getName());
                if (method == null) {
                    cellSchema.setValue(cell, null);
                } else {
                    try {
                        cellSchema.setValue(cell, method.invoke(value));
                    } catch (IllegalAccessException e) {
                        cellSchema.setValue(cell, null);
                    } catch (IllegalArgumentException e) {
                        cellSchema.setValue(cell, null);
                    } catch (InvocationTargetException e) {
                        cellSchema.setValue(cell, null);
                    }
                }
            }
        }
    }

    public void autoSizeColumn() {
        for (int i = cellSchemaList.size() - 1; i > -1; i--) {
            sheet.autoSizeColumn(i, true);
        }
    }

    public void setAutoFilter() {
        sheet.setAutoFilter(new CellRangeAddress(0, rowNumber - 1, 0, cellSchemaList.size() - 1));
    }
}
