package zss.tool.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

import zss.tool.Version;

@Version("2017.05.11")
public class CellSchema {
    private final String name;
    private String title;
    private CellStyle titleStyle;
    private CellStyle style;
    private Object defaultValue;
    private IConverter converter = DefaultConverter.getInstance();

    public CellSchema setTitleStyle(CellStyle titleStyle) {
        this.titleStyle = titleStyle;
        return this;
    }

    public CellSchema setConverter(IConverter converter) {
        this.converter = converter == null ? DefaultConverter.getInstance() : converter;
        return this;
    }

    public CellSchema setDefaultValue(Object defaultValue) {
        this.defaultValue = defaultValue;
        return this;
    }

    public CellSchema setStyle(CellStyle style) {
        this.style = style;
        return this;
    }

    public CellStyle getStyle() {
        return style;
    }

    public String getName() {
        return name;
    }

    public String getTitle() {
        return title;
    }

    public CellSchema setTitle(String title) {
        this.title = title;
        return this;
    }

    public CellSchema(final String name) {
        this.name = name;
    }

    public void setStyle(final Cell cell) {
        if (style != null) {
            cell.setCellStyle(style);
        }
    }

    public void setValue(Cell cell, Object value) {
        if (value == null) {
            if (defaultValue != null) {
                converter.convert(cell, defaultValue);
            }
        } else {
            converter.convert(cell, value);
        }
    }

    public void setTitle(final Cell cell) {
        if (titleStyle != null) {
            cell.setCellStyle(titleStyle);
        }
        if (title == null) {
            cell.setCellValue(name);
        } else {
            cell.setCellValue(title);
        }
    }
}
