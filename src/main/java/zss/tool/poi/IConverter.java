package zss.tool.poi;

import org.apache.poi.ss.usermodel.Cell;

import zss.tool.Version;

@Version("2017.05.11")
public interface IConverter {
    public void convert(final Cell cell, final Object value);
}
