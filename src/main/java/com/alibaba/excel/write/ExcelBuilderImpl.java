package com.alibaba.excel.write;

import com.alibaba.excel.metadata.ExcelColumnProperty;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.support.ExcelType;
import com.alibaba.excel.util.EasyExcelTempFile;
import com.alibaba.excel.write.context.GenerateContext;
import com.alibaba.excel.write.context.GenerateContextImpl;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

/**
 * @author jipengfei
 */
public class ExcelBuilderImpl implements ExcelBuilder {

    private GenerateContext context;

    public void init(OutputStream out, ExcelType excelType, boolean needHead) {
        //初始化时候创建临时缓存目录，用于规避POI在并发写bug
        EasyExcelTempFile.createPOIFilesDirectory();

        context = new GenerateContextImpl(out, excelType, needHead);
    }

    public void addContent(List data) {
        if (data != null && !data.isEmpty()) {
            int rowNum = context.getCurrentSheet().getLastRowNum();
            if (rowNum == 0) {
                Row row = context.getCurrentSheet().getRow(0);
                if (row == null) {
                    if (context.getExcelHeadProperty() == null || !context.needHead()) {
                        rowNum = -1;
                    }
                }
            }
            for (int i = 0; i < data.size(); i++) {
                int n = i + rowNum + 1;
                addOneRowOfDataToExcel(data.get(i), n);
            }
        }
    }

    public void addContent(List data, Sheet sheetParam) {
        context.buildCurrentSheet(sheetParam);
        addContent(data);
    }

    public void addContent(List data, Sheet sheetParam, Table table) {
        context.buildCurrentSheet(sheetParam);
        context.buildTable(table);
        addContent(data);
    }

    public void finish() {
        try {
            context.getWorkbook().write(context.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void addOneRowOfDataToExcel(List<String> oneRowData, Row row) {
        if (oneRowData != null && oneRowData.size() > 0) {
            for (int i = 0; i < oneRowData.size(); i++) {
                Cell cell = row.createCell(i);
                cell.setCellStyle(context.getCurrentContentStyle());
                cell.setCellValue(oneRowData.get(i));
            }
        }
    }

    private void addOneRowOfDataToExcel(Object oneRowData, Row row) {
        int i = 0;
        for (ExcelColumnProperty excelHeadProperty : context.getExcelHeadProperty().getColumnPropertyList()) {
            Cell cell = row.createCell(i);
            cell.setCellStyle(context.getCurrentContentStyle());
            String cellValue = null;
            try {
                final Field field = excelHeadProperty.getField();
                if (Date.class == field.getType()
                        && excelHeadProperty.getFormat() != null
                        && excelHeadProperty.getFormat().length() > 0) {
                    field.setAccessible(true);
                    final Object o = field.get(oneRowData);
                    cellValue = new SimpleDateFormat(excelHeadProperty.getFormat()).format(o);
                } else {
                    cellValue = BeanUtils.getProperty(oneRowData, field.getName());
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            if (cellValue != null) {
                cell.setCellValue(cellValue);
            } else {
                cell.setCellValue("");
            }
            i++;
        }
    }

    private void addOneRowOfDataToExcel(Object oneRowData, int n) {
        Row row = context.getCurrentSheet().createRow(n);
        if (oneRowData instanceof List) {
            addOneRowOfDataToExcel((List<String>) oneRowData, row);
        } else {
            addOneRowOfDataToExcel(oneRowData, row);
        }
    }

}
