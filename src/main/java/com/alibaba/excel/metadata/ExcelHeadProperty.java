package com.alibaba.excel.metadata;

import com.alibaba.excel.annotation.ExcelColumnNum;
import com.alibaba.excel.annotation.ExcelProperty;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 表头信息
 *
 * @author jipengfei
 */
public class ExcelHeadProperty {

    /**
     * 表头数据对应的Class
     */
    private Class<? extends ExcelRowModel> headClazz;

    /**
     * 表头名称
     */
    private List<List<String>> head;

    /**
     * Excel每列表头数据
     */
    private List<ExcelColumnProperty> columnPropertyList = new ArrayList<>();

    /**
     * key:Excel列号，value:表头数据
     */
    private final Map<Integer, ExcelColumnProperty> excelColumnPropertyMap1 = new HashMap<>();

    public ExcelHeadProperty(Class<? extends ExcelRowModel> headClazz, List<List<String>> head) {
        this.headClazz = headClazz;
        this.head = head;
        initColumnProperties();
    }

    /**
     * 初始化每列
     */
    private void initColumnProperties() {
        if (this.headClazz != null) {
            Field[] fields = this.headClazz.getDeclaredFields();
            List<List<String>> headList = new ArrayList<>();

            for (Field field : fields) {
                initOneColumnProperty(field);
            }
            //对列排序
            Collections.sort(columnPropertyList);
            if (head == null || head.isEmpty()) {
                for (ExcelColumnProperty property : columnPropertyList) {
                    headList.add(property.getHead());
                }
                this.head = headList;
            }
        }
    }

    /**
     * 初始化一列
     */
    private void initOneColumnProperty(Field f) {
        ExcelProperty p = f.getAnnotation(ExcelProperty.class);
        ExcelColumnProperty property = null;
        if (p != null) {
            property = new ExcelColumnProperty();
            property.setField(f);
            property.setHead(Arrays.asList(p.value()));
            final int index = p.index() == -1 ? columnPropertyList.size() + 1 : p.index();
            property.setIndex(index);
            property.setFormat(p.format());
            excelColumnPropertyMap1.put(index, property);
        } else {
            ExcelColumnNum columnNum = f.getAnnotation(ExcelColumnNum.class);
            if (columnNum != null) {
                property = new ExcelColumnProperty();
                property.setField(f);
                property.setIndex(columnNum.value());
                property.setFormat(columnNum.format());
                excelColumnPropertyMap1.put(columnNum.value(), property);
            }
        }
        if (property != null) {
            this.columnPropertyList.add(property);
        }

    }

    /**
     * 将表头的一行数据，转换为一列一列形式，组成表头
     *
     * @param row 表头中的一行数据
     */
    public void appendOneRow(List<String> row) {
        for (int i = 0; i < row.size(); i++) {
            List<String> list;
            if (head.size() <= i) {
                list = new ArrayList<>();
                head.add(list);
            } else {
                list = head.get(0);
            }
            list.add(row.get(i));
        }

    }

    /**
     * 根据Excel中的列号，获取Excel的表头信息
     *
     * @param columnNum 列号
     * @return ExcelColumnProperty
     */
    public ExcelColumnProperty getExcelColumnProperty(int columnNum) {
        ExcelColumnProperty property = excelColumnPropertyMap1.get(columnNum);
        if (property != null) {
            return property;
        }

        if (head != null && head.size() > columnNum) {
            List<String> columnHead = head.get(columnNum);
            for (ExcelColumnProperty columnProperty : columnPropertyList) {
                if (headEquals(columnHead, columnProperty.getHead())) {
                    property = columnProperty;
                }
            }
        }

        return property;
    }

    /**
     * 根据Excel中的列号，获取Excel的表头信息
     *
     * @param columnNum 列号
     * @return ExcelColumnProperty
     */
    public ExcelColumnProperty getExcelColumnProperty1(int columnNum) {
        return excelColumnPropertyMap1.get(columnNum);

    }

    /**
     * 判断表头是否相同
     */
    private boolean headEquals(List<String> columnHead, List<String> head) {
        boolean result = true;
        if (columnHead == null || head == null || columnHead.size() != head.size()) {
            return false;
        } else {
            for (int i = 0; i < head.size(); i++) {
                if (!head.get(i).equals(columnHead.get(i))) {
                    result = false;
                    break;
                }
            }
        }
        return result;
    }

    public Class<? extends ExcelRowModel> getHeadClazz() {
        return headClazz;
    }

    public void setHeadClazz(Class<? extends ExcelRowModel> headClazz) {
        this.headClazz = headClazz;
    }

    public List<List<String>> getHead() {
        return this.head;
    }

    public void setHead(List<List<String>> head) {
        this.head = head;
    }

    public List<ExcelColumnProperty> getColumnPropertyList() {
        return columnPropertyList;
    }

    public void setColumnPropertyList(List<ExcelColumnProperty> columnPropertyList) {
        this.columnPropertyList = columnPropertyList;
    }

    public List<CellRange> getCellRangeModels() {
        List<CellRange> ranges = new ArrayList<>();
        for (int col = 0; col < head.size(); col++) {
            List<String> columnValues = head.get(col);
            for (int row = 0; row < columnValues.size(); row++) {
                int lastRow = getLastRangRow(row, columnValues.get(row), columnValues);
                int lastColumn = getLastRangColumn(columnValues.get(row), getHeadByRowNum(row), col, row);
                if (lastRow >= 0 && lastColumn >= 0 && (lastRow > row || lastColumn > col)) {
                    ranges.add(new CellRange(row, lastRow, col, lastColumn));
                }
            }
        }
        return ranges;
    }

    public List<String> getHeadByRowNum(int rowNum) {
        List<String> l = new ArrayList<>(head.size());
        for (List<String> list : head) {
            if (list.size() > rowNum) {
                l.add(list.get(rowNum));
            } else {
                l.add(list.get(list.size() - 1));
            }
        }
        return l;
    }

    private int getLastRangColumn(String value, List<String> headByRowNum, int col, int row) {
        if (headByRowNum.indexOf(value) < col) {
            return -1;
        }
        final List<String> firstColumn = head.get(col);

        // 遍历获取当前列合并了几个单元格
        int lastIndex = col;
        for (int j = col + 1; j < headByRowNum.size(); j++) {
            if (!value.equals(headByRowNum.get(j))) {
                // 当前行不相同肯定不是同一列
                break;
            }

            final List<String> lastColumn = head.get(j);
            // 验证是否是相同的列
            for (int i = 0; i < row; i++) {
                if (!firstColumn.get(i).equals(lastColumn.get(i))) {
                    // 当前列与第一列不相同, 结束
                    return lastIndex;
                }
            }
            // 列相同, 是同一列
            lastIndex = j;
        }
        return lastIndex;

    }

    private int getLastRangRow(int j, String value, List<String> columnValue) {
        if (columnValue.indexOf(value) < j) {
            return -1;
        }
        if (value != null && value.equals(columnValue.get(columnValue.size() - 1))) {
            return getRowNum() - 1;
        } else {
            return columnValue.lastIndexOf(value);
        }
    }

    public int getRowNum() {
        int headRowNum = 0;
        for (List<String> list : head) {
            if (list != null && list.size() > headRowNum) {
                headRowNum = list.size();
            }
        }
        return headRowNum;
    }

}
