package com.alibaba.excel.support;

/**
 * 支持读写的数据格式
 *
 * @author jipengfei
 */
public enum ExcelType {
    XLS(".xls"),
    XLSX(".xlsx");
    private final String value;

    ExcelType(String value) {
        this.value = value;
    }

    /**
     * 根据文件名获取类型
     *
     * @param filename 文件名
     * @return Excel类型
     */
    public static ExcelType fromFilename(String filename) {
        if (filename == null) {
            return null;
        }

        for (ExcelType type : values()) {
            if (filename.endsWith(type.value)) {
                return type;
            }
        }

        return null;
    }

    public String getValue() {
        return value;
    }

}
