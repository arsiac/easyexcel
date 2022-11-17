package com.alibaba.excel.read.v07;

import com.alibaba.excel.annotation.FieldType;
import com.alibaba.excel.read.context.AnalysisContext;
import com.alibaba.excel.read.event.AnalysisEventRegisterCenter;
import com.alibaba.excel.read.event.OneRowAnalysisFinishEvent;
import com.alibaba.excel.util.ExcelXmlConstants;
import com.alibaba.excel.util.PositionUtils;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.Arrays;
import java.util.List;

import static com.alibaba.excel.util.ExcelXmlConstants.CELL_VALUE_TAG;
import static com.alibaba.excel.util.ExcelXmlConstants.CELL_VALUE_TAG_1;
import static com.alibaba.excel.util.ExcelXmlConstants.DIMENSION;
import static com.alibaba.excel.util.ExcelXmlConstants.DIMENSION_REF;
import static com.alibaba.excel.util.ExcelXmlConstants.ROW_TAG;

/**
 * @author jipengfei
 */
public class RowHandler extends DefaultHandler {

    private String currentCellIndex;

    private FieldType currentCellType;

    private int curRow;

    private int curCol;

    private String[] curRowContent = new String[30];

    private String currentCellValue;

    private SharedStringsTable sst;

    private final AnalysisContext analysisContext;

    private final AnalysisEventRegisterCenter registerCenter;

    private final List<String> sharedStringList;

    public RowHandler(AnalysisEventRegisterCenter registerCenter, SharedStringsTable sst,
                      AnalysisContext analysisContext, List<String> sharedStringList) {
        this.registerCenter = registerCenter;
        this.analysisContext = analysisContext;
        this.sst = sst;
        this.sharedStringList = sharedStringList;

    }

    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {

        setTotalRowCount(name, attributes);

        startCell(name, attributes);

        startCellValue(name);

    }

    private void startCellValue(String name) {
        if (name.equals(CELL_VALUE_TAG) || name.equals(CELL_VALUE_TAG_1)) {
            // initialize current cell value
            currentCellValue = "";
        }
    }

    private void startCell(String name, Attributes attributes) {
        if (ExcelXmlConstants.CELL_TAG.equals(name)) {
            currentCellIndex = attributes.getValue(ExcelXmlConstants.POSITION);
            int nextRow = PositionUtils.getRow(currentCellIndex);
            if (nextRow > curRow) {
                curRow = nextRow;
            }
            analysisContext.setCurrentRowNum(curRow);
            curCol = PositionUtils.getCol(currentCellIndex);

            String cellType = attributes.getValue("t");
            currentCellType = FieldType.EMPTY;
            if (cellType != null && cellType.equals("s")) {
                currentCellType = FieldType.STRING;
            }
        }
    }

    private void endCellValue(String name) {
        // ensure size
        if (curCol >= curRowContent.length) {
            curRowContent = Arrays.copyOf(curRowContent, (int)(curCol * 1.5));
        }
        if (CELL_VALUE_TAG.equals(name)) {
            if (currentCellType == FieldType.STRING) {
                int idx = Integer.parseInt(currentCellValue);
                if (idx < sharedStringList.size()) {
                    currentCellValue = sharedStringList.get(idx);
                } else {
                    currentCellValue = "";
                }
                currentCellType = FieldType.EMPTY;
            }
            curRowContent[curCol] = currentCellValue;
        } else if (CELL_VALUE_TAG_1.equals(name)) {
            curRowContent[curCol] = currentCellValue;
        }
    }

    @Override
    public void endElement(String uri, String localName, String name) throws SAXException {
        endRow(name);
        endCellValue(name);
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        currentCellValue += new String(ch, start, length);
    }

    private void setTotalRowCount(String name, Attributes attributes) {
        if (DIMENSION.equals(name)) {
            String d = attributes.getValue(DIMENSION_REF);
            String totalStr = d.substring(d.indexOf(":") + 1, d.length());
            String c = totalStr.toUpperCase().replaceAll("[A-Z]", "");
            analysisContext.setTotalCount(Integer.parseInt(c));
        }

    }

    private void endRow(String name) {
        if (name.equals(ROW_TAG)) {
            registerCenter.notifyListeners(new OneRowAnalysisFinishEvent(Arrays.asList(curRowContent)));
            curRowContent = new String[30];
        }
    }

}

