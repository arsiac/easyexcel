package com.alibaba.excel.read;

import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.read.context.AnalysisContext;
import com.alibaba.excel.read.exception.ExcelAnalysisException;
import com.alibaba.excel.read.v07.RowHandler;
import com.alibaba.excel.read.v07.XMLTempFile;
import com.alibaba.excel.read.v07.XmlParserFactory;
import com.alibaba.excel.util.FileUtil;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbookPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

/**
 * @author jipengfei
 */
public class SaxAnalyserV07 extends BaseSaxAnalyser {

    private SharedStringsTable sharedStringsTable;

    private final List<String> sharedStringList = new LinkedList<>();

    private List<SheetSource> sheetSourceList = new ArrayList<>();

    private boolean use1904WindowDate = false;

    private final String path;

    private final File tmpFile;

    private final String workBookXMLFilePath;

    private final String sharedStringXMLFilePath;

    public SaxAnalyserV07(AnalysisContext analysisContext) throws Exception {
        this.analysisContext = analysisContext;
        this.path = XMLTempFile.createPath();
        this.tmpFile = new File(XMLTempFile.getTmpFilePath(path));
        this.workBookXMLFilePath = XMLTempFile.getWorkBookFilePath(path);
        this.sharedStringXMLFilePath = XMLTempFile.getSharedStringFilePath(path);
        start();
    }

    @Override
    protected void execute() {
        try {
            Sheet sheet = analysisContext.getCurrentSheet();
            if (!isAnalysisAllSheets(sheet)) {
                if (this.sheetSourceList.size() < sheet.getSheetNo() || sheet.getSheetNo() == 0) {
                    return;
                }
                InputStream sheetInputStream = this.sheetSourceList.get(sheet.getSheetNo() - 1).getInputStream();
                parseXmlSource(sheetInputStream);
                return;
            }
            int i = 0;
            for (SheetSource sheetSource : this.sheetSourceList) {
                i++;
                this.analysisContext.setCurrentSheet(new Sheet(i));
                parseXmlSource(sheetSource.getInputStream());
            }
        } catch (Exception e) {
            stop();
            throw new ExcelAnalysisException(e);
        }
    }

    private boolean isAnalysisAllSheets(Sheet sheet) {
        if (sheet == null) {
            return true;
        }
        return sheet.getSheetNo() < 0;
    }

    public void stop() {
        for (SheetSource sheet : sheetSourceList) {
            if (sheet.getInputStream() != null) {
                try {
                    sheet.getInputStream().close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        FileUtil.deleteFile(path);
    }

    private void parseXmlSource(InputStream inputStream) {
        try {
            ContentHandler handler =
                    new RowHandler(this, this.sharedStringsTable, this.analysisContext, sharedStringList);
            XmlParserFactory.parse(inputStream, handler);
            inputStream.close();
        } catch (Exception e) {
            try {
                inputStream.close();
            } catch (IOException e1) {
                e1.printStackTrace();
            }
            throw new ExcelAnalysisException(e);
        }
    }

    public List<Sheet> getSheets() {
        List<Sheet> sheets = new ArrayList<>();
        try {
            int i = 1;
            for (SheetSource sheetSource : this.sheetSourceList) {
                Sheet sheet = new Sheet(i, 0);
                sheet.setSheetName(sheetSource.getSheetName());
                i++;
                sheets.add(sheet);
            }
        } catch (Exception e) {
            stop();
            throw new ExcelAnalysisException(e);
        }

        return sheets;
    }

    private void start() throws IOException, XmlException, ParserConfigurationException, SAXException {
        createTmpFile();
        unZipTempFile();
        initSharedStringsTable();
        initUse1904WindowDate();
        initSheetSourceList();
    }

    private void createTmpFile() {
        FileUtil.writeFile(tmpFile, analysisContext.getInputStream());
    }

    private void unZipTempFile() throws IOException {
        FileUtil.unzip(path, tmpFile);
    }

    private void initSheetSourceList() throws IOException, ParserConfigurationException, SAXException {
        this.sheetSourceList = new ArrayList<>();
        InputStream workbookXml = Files.newInputStream(Paths.get(this.workBookXMLFilePath));
        XmlParserFactory.parse(workbookXml, new WorkBookHandler());
        workbookXml.close();
        Collections.sort(sheetSourceList);
    }

    private void initUse1904WindowDate() throws IOException, XmlException {
        InputStream workbookXml = Files.newInputStream(Paths.get(workBookXMLFilePath));
        WorkbookDocument ctWorkbook = WorkbookDocument.Factory.parse(workbookXml);
        CTWorkbook wb = ctWorkbook.getWorkbook();
        CTWorkbookPr prefix = wb.getWorkbookPr();
        if (prefix != null) {
            this.use1904WindowDate = prefix.getDate1904();
        }
        this.analysisContext.setUse1904WindowDate(use1904WindowDate);
        workbookXml.close();
    }

    private void initSharedStringsTable() throws IOException, ParserConfigurationException, SAXException {
        InputStream inputStream = Files.newInputStream(Paths.get(this.sharedStringXMLFilePath));
        XmlParserFactory.parse(inputStream, new SharedStringHandler());
        inputStream.close();
    }

    private class SharedStringHandler extends DefaultHandler {
        private static final int SHARED_NONE = 0;

        private static final int SHARED_NEW = 1;

        private static final int SHARED_APPEND = 2;

        private String currentQName = "";

        private int currentType = SHARED_NONE;

        private String currentValue = "";

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            final String tagSi = "si";
            final String tagT = "t";

            if (tagSi.equals(qName) || tagT.equals(qName)) {
                String beforeQName = currentQName;
                currentQName = qName;

                if (tagT.equals(currentQName) && (tagT.equals(beforeQName))) {
                    currentType = SHARED_APPEND;
                } else if (tagT.equals(currentQName) && (tagSi.equals(beforeQName))) {
                    currentType = SHARED_NEW;
                }
            }
        }

        @Override
        public void endElement(String uri, String localName, String qName) {
            if (Objects.equals(SHARED_APPEND, currentType) || Objects.equals(SHARED_NEW, currentType)) {
                sharedStringList.add(currentValue);
                currentValue = "";
                currentType = SHARED_NONE;
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            if (Objects.equals(SHARED_APPEND, currentType)) {
                String pre = sharedStringList.get(sharedStringList.size() - 1);
                currentValue = pre + new String(ch, start, length);
                sharedStringList.remove(sharedStringList.size() - 1);
            } else if (Objects.equals(SHARED_NEW, currentType)) {
                currentValue = new String(ch, start, length);
            }
        }

    }

    private class WorkBookHandler extends DefaultHandler {
        private int id = 0;

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attrs) throws SAXException {
            if (qName.toLowerCase(Locale.US).equals("sheet")) {
                String name = null;
                id++;
                for (int i = 0; i < attrs.getLength(); i++) {
                    final String attrName = getAttributeName(attrs, i);
                    if (attrName.equals("name")) {
                        name = attrs.getValue(i);
                    } else if (attrName.equals("r:id")) {
                        final String filepath = XMLTempFile.getSheetFilePath(path, id);
                        try {
                            InputStream inputStream = new FileInputStream(filepath);
                            sheetSourceList.add(new SheetSource(id, name, inputStream));
                        } catch (FileNotFoundException e) {
                            throw new IllegalStateException("file not found: " + filepath, e);
                        }
                    }
                }

            }
        }

        private String getAttributeName(Attributes attributes, int i) {
            String attrName = attributes.getLocalName(i);
            if (attrName == null || attrName.length() == 0) {
                attrName = attributes.getQName(i);
            }

            return attrName.toLowerCase(Locale.US);
        }
    }

    private static class SheetSource implements Comparable<SheetSource> {

        private int id;

        private String sheetName;

        private InputStream inputStream;

        public SheetSource(int id, String sheetName, InputStream inputStream) {
            this.id = id;
            this.sheetName = sheetName;
            this.inputStream = inputStream;
        }

        public String getSheetName() {
            return sheetName;
        }

        public void setSheetName(String sheetName) {
            this.sheetName = sheetName;
        }

        public InputStream getInputStream() {
            return inputStream;
        }

        public void setInputStream(InputStream inputStream) {
            this.inputStream = inputStream;
        }

        public int getId() {
            return id;
        }

        public void setId(int id) {
            this.id = id;
        }

        public int compareTo(SheetSource o) {
            return Integer.compare(this.id, o.id);
        }

    }

}
