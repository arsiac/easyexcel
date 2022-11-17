package com.alibaba.excel;

import com.alibaba.excel.read.context.AnalysisContext;
import com.alibaba.excel.read.event.AnalysisEventListener;
import com.alibaba.excel.support.ExcelType;
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;

public class ExcelReaderTest {

    @Test
    public void read() throws IOException {
        try (final InputStream stream = this.getClass().getResourceAsStream("/excel/规则导出.xlsx")) {
            final ExcelReader reader = new ExcelReader(stream, ExcelType.XLSX, new AnalysisEventListener() {
                @Override
                public void invoke(Object object, AnalysisContext context) {
                    System.out.println(object);
                }
            });

            reader.read();
            reader.finish();
        }
    }

}