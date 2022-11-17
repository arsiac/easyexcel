package com.alibaba.excel.read.modelbuild;

import com.alibaba.excel.metadata.ExcelColumnProperty;
import com.alibaba.excel.metadata.ExcelHeadProperty;
import com.alibaba.excel.metadata.ExcelRowModel;
import com.alibaba.excel.read.context.AnalysisContext;
import com.alibaba.excel.read.event.AnalysisEventListener;
import com.alibaba.excel.util.TypeUtil;
import com.alibaba.excel.write.exception.ExcelGenerateException;
import org.apache.commons.beanutils.BeanUtils;

import java.util.List;

/**
 * 监听POI Sax解析的每行结果
 *
 * @author jipengfei
 */
public class ModelBuildEventListener extends AnalysisEventListener<List<String>> {

    @Override
    public void invoke(List<String> object, AnalysisContext context) {
        if (context.getExcelHeadProperty() != null && context.getExcelHeadProperty().getHeadClazz() != null) {
            try {
                Object resultModel = buildUserModel(context, object);
                context.setCurrentRowAnalysisResult(resultModel);
                context.setCurrentException(null);
            } catch (Exception e) {
                context.setCurrentRowAnalysisResult(null);
                context.setCurrentException(e);
            }
        }
    }

    private Object buildUserModel(AnalysisContext context, List<String> stringList) {
        ExcelHeadProperty excelHeadProperty = context.getExcelHeadProperty();

        final Class<? extends ExcelRowModel> clazz = context.getExcelHeadProperty().getHeadClazz();
        Object model = createInstance(clazz);
        for (int i = 0; i < stringList.size(); i++) {
            ExcelColumnProperty columnProperty = excelHeadProperty.getExcelColumnProperty1(i);
            if (columnProperty != null) {
                Object value = TypeUtil.convert(stringList.get(i), columnProperty.getField(),
                        columnProperty.getFormat(), context.use1904WindowDate());
                if (value != null) {
                    try {
                        BeanUtils.setProperty(model, columnProperty.getField().getName(), value);
                    } catch (Exception e) {
                        final String msg = columnProperty.getField().getName() + " can not set value " + value;
                        throw new ExcelGenerateException(msg, e);
                    }
                }
            }
        }

        return model;
    }

    private static Object createInstance(Class<? extends ExcelRowModel> clazz) {
        try {
            return clazz.newInstance();
        } catch (Exception e) {
            throw new ExcelGenerateException(e);
        }
    }

}
