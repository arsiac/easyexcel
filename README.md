# easyexcel 1.0.4.1

维护EasyExcel 1.0.4 的版本

该版本依赖 poi-3.15

## Fix

1. Excel 导入无法获取到数据

2. 注解在不设置 `index` 时无法进行解析

3. 相同名称的列会被合并

4. 使用日期严格模式

5. sharedStrings中空白字符会被忽略
   ```xml
   <!-- example -->
   <si>
       <t></t>
   </si>
   ```