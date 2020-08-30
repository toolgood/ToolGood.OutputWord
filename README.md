# ToolGood.OutputWord


Word模板导出组件，采用ToolGood.Algorithm计算引擎，支持常用Excel公式，做到模板与代码分离的效果。


### 快速上手
Nuget 命令行
```
    Install-Package ToolGood.OutputWord 
```

后台代码
```` csharp
            // 获取数据
            var helper = SqlHelperFactory.OpenSqliteFile("test.db");
            var dt = helper.ExecuteDataTable("select * from Introduction");
            var tableTests = helper.Select<TableTest>("select * from TableTest");

            ToolGood.OutputWord.WordTemplate openXmlTemplate = new ToolGood.OutputWord.WordTemplate();
            // 加载数据
            openXmlTemplate.SetData(dt);
            openXmlTemplate.SetListData("list", JsonConvert.SerializeObject(tableTests));

            // 生成模板 一
            openXmlTemplate.BuildTemplate("test.docx", "openxml_2.docx");

            // 生成模板 二
            var bs = openXmlTemplate.BuildTemplate("test.docx");
            File.WriteAllBytes("openxml_1.docx", bs);

````

Word模板设置

a) 普通变量：`{变量名}`    

b) 使用公式：`#公式#`

c) 名称简化：在文档最后添加 `###变量名：公式`

word模板生成后，会自动删除`###变量名：公式`

d) 表格内插入多条数据：`{{公式}}`

例：{{list[i].Id}}

其中 list 为`SetListData`方法中的第一个参数，[i] 为第某行

e) 插入图片：`<% 图片 %>` ，注意此标签会占整个段落，先清空段落，再插入图片