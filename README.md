# ToolGood.OutputWord


Wordģ�嵼�����������ToolGood.Algorithm�������棬֧�ֳ���Excel��ʽ������ģ�����������Ч����


### ��������
Nuget ������
```
    Install-Package ToolGood.OutputWord 
```

��̨����
```` csharp
            // ��ȡ����
            var helper = SqlHelperFactory.OpenSqliteFile("test.db");
            var dt = helper.ExecuteDataTable("select * from Introduction");
            var tableTests = helper.Select<TableTest>("select * from TableTest");

            ToolGood.OutputWord.WordTemplate openXmlTemplate = new ToolGood.OutputWord.WordTemplate();
            // ��������
            openXmlTemplate.SetData(dt);
            openXmlTemplate.SetListData("list", JsonConvert.SerializeObject(tableTests));

            // ����ģ�� һ
            openXmlTemplate.BuildTemplate("test.docx", "openxml_2.docx");

            // ����ģ�� ��
            var bs = openXmlTemplate.BuildTemplate("test.docx");
            File.WriteAllBytes("openxml_1.docx", bs);

````

Wordģ������

a) ��ͨ������`{������}`    

b) ʹ�ù�ʽ��`#��ʽ#`

c) ���Ƽ򻯣����ĵ������� `###����������ʽ`

wordģ�����ɺ󣬻��Զ�ɾ��`###����������ʽ`

d) ����ڲ���������ݣ�`{{��ʽ}}`

����{{list[i].Id}}

���� list Ϊ`SetListData`�����еĵ�һ��������[i] Ϊ��ĳ��

e) ����ͼƬ��`<% ͼƬ %>` ��ע��˱�ǩ��ռ�������䣬����ն��䣬�ٲ���ͼƬ