using System;
using System.IO;
using ToolGood.ReadyGo3;
using Newtonsoft.Json;

namespace ToolGood.WordTemplate
{
    class Program
    {
        static void Main(string[] args)
        {
            var helper = SqlHelperFactory.OpenSqliteFile("test.db");

            helper._TableHelper.TryCreateTable(typeof(Introduction));
            helper._TableHelper.TryCreateTable(typeof(TableTest));
            helper.Insert(new Introduction() {
                Name = "ToolGood",
                Achievement1 = "ToolGood.Words 类库 Star 超过1300",
                Achievement2 = "Arctic Code Vault Contributor",
                Achievement3 = " ToolGood contributed code to several repositories in the 2020 GitHub Archive Program: toolgood/ToolGood.Words, toolgood/ToolGood.Algorithm, and more! ",
                Appraisal = "懒人，挖坑党，toolgood/ToolGood.Words库golang版本没更新到最新，toolgood/ToolGood.Algorithm库javasrcipt还没写完,java没开写",
                TestInt = 18
            });
            helper.Insert(new TableTest() { Name = "行1" });
            helper.Insert(new TableTest() { Name = "行2" });
            helper.Insert(new TableTest() { Name = "行3" });


            var dt = helper.ExecuteDataTable("select * from Introduction");
            var tableTests = helper.Select<TableTest>("select * from TableTest");


            //DocxTemplate docxTemplate = new DocxTemplate();
            //docxTemplate.SetData(dt);
            ////docxTemplate.SetJsonData(JsonConvert.SerializeObject(tableTests));
            //var bs = docxTemplate.BuildTemplate("test.docx");
            //File.WriteAllBytes("docx_1.docx", bs);
            //docxTemplate.BuildTemplate("test.docx", "docx_2.docx");

            OpenXmlTemplate openXmlTemplate = new OpenXmlTemplate();
            openXmlTemplate.SetData(dt);
            openXmlTemplate.SetListData("list", JsonConvert.SerializeObject(tableTests));

            openXmlTemplate.BuildTemplate("test.docx", "openxml_2.docx");


            var bs = openXmlTemplate.BuildTemplate("test.docx");
            File.WriteAllBytes("openxml_1.docx", bs);



        }
    }


    public class Introduction
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public string Achievement1 { get; set; }
        public string Achievement2 { get; set; }
        public string Achievement3 { get; set; }

        public string Appraisal { get; set; }

        public int TestInt { get; set; }
    }

    public class TableTest
    {
        public int Id { get; set; }

        public string Name { get; set; }

    }

}
