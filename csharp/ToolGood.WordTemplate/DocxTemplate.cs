using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using Novacode;
using ToolGood.Algorithm;

namespace ToolGood.WordTemplate
{
    /// <summary>
    /// 原Docx 组件，个人免费，不能商用。 
    /// 还有一点，经测试文件会变大。
    /// </summary>
    public class DocxTemplate : AlgorithmEngine
    {
        private readonly static Regex _tempEngine = new Regex("^###([^:]*):(.*)$");// 定义临时变量
        private readonly static Regex _tempMatch = new Regex("(#[^#]*#)");// 
        private readonly static Regex _simplifyMatch = new Regex(@"(\{[^\}]*\})");//简化文本 只读取字段
        private DataTable _dt;

        public byte[] BuildTemplate(DataTable dataTable, string fileName)
        {
            _dt = dataTable;
            using (DocX document = DocX.Load(fileName))
            {
                ReplaceTemplate(document);
                using (var ms = new MemoryStream())
                {
                    document.SaveAs(ms);
                    return ms.ToArray();
                }
            }
        }

        public void BuildTemplate(DataTable dataTable, string fileName, string newFilePath)
        {
            _dt = dataTable;
            using (DocX document = DocX.Load(fileName))
            {
                ReplaceTemplate(document);
                document.SaveAs(newFilePath);
            }
        }

        public byte[] BuildTemplate(string jsonData, string fileName)
        {
            _dt = null;
            this.AddParameterFromJson(jsonData);
            using (DocX document = DocX.Load(fileName))
            {
                ReplaceTemplate(document);
                using (var ms = new MemoryStream())
                {
                    document.SaveAs(ms);
                    return ms.ToArray();
                }
            }
        }

        public void BuildTemplate(string jsonData, string fileName, string newFilePath)
        {
            _dt = null;
            this.AddParameterFromJson(jsonData);
            using (DocX document = DocX.Load(fileName))
            {
                ReplaceTemplate(document);
                document.SaveAs(newFilePath);
            }
        }

        private void ReplaceTemplate(DocX document)
        {
            var tempMatches = new List<string>();
            List<Paragraph> deleteParagraph = new List<Paragraph>();
            foreach (var paragraph in document.Paragraphs)
            {
                var text = paragraph.Text.Trim();
                var m = _tempEngine.Match(text);
                if (m.Success)
                {
                    var name = m.Groups[1].Value.Trim();
                    var engine = m.Groups[2].Value.Trim();
                    var value = this.TryEvaluate(engine, "");
                    this.AddParameter(name, value);
                    deleteParagraph.Add(paragraph);
                    continue;
                }
                var m2 = _tempMatch.Match(text);
                if (m2.Success)
                {
                    tempMatches.Add(m2.Groups[1].Value);
                    continue;
                }
                var m3 = _simplifyMatch.Match(text);
                if (m3.Success)
                {
                    tempMatches.Add(m3.Groups[1].Value);
                    continue;
                }
            }
            foreach (var paragraph in deleteParagraph)
            {
                paragraph.Remove(false);
            }
            foreach (var m in tempMatches)
            {
                string value;
                if (m.StartsWith("#"))
                {
                    value = this.TryEvaluate(m.Trim('#'), "");
                }
                else
                {
                    value = this.TryEvaluate(m.Replace("{", "[").Replace("}", "]"), "");
                }
                document.ReplaceText(m, value);
            }
        }
        protected override Operand GetParameter(string parameter)
        {
            parameter = parameter.Trim();
            if (_dt != null && _dt.Rows.Count > 0 && _dt.Columns.Contains(parameter))
            {
                if (_dt.Rows[0].IsNull(parameter))
                {
                    return Operand.CreateNull();
                }
                var obj = _dt.Rows[0][parameter];
                { if (obj is Int16 val) { return val; } }
                { if (obj is Int32 val) { return val; } }
                { if (obj is Int64 val) { return val; } }
                { if (obj is UInt16 val) { return val; } }
                { if (obj is UInt32 val) { return val; } }
                { if (obj is UInt64 val) { return val; } }
                { if (obj is Single val) { return val; } }
                { if (obj is Double val) { return val; } }
                { if (obj is Decimal val) { return val; } }
                { if (obj is DateTime val) { return val; } }
                { if (obj is TimeSpan val) { return val; } }
                { if (obj is Boolean val) { return val; } }
                { if (obj is String val) { return val; } }

                return _dt.Rows[0][parameter]?.ToString() ?? "";
            }
            return base.GetParameter(parameter);
        }
    }
}
