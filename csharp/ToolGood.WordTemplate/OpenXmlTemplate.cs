﻿using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ToolGood.Algorithm;

namespace ToolGood.WordTemplate
{
    /// <summary>
    /// DocumentFormat.OpenXml MIT协议，可以商用
    /// </summary>
    public class OpenXmlTemplate : AlgorithmEngine
    {
        private readonly static Regex _tempEngine = new Regex("^###([^:：]*)[:：](.*)$");// 定义临时变量
        private readonly static Regex _tempMatch = new Regex("(#[^#]*#)");// 
        private readonly static Regex _simplifyMatch = new Regex(@"(\{[^\}]*\})");//简化文本 只读取字段
        private DataTable _dt;

        public byte[] BuildTemplate(DataTable dataTable, string fileName)
        {
            _dt = dataTable;
            this.ClearParameters();
            var bs = File.ReadAllBytes(fileName);
            using (var ms = new MemoryStream(bs))
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(ms, true)) {
                var body = wordDoc.MainDocumentPart.Document.Body;
                ReplaceTemplate(body);

                using (var ms2 = new MemoryStream()) {
                    wordDoc.Clone(ms2);
                    return ms2.ToArray();
                }
            };
        }

        public void BuildTemplate(DataTable dataTable, string fileName, string newFilePath)
        {
            _dt = dataTable;
            this.ClearParameters();
            var bs = File.ReadAllBytes(fileName);
            using (var ms = new MemoryStream(bs))
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(ms, true)) {
                var body = wordDoc.MainDocumentPart.Document.Body;
                ReplaceTemplate(body);
                wordDoc.SaveAs(newFilePath);
            };
        }

        public byte[] BuildTemplate(string jsonData, string fileName)
        {
            _dt = null;
            this.ClearParameters();
            this.AddParameterFromJson(jsonData);
            var bs = File.ReadAllBytes(fileName);
            using (var ms = new MemoryStream(bs))
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(ms, true)) {
                var body = wordDoc.MainDocumentPart.Document.Body;
                ReplaceTemplate(body);
                return ms.ToArray();
            };
        }

        public void BuildTemplate(string jsonData, string fileName, string newFilePath)
        {
            _dt = null;
            this.ClearParameters();
            this.AddParameterFromJson(jsonData);
            var bs = File.ReadAllBytes(fileName);
            using (var ms = new MemoryStream(bs))
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(ms, true)) {
                var body = wordDoc.MainDocumentPart.Document.Body;
                ReplaceTemplate(body);
                wordDoc.SaveAs(newFilePath);
            };
        }



        private void ReplaceTemplate(Body body)
        {
            var tempMatches = new List<string>();
            List<Paragraph> deleteParagraph = new List<Paragraph>();
            foreach (var paragraph in body.Descendants<Paragraph>()) {
                var text = paragraph.InnerText.Trim();
                var m = _tempEngine.Match(text);
                if (m.Success) {
                    var name = m.Groups[1].Value.Trim();
                    var engine = m.Groups[2].Value.Trim();
                    var value = this.TryEvaluate(engine, "");
                    this.AddParameter(name, value);
                    deleteParagraph.Add(paragraph);
                    continue;
                }
                var m2 = _tempMatch.Match(text);
                if (m2.Success) {
                    tempMatches.Add(m2.Groups[1].Value);
                    continue;
                }
                var m3 = _simplifyMatch.Match(text);
                if (m3.Success) {
                    tempMatches.Add(m3.Groups[1].Value);
                    continue;
                }
            }
            foreach (var paragraph in deleteParagraph) {
                paragraph.Remove();
            }

            foreach (var m in tempMatches) {
                string value;
                if (m.StartsWith("#")) {
                    value = this.TryEvaluate(m.Trim('#'), "");
                } else {
                    value = this.TryEvaluate(m.Replace("{", "[").Replace("}", "]"), "");
                }
                foreach (var paragraph in body.Descendants<Paragraph>()) {
                    ReplaceText(paragraph, m, value);
                }
            }
        }

        protected override Operand GetParameter(string parameter)
        {
            parameter = parameter.Trim();
            if (_dt != null && _dt.Rows.Count > 0 && _dt.Columns.Contains(parameter)) {
                if (_dt.Rows[0].IsNull(parameter)) {
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


        #region OpenXml ReplaceText
        // 代码来源 https://stackoverflow.com/questions/19094388/openxml-replace-text-in-all-document

        /// <summary>
        /// Find/replace within the specified paragraph.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="find"></param>
        /// <param name="replaceWith"></param>
        private void ReplaceText(Paragraph paragraph, string find, string replaceWith)
        {
            var texts = paragraph.Descendants<Text>();
            for (int t = 0; t < texts.Count(); t++) {   // figure out which Text element within the paragraph contains the starting point of the search string
                Text txt = texts.ElementAt(t);
                for (int c = 0; c < txt.Text.Length; c++) {
                    var match = IsMatch(texts, t, c, find);
                    if (match != null) {   // now replace the text
                        string[] lines = replaceWith.Replace(Environment.NewLine, "\r").Split('\n', '\r'); // handle any lone n/r returns, plus newline.

                        int skip = lines[lines.Length - 1].Length - 1; // will jump to end of the replacement text, it has been processed.

                        if (c > 0)
                            lines[0] = txt.Text.Substring(0, c) + lines[0];  // has a prefix
                        if (match.EndCharIndex + 1 < texts.ElementAt(match.EndElementIndex).Text.Length)
                            lines[lines.Length - 1] = lines[lines.Length - 1] + texts.ElementAt(match.EndElementIndex).Text.Substring(match.EndCharIndex + 1);

                        txt.Space = new EnumValue<SpaceProcessingModeValues>(SpaceProcessingModeValues.Preserve); // in case your value starts/ends with whitespace
                        txt.Text = lines[0];

                        // remove any extra texts.
                        for (int i = t + 1; i <= match.EndElementIndex; i++) {
                            texts.ElementAt(i).Text = string.Empty; // clear the text
                        }

                        // if 'with' contained line breaks we need to add breaks back...
                        if (lines.Count() > 1) {
                            OpenXmlElement currEl = txt;
                            Break br;

                            // append more lines
                            var run = txt.Parent as Run;
                            for (int i = 1; i < lines.Count(); i++) {
                                br = new Break();
                                run.InsertAfter<Break>(br, currEl);
                                currEl = br;
                                txt = new Text(lines[i]);
                                run.InsertAfter<Text>(txt, currEl);
                                t++; // skip to this next text element
                                currEl = txt;
                            }
                            c = skip; // new line
                        } else {   // continue to process same line
                            c += skip;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Determine if the texts (starting at element t, char c) exactly contain the find text
        /// </summary>
        /// <param name="texts"></param>
        /// <param name="t"></param>
        /// <param name="c"></param>
        /// <param name="find"></param>
        /// <returns>null or the result info</returns>
        private Match IsMatch(IEnumerable<Text> texts, int t, int c, string find)
        {
            int ix = 0;
            for (int i = t; i < texts.Count(); i++) {
                for (int j = c; j < texts.ElementAt(i).Text.Length; j++) {
                    if (find[ix] != texts.ElementAt(i).Text[j]) {
                        return null; // element mismatch
                    }
                    ix++; // match; go to next character
                    if (ix == find.Length)
                        return new Match() { EndElementIndex = i, EndCharIndex = j }; // full match with no issues
                }
                c = 0; // reset char index for next text element
            }
            return null; // ran out of text, not a string match
        }

        /// <summary>
        /// Defines a match result
        /// </summary>
        class Match
        {
            /// <summary>
            /// Last matching element index containing part of the search text
            /// </summary>
            public int EndElementIndex { get; set; }
            /// <summary>
            /// Last matching char index of the search text in last matching element
            /// </summary>
            public int EndCharIndex { get; set; }
        }
        #endregion

    }
}
