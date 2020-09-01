using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ToolGood.Algorithm;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace ToolGood.OutputWord
{
    /// <summary>
    /// DocumentFormat.OpenXml MIT协议，可以商用
    /// </summary>
    public class WordTemplate : AlgorithmEngine
    {
        private readonly static Regex _tempEngine = new Regex("^###([^:：]*)[:：](.*)$");// 定义临时变量
        private readonly static Regex _tempMatch = new Regex("(#[^#]+#)");// 
        private readonly static Regex _simplifyMatch = new Regex(@"(\{[^\{\}]*\})");//简化文本 只读取字段
        private readonly static Regex _rowMatch = new Regex(@"({{(.*?)}})");// 
        private readonly static Regex _imageMatch = new Regex(@"(<%(.*?)%>)");//图片 


        private DataTable _dt;
        private int _idx;
        private List<string> listNames = new List<string>();

        public WordTemplate()
        {
            listNames.Add("\\[i\\]");
        }

        public void Reset()
        {
            _dt = null;
            listNames.Clear();
            listNames.Add("\\[i\\]");
            ClearParameters();
        }
        public void SetData(DataTable dataTable)
        {
            _dt = dataTable;
        }
        public void SetListData(string listName, string jsonData)
        {
            var name = listName.Replace("\\", "\\\\").Replace("[", "\\[").Replace("]", "\\]").Replace("(", "\\(").Replace(")", "\\)");
            listNames.Add("\\b" + name + "\\b");
            AddParameter(listName, Operand.CreateJson(jsonData));
        }
        public void SetImageFile(string name,string file)
        {
            var bytes = File.ReadAllBytes(file);
            AddParameter(name, Convert.ToBase64String(bytes));
        }
        public void SetImage(string name, byte[] fileBytes)
        {
            AddParameter(name, Convert.ToBase64String(fileBytes));
        }
        public void SetImageBase64(string name, string imageBase64)
        {
            AddParameter(name, imageBase64);
        }

        public byte[] BuildTemplate(string fileName)
        {
            var bytes = File.ReadAllBytes(fileName);
            var ms = new MemoryStream();
            ms.Write(bytes, 0, bytes.Length);
            ms.Position = 0;

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(ms, true)) {
                var body = wordDoc.MainDocumentPart.Document.Body;
                ReplaceTable(body);
                ReplaceTemplate(body);
                ReplacePicture(wordDoc.MainDocumentPart, body);
                using (var ms2 = new MemoryStream()) {
                    wordDoc.Clone(ms2);
                    return ms2.ToArray();
                }
            };
            ms.Dispose();
        }

        public void BuildTemplate(string fileName, string newFilePath)
        {
            var bytes = File.ReadAllBytes(fileName);
            var ms = new MemoryStream();
            ms.Write(bytes, 0, bytes.Length);
            ms.Position = 0;
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(ms, true)) {
                var body = wordDoc.MainDocumentPart.Document.Body;
                ReplaceTable(body);
                ReplaceTemplate(body);
                ReplacePicture(wordDoc.MainDocumentPart, body);
                wordDoc.SaveAs(newFilePath);
            };
            ms.Dispose();
        }

        #region ReplaceTable
        private void ReplaceTable(Body body)
        {
            foreach (Table table in body.Descendants<Table>()) {

                foreach (TableRow row in table.Descendants<TableRow>()) {
                    bool isRowData = false;
                    foreach (var paragraph in row.Descendants<Paragraph>()) {
                        var text = paragraph.InnerText.Trim();
                        if (_rowMatch.IsMatch(text)) {
                            isRowData = true;
                            break;
                        }
                    }
                    if (isRowData) {
                        // 防止 list[i].Id 写成  [list][[i]].Id 这种繁杂的方式
                        Regex nameReg = new Regex(string.Join("|", listNames));
                        Dictionary<string, string> tempMatches = new Dictionary<string, string>();
                        foreach (Paragraph ph in row.Descendants<Paragraph>()) {
                            var ms2 = _rowMatch.Matches(ph.InnerText.Trim());
                            foreach (System.Text.RegularExpressions.Match m2 in ms2)
                            {
                                var txt = m2.Groups[1].Value;
                                var eval = txt.Substring(2, txt.Length - 4).Trim();
                                eval = nameReg.Replace(eval, new MatchEvaluator((k) => {
                                    return "[" + k.Value + "]";
                                }));
                                tempMatches[txt] = eval;
                            }
                        }

                        TableRow tpl = row.CloneNode(true) as TableRow;
                        TableRow lastRow = row;
                        TableRow opRow = row;
                        var startIndex = UseExcelIndex ? 1 : 0;
                        _idx = startIndex;

                        while (true) {
                            if (_idx > startIndex) { opRow = tpl.CloneNode(true) as TableRow; }

                            bool isMatch = true;
                            foreach (var m in tempMatches) {
                                string value = this.TryEvaluate(m.Value, null);
                                if (value == null) {
                                    isMatch = false;
                                    break;
                                }
                                foreach (var ph in opRow.Descendants<Paragraph>()) {
                                    ReplaceText(ph, m.Key, value);
                                }
                            }
                            if (isMatch == false) {
                                //当数据为空时，清空数据
                                if (_idx == startIndex) {
                                    foreach (var ph in opRow.Descendants<Paragraph>()) {
                                        ph.RemoveAllChildren();
                                    }
                                }
                                break;
                            }

                            if (_idx > startIndex) { table.InsertAfter(opRow, lastRow); }
                            lastRow = opRow;
                            _idx++;
                        }

                    }
                }
            }
        }

        #endregion

        #region ReplaceTemplate
        private void ReplaceTemplate(Body body)
        {
            _idx = UseExcelIndex ? 1 : 0;
            var tempMatches = new List<string>();
            List<Paragraph> deleteParagraph = new List<Paragraph>();
            foreach (var paragraph in body.Descendants<Paragraph>()) {
                var text = paragraph.InnerText.Trim();
                var ms = _tempEngine.Matches(text);
                foreach (System.Text.RegularExpressions.Match m in ms)
                {
                    var name = m.Groups[1].Value.Trim();
                    var engine = m.Groups[2].Value.Trim();
                    var value = this.TryEvaluate(engine, "");
                    this.AddParameter(name, value);
                    deleteParagraph.Add(paragraph);
                }
        
                var ms2 = _tempMatch.Matches(text);
                foreach (System.Text.RegularExpressions.Match m2 in ms2)
                {
                    tempMatches.Add(m2.Groups[1].Value);
                }
               
                var ms3 = _simplifyMatch.Matches(text);
                foreach (System.Text.RegularExpressions.Match m3 in ms3)
                {
                    tempMatches.Add(m3.Groups[1].Value);
                }
            }
            foreach (var paragraph in deleteParagraph) {
                paragraph.Remove();
            }

            Regex nameReg = new Regex(string.Join("|", listNames));
            foreach (var m in tempMatches) {
                string value;
                if (m.StartsWith("#")) {
                    var eval = m.Trim('#');
                    // 防止 list[i].Id 写成  [list][[i]].Id 这种繁杂的方式
                    eval = nameReg.Replace(eval, new MatchEvaluator((k) => {
                        return "[" + k.Value + "]";
                    }));
                    value = this.TryEvaluate(eval, "");
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
            if (parameter.Equals("i", StringComparison.OrdinalIgnoreCase)) {
                return _idx;
            }
            if (_dt != null && _dt.Rows.Count > 0 && _dt.Columns.Contains(parameter)) {
                if (_dt.Rows[0].IsNull(parameter)) {
                    return Operand.CreateNull();
                }
                var obj = _dt.Rows[0][parameter];
                { if (obj is Byte val) { return (int)val; } }
                { if (obj is SByte val) { return (int)val; } }
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
            var p = base.GetParameter(parameter);
            return p;
        }


        #endregion

        #region ReplacePicture
        private void ReplacePicture(MainDocumentPart mainPart, Body body)
        {
            foreach (var paragraph in body.Descendants<Paragraph>()) {
                var text = paragraph.InnerText.Trim();
                var m = _imageMatch.Match(text);
                if (m.Success) {
                    //var name = m.Groups[1].Value.Trim();
                    var engine = m.Groups[2].Value.Trim();
                    var value = this.TryEvaluate("[" + engine + "]", "");
                    paragraph.RemoveAllChildren();
                    if (string.IsNullOrEmpty(value) == false) {
                        InsertPicture(mainPart, paragraph, value);
                    }
                }
            }
        }

        private void InsertPicture(MainDocumentPart mainPart, Paragraph paragraph, string fileBase64)
        {
            var bytes = Convert.FromBase64String(fileBase64);
            ImagePartType type = CheckImageType(bytes);
            ImagePart imagePart = mainPart.AddImagePart(type);

            using (MemoryStream stream = new MemoryStream(bytes)) {
                imagePart.FeedData(stream);
            }
            var element = CreateDrawing(mainPart.GetIdOfPart(imagePart));

            var run = new Run(element);
            paragraph.AppendChild(run);
        }

        private ImagePartType CheckImageType(byte[] buf)
        {
            if (buf == null || buf.Length < 2) {
                return ImagePartType.Jpeg;
            }
            Dictionary<int, ImagePartType> _imageTag = new Dictionary<int, ImagePartType>();
            _imageTag.Add((int)ImageType.BMP, ImagePartType.Bmp);
            _imageTag.Add((int)ImageType.JPG, ImagePartType.Jpeg);
            _imageTag.Add((int)ImageType.GIF, ImagePartType.Gif);
            _imageTag.Add((int)ImageType.PCX, ImagePartType.Pcx);
            _imageTag.Add((int)ImageType.PNG, ImagePartType.Png);
            _imageTag.Add((int)ImageType.TIFF, ImagePartType.Tiff);

            int key = (buf[1] << 8) + buf[0];
            ImagePartType s;
            if (_imageTag.TryGetValue(key, out s)) {
                return s;
            }
            return ImagePartType.Jpeg;
        }
        public enum ImageType
        {
            None = 0,
            BMP = 0x4D42,
            JPG = 0xD8FF,
            GIF = 0x4947,
            PCX = 0x050A,
            PNG = 0x5089,
            PSD = 0x4238,
            RAS = 0xA659,
            SGI = 0xDA01,
            TIFF = 0x4949
        }

        private Drawing CreateDrawing(string relationshipId)
        {
            var element = new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = 990000L, Cy = 792000L },
                         new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                         new DW.DocProperties() { Id = (UInt32Value)1U, Name = "Picture" },
                         new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Image.jpg" },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" })) {
                                             Embed = relationshipId,
                                             CompressionState = A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(new A.Offset() { X = 0L, Y = 0L }, new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new A.PresetGeometry(new A.AdjustValueList()) {
                                             Preset = A.ShapeTypeValues.Rectangle
                                         }))
                             ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     ) { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, });
            return element;
        }
        #endregion

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
