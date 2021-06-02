using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Xml;
using clojure.lang;
using Newtonsoft.Json;
using Microsoft.Scripting.Utils;
using System.Diagnostics;

namespace WordAnalysis
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public void Form1_Load(object sender, EventArgs e)
        {
            //Word.Application app = new Word.Application();
            //Word.Document doc = null;
            //app.Visible = true;
            //string path = "D:/MyConfiguration/lzy13870/Documents/test.docx";
            //var document = Open(path, app);
            //StringBuilder buf = new StringBuilder();
            //string xml = document.Content.XML;
            //buf.AppendLine(Str(Xml(xml).InnerXml));
            ////buf.AppendLine(Str(xml));
            ////buf.AppendLine(
            ////    Str("Sentences", Str(Loop(document.Content.Sentences).Select(range => PP(range, 1)).ToList().ToArray())));
            //textBox1.Text = buf.ToString();
            //document.Close();
            //textBox1.Text = WordDocument("D:\\MyConfiguration\\lzy13870\\Desktop\\sent\\桐庐2日.docx");
        }

        public static string WordDocument(string path)
        {
            if (!File.Exists(path))
            {
                return string.Empty;
            }
            //Process.Start("taskkill", " /f /t /im WINWORD.EXE");
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document document = null;
            app.Visible = true;
            document = Open(path, app);
            StringBuilder buf = new StringBuilder();
            object oMissing = System.Reflection.Missing.Value;

            var xmlText =
                document.Content.XML.ToString();
            document.Close();
            app.Quit();

            //File.WriteAllText("temp.xml", xmlText);
            //Process.Start("temp.xml");
            //buf.AppendLine(xml.SelectNodes("//tbl").ToEnumerable().Take(1).FirstOrDefault().OuterXml);
            buf.AppendLine(Parse(xmlText));
            Clipboard.SetText(buf.ToString());
            return buf.ToString();
        }

        private static string Parse(string xmlText)
        {
            var xml = Xml(
                xmlText.Replace("w:", string.Empty).Replace("wx:", string.Empty).Replace("wsp:", string.Empty));
            StringBuilder buf = new StringBuilder();
            buf.AppendLine(
                xml.Select("/wordDocument/body/sect/sub-section/sub-section/p[1]/r/t")
                    .Select(p => p.InnerText)
                    .JoinStrings());
            buf.AppendLine();
            buf.AppendLine(
                xml.Select("//tr")
                    .Select(
                        tr =>
                        tr.Select("tc/p")
                            .Select(p => p.Select("r/t").Select(t => t.InnerText).JoinStrings())
                            .JoinStrings("|"))
                    .JoinStrings("\r\n"));
            var s = buf.ToString();
            return Str(s, "\r\n", s.Split("\r\n".ToCharArray()).Select(
                (string line) =>
                {
                    line = line.ReplaceRegex("服务标准");
                    var resourceTypePattern = "(门票|用餐|住宿|导游|保险|其他|交通){1}";
                    var unitPattern = "((元|辆|人|餐|间|车|天)+\/*)+";
                    var numberPattern = @"([0-9]+[\,\.\+\-\*\/\^\=\s]*)+";
                    var resourceTypeColumnPattern = @"\|*[0-9]{1,3}、*(门票|用餐|住宿|导游|保险|其他|交通){1}\|*";
                    var fastText =
                        line.Matches(resourceTypeColumnPattern)
                            .JoinStrings("\t")
                            .Matches(resourceTypePattern)
                            .JoinStrings("\t");
                    if (fastText.IsNotEmpty())
                    {
                        var separator = @"(\||\s)+";
                        var separator2 = @"(\||\s)*";

                        var row = line.Split(new []{'|',' '}, StringSplitOptions.RemoveEmptyEntries).ToList();
                        //var resourceNameColumnPattern = Str(separator2, @"(\s*\d*\s*\.*\；*\(*[\u4e00-\u9fa5]+\)*\s*\d*\s*\s*\d*\s*\.*\；*)+", separator2);
                        //var resourceType = line.MatchesJoinTrim(resourceTypeColumnPattern).MatchesJoin(Str(resourceTypePattern));
                        //line = line.ReplaceRegex(resourceTypeColumnPattern);
                        //var price = line.MatchesJoin(Str(separator, pricePattern,unitPattern, separator2), " , ").MatchesJoinTrim(pricePattern);
                        //line = line.ReplaceRegex(Str(separator, pricePattern, unitPattern, separator2));
                        //var count = line.MatchesJoinTrim(Str(separator, pricePattern, separator)).MatchesJoinTrim(pricePattern);
                        //var total = line.MatchesJoinTrim(Str(separator, pricePattern, separator2)).MatchesJoinTrim(pricePattern);
                        //var unit = line.MatchesJoinTrim(Str(separator, unitPattern, separator));
                        //line = line.ReplaceRegex(Str(separator, unitPattern, separator));
                        //var resourceName = line.Matches(resourceNameColumnPattern).Select(txt => txt.Trim().Trim('|')).JoinStrings("+");
                        //line = line.ReplaceRegex(resourceNameColumnPattern);
                        return string.Join(
                            " ",
                            line,
                            Json(row),
                            "\r\n",
                            "resourceType:",
                            row.Get(0).MatchesJoinTrim(resourceTypePattern),
                            "resourceName:",
                            row.Get(1),
                            "price:",
                            row.Get(2).MatchesJoinTrim(numberPattern),
                            "unit:",
                            row.Get(3).MatchesJoinTrim(unitPattern),
                            "count:",
                            row.Get(4).MatchesJoinTrim(numberPattern),
                            "total:",
                            row.Get(5).MatchesJoinTrim(numberPattern));
                    }
                    else if (line.MatchesJoinTrim(@"\|*(D|第)*[0-9]+天*：*\|*").IsNotEmpty())
                    {
                        var dayPattern = @"\|*\s*(D|第)+\s*[0-9]+\s*天*\s*(\：|\:)*\s*\|*";
                        var timePattern = @"(([0-9]{1,2}\s*(\:|\：)\s*[0-9]{1,2})+(\s*(\-|\—)+\s*[0-9]{1,2}\s*(\:|\：)+\s*[0-9]{1,2})*)+";
                        return
                            string.Join(
                                "\r\n",
                                line.MatchesSplit(dayPattern)
                                    .Select(
                                        day =>
                                        Str(
                                            "\r\n",
                                            day.MatchesJoin(@"([\u4e00-\u9fa5]+(\s*(\-|\—)+\s*[\u4e00-\u9fa5]+)+\s*\|+)+", ","),
                                            "\r\n",
                                            day.MatchesJoin(dayPattern, ","),
                                            "\r\n",
                                            day
                                            .MatchesSplit(timePattern)
                                            .Select(
                                                time =>
                                                Str(
                                                    "\r\n\t",
                                                    time.MatchesJoin(timePattern, ","),
                                                    @"    ",
                                                    ParseResourceType(time),
                                                    @"    ",
                                                    time))
                                            .JoinStrings("\r\n")))
                                    .JoinStrings("\r\n")).Replace("|", string.Empty);
                    }
                    return fastText;
                }).JoinStrings("\r\n"));
        }

        private static string ParseResourceType(string time)
        {
            var map = new Dictionary<string, string>()
                          {
                              { "(出发|返回|乘|指定地点)+", "交通" },
                              { "(游览|游玩)+", "景点" },
                              { "(酒店)+", "酒店" },
                              { "(餐)+", "餐饮" },
                          };
            return
                map.Keys.Select(key => time.MatchesJoinTrim(key).IsNotEmpty() ? map[key] : string.Empty)
                    .Where(p => p.IsNotEmpty())
                    .Take(1)
                    .JoinStrings();
        }

        public static XmlElement Xml(string xml)
        {
            try
            {
                XmlDocument document = new XmlDocument();
                document.LoadXml(xml);
                return document.DocumentElement;
            }
            catch (Exception exception)
            {
                Exceptions.Add(exception);
            }
            return null;
        }

        public static IEnumerable<Microsoft.Office.Interop.Word.Range> Loop(Microsoft.Office.Interop.Word.Sentences sentences)
        {
            for (int i = 1; i < sentences.Count; i++)
            {
                yield return sentences[i];
            }
        }

        public static Microsoft.Office.Interop.Word.Document Open(string path, Microsoft.Office.Interop.Word.Application app)
        {
            object file = path;
            object unknow = Type.Missing;
            Microsoft.Office.Interop.Word.Document document = app.Documents.Open(
                ref file,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow);
            return document;
        }

        public static IEnumerable<int> Range(int start, int end)
        {
            for (int i = start; i <= end; i++)
            {
                yield return i;
            }
        }

        public static string Repeat(string text, int n)
        {
            return Str(Enumerable.Select(Range(0, n), i => text).ToList().ToArray());
        }

        public static string PP(Microsoft.Office.Interop.Word.Range range, int n)
        {
            int current = n + 1;
            StringBuilder buf = new StringBuilder(Str(Repeat("\t", n), "Range", range.Text));

            buf.AppendLine(Str(Repeat("\t", current), "Tables",
                Str(
                    Try(
                        range,
                        r =>
                        Loop(r.Tables)
                            .Select(table => Str(PP(table, current)))
                            .ToList()
                            .ToArray()))));
            buf.AppendLine(Str(Repeat("\t", current), "Bookmarks",
                Str(
                    Loop(range.Bookmarks)
                        .Select(bookmark => Str(PP(bookmark, current)))
                        .ToList()
                        .ToArray())));

            buf.AppendLine(Str(Repeat("\t", current), "Words",
                Str(
                    Try(
                        range,
                        r =>
                        Loop(r.Words)
                            .Select(bookmark => Str(PP(bookmark, current)))
                            .ToList()
                            .ToArray()))));
            range.Select();
            buf.AppendLine(Str(Repeat("\t", current), "Cells",
                Str(
                    Try(
                        range,
                        r =>
                        Loop(r.Cells)
                            .Select(bookmark => Str(PP(bookmark, current)))
                            .ToList()
                            .ToArray()))));
            return buf.ToString();
        }

        public static IEnumerable<Microsoft.Office.Interop.Word.Range> Loop(Microsoft.Office.Interop.Word.Words words)
        {
            for (int i = 0; i < words.Count; i++)
            {
                yield return words[i];
            }
        }

        public static IEnumerable<Microsoft.Office.Interop.Word.Cell> Loop(Microsoft.Office.Interop.Word.Cells cells)
        {
            for (int i = 1; i < cells.Count; i++)
            {
                yield return cells[i];
            }
        }

        public static IEnumerable<Microsoft.Office.Interop.Word.Bookmark> Loop(Microsoft.Office.Interop.Word.Bookmarks bookmarks)
        {
            for (int j = 1; j < bookmarks.Count; j++)
            {
                yield return bookmarks[j];
            }
        }

        public static IEnumerable<Microsoft.Office.Interop.Word.Table> Loop(Microsoft.Office.Interop.Word.Tables tables)
        {
            for (int j = 1; j < tables.Count; j++)
            {
                yield return tables[j];
            }
        }

        public static List<Exception> Exceptions = new List<Exception>();

        public static TOut Try<TIn, TOut>(TIn source, Func<TIn, TOut> fn) where TOut : class where TIn : class
        {
            try
            {
                return fn(source);
            }
            catch (Exception exception)
            {
                Exceptions.Add(exception);
            }

            return null;
        }

        public static string DrawLine(int current, Microsoft.Office.Interop.Word.Cell cell)
        {
            return Str(Repeat("\t", current), PP(cell, current));
        }

        public static string PP(Microsoft.Office.Interop.Word.Cell cell, int n)
        {
            int current = n + 1;
            StringBuilder buf = new StringBuilder();
            buf.Append(Str("Cell", Repeat("\t", current), PP(cell.Range, current)));
            return buf.ToString();
        }

        public static string PP(Microsoft.Office.Interop.Word.Bookmark bookmark, int n)
        {
            int current = n + 1;
            StringBuilder buf = new StringBuilder();
            buf.AppendLine(Str("Bookmark", Repeat("\t", current), bookmark.Name, PP(bookmark.Range, current)));
            return buf.ToString();
        }

        public static string PP(Microsoft.Office.Interop.Word.Table table, int n)
        {
            int current = n + 1;
            return Str(Loop(table.Rows).Select(r => Str("Row", Repeat("\t", current), PP(r, current))).ToList().ToArray());
        }

        public static IEnumerable<Microsoft.Office.Interop.Word.Row> Loop(Microsoft.Office.Interop.Word.Rows rows)
        {
            for (int k = 1; k < rows.Count; k++)
            {
                yield return rows[k];
            }
        }

        public static string PP(Microsoft.Office.Interop.Word.Row row, int n)
        {
            int current = n + 1;
            StringBuilder buf = new StringBuilder();
            if (row.Cells.Count > 0)
            {
                for (int l = 1; l < row.Cells.Count; l++)
                {
                    try
                    {
                        Microsoft.Office.Interop.Word.Cell cell = row.Cells[l];
                        buf.AppendLine(Str(Repeat("\t", current), PP(cell.Range, current)));
                    }
                    catch (Exception exception)
                    {
                        buf.AppendLine(exception.Message);
                        throw exception;
                    }
                }
            }

            return buf.ToString();
        }

        public static string Str(params object[] strings)
        {
            if (strings != null)
            {
                return string.Join(string.Empty, strings);
            }

            return string.Empty;
        }

        public static string Json(object obj)
        {
            return JsonConvert.SerializeObject(obj);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void textBox1_DoubleClick(object sender, EventArgs e)
        {
            var result = this.openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
                var document = Open(this.openFileDialog1.FileName, app);
                var xmlText = document.Content.XML;
                document.Close();
                app.Quit();
                this.textBox1.Text = Parse(xmlText);
            }
        }
    }
}