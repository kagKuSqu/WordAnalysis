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
using Word = Microsoft.Office.Interop.Word;

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
            Word.Application app = new Word.Application();
            Word.Document doc = null;
            app.Visible = true;
            string path = "D:/MyConfiguration/lzy13870/Documents/test.docx";
            var document = Open(path, app);
            StringBuilder buf = new StringBuilder();
            string xml = document.Content.XML;
            buf.AppendLine(Str(Xml(xml).InnerXml));

            //buf.AppendLine(Str(xml));
            //buf.AppendLine(
            //    Str("Sentences", Str(Loop(document.Content.Sentences).Select(range => PP(range, 1)).ToList().ToArray())));
            textBox1.Text = buf.ToString();
            document.Close();
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

        public static IEnumerable<Word.Range> Loop(Word.Sentences sentences)
        {
            for (int i = 1; i < sentences.Count; i++)
            {
                yield return sentences[i];
            }
        }

        public static Word.Document Open(string path, Word.Application app)
        {
            object file = path;
            object unknow = Type.Missing;
            Word.Document document = app.Documents.Open(
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
            return Str(Range(0, n).Select(i => text).ToList().ToArray());
        }

        public static string PP(Word.Range range, int n)
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

        public static IEnumerable<Word.Range> Loop(Word.Words words)
        {
            for (int i = 0; i < words.Count; i++)
            {
                yield return words[i];
            }
        }

        public static IEnumerable<Word.Cell> Loop(Word.Cells cells)
        {
            for (int i = 1; i < cells.Count; i++)
            {
                yield return cells[i];
            }
        }

        public static IEnumerable<Word.Bookmark> Loop(Word.Bookmarks bookmarks)
        {
            for (int j = 1; j < bookmarks.Count; j++)
            {
                yield return bookmarks[j];
            }
        }

        public static IEnumerable<Word.Table> Loop(Word.Tables tables)
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

        public static string DrawLine(int current, Word.Cell cell)
        {
            return Str(Repeat("\t", current), PP(cell, current));
        }

        public static string PP(Word.Cell cell, int n)
        {
            int current = n + 1;
            StringBuilder buf = new StringBuilder();
            buf.Append(Str("Cell",Repeat("\t", current), PP(cell.Range, current)));
            return buf.ToString();
        }

        public static string PP(Word.Bookmark bookmark, int n)
        {
            int current = n + 1;
            StringBuilder buf = new StringBuilder();
            buf.AppendLine(Str("Bookmark", Repeat("\t", current), bookmark.Name, PP(bookmark.Range, current)));
            return buf.ToString();
        }

        public static string PP(Word.Table table, int n)
        {
            int current = n + 1;
            return Str(Loop(table.Rows).Select(r => Str("Row", Repeat("\t", current), PP(r, current))).ToList().ToArray());
        }

        public static IEnumerable<Word.Row> Loop(Word.Rows rows)
        {
            for (int k = 1; k < rows.Count; k++)
            {
                yield return rows[k];
            }
        }

        public static string PP(Word.Row row, int n)
        {
            int current = n + 1;
            StringBuilder buf = new StringBuilder();
            if (row.Cells.Count > 0)
            {
                for (int l = 1; l < row.Cells.Count; l++)
                {
                    try
                    {
                        Word.Cell cell = row.Cells[l];
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
    }
}