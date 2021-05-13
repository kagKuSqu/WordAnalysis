using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordAnalysis
{
    using Newtonsoft.Json;

    using Word = Microsoft.Office.Interop.Word;


    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Word.Application app = new Word.Application();
            Word.Document doc = null;
            app.Visible = true;
            string path = "D:/MyConfiguration/lzy13870/Documents/test.docx";
            var document = Open(path, app);
            StringBuilder buf = new StringBuilder(Json(Range(2,7)));

            foreach (Word.Range range in document.Content.Sentences)
            {
                buf.AppendLine(PP(range, 1));
            }
            buf.AppendLine(Json(document.Content.Tables));
            textBox1.Text = buf.ToString();
            document.Close();
        }

        private static Word.Document Open(string path, Word.Application app)
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

        private static string PP(Word.Range range,int n)
        {
            int current = n+1;
            StringBuilder buf = new StringBuilder(Str(Repeat("\t",current), range.Text));
            for (int j = 1; j < range.Tables.Count; j++)
            {
                try
                {
                    buf.Append("\t");
                    Word.Table table = range.Tables[j];
                    buf.AppendLine(Str(Repeat("\t", current), PP(table, current)));
                    buf.AppendLine(Str(Repeat("\t",current),PP(table.Range,current)));
                }
                catch (Exception exception)
                {
                    buf.AppendLine(exception.Message);
                    throw exception;
                }
            }
            for (int j = 1; j < range.Bookmarks.Count; j++)
            {
                try
                {
                    Word.Bookmark bookmark = range.Bookmarks[j];
                    buf.AppendLine(Str(Repeat("\t", current), PP(bookmark, current)));
                }
                catch (Exception exception)
                {
                    buf.AppendLine(exception.Message);
                    throw exception;
                }
            }

            foreach (Word.Cell cell in range.Cells)
            {
                buf.AppendLine(Str(Repeat("\t", current), PP(cell, current)));
            }
            return buf.ToString();
        }

        private static string PP(Word.Cell cell, int n)
        {
            int current = n + 1;
            StringBuilder buf = new StringBuilder();
            buf.Append(Str(Repeat("\t", current), PP(cell.Range, current)));
            return buf.ToString();
        }

        private static string PP(Word.Bookmark bookmark,int n)
        {
            int current = n + 1;
            StringBuilder buf = new StringBuilder();
            buf.AppendLine(Str(Repeat("\t", current), bookmark.Name, PP(bookmark.Range, current)));
            return buf.ToString();
        }

        public static string PP(Word.Table table,int n)
        {
            int current = n + 1;
            StringBuilder buf = new StringBuilder();
            if (table.Rows.Count > 0)
            {
                for (int k = 1; k < table.Rows.Count; k++)
                {
                    try
                    {
                        Word.Row row = table.Rows[k];
                        buf.Append(Str(Repeat("\t", current), PP(row,current)));
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

        private static string PP(Word.Row row,int n)
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
                        buf.AppendLine(Str(Repeat("\t", current), PP(cell.Range,current)));
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
            return string.Join(string.Empty, strings);
        }

        public static string Json(object obj)
        {
            return JsonConvert.SerializeObject(obj);
        }
    }
}
