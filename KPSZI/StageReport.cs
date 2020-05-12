using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing;

namespace KPSZI
{
    abstract class StageReport : Stage
    {
        protected string templatePath;
        protected string sourcePath;

        protected WebBrowser wbReport;

        protected Word.Document doc = null;
        protected Word.Application app;
        protected override ImageList imageListForTabPage { get; set; }
        protected abstract Button BtnExportToWord { get; }
        protected abstract Encoding htmlEncoding { get; }

        public StageReport(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS, string template, WebBrowser wb)
            : base(stageTab, stageNode, mainForm, IS)
        {
            templatePath = template;
            wbReport = wb;
        }

        public override void saveChanges()
        {

        }

        public override void enterTabPage()
        {
            wbReport.DocumentText = string.Empty;
            if (SetSourcePath())
            {
                FillData(wbReport, sourcePath, htmlEncoding);
                BtnExportToWord.Visible = true;
                BtnExportToWord.Click += new EventHandler(ExportToWord);
            }
        }

        protected override void initTabPage()
        {
            
        }

        protected void FillData(WebBrowser wb, string path, Encoding encoding)
        {
            StreamReader sr = new StreamReader(path, encoding);
            string htmlText = sr.ReadToEnd();
            sr.Close();
            wb.DocumentText = htmlText;
        }

        bool SetSourcePath()
        {
            if (mf.ofdReport.ShowDialog() == DialogResult.Cancel)
                return false;
            sourcePath = mf.ofdReport.FileName;
            return true;
        }

        public abstract void ReportToWord(string nameWord, bool groupExport = false, Word.Document doc = null, Word.Application app = null, Word.Paragraph paragraph = null);

        public abstract void Parce(string pathHTML);

        protected void ExportToWord(object sender, EventArgs e)
        {
            ReportToWord(templatePath);
        }

        protected void FillRangeInWord(Word.Range range, string text, string fontFamily, int fontSize, byte bold, Word.WdParagraphAlignment alignment, Word.WdColor color, bool isTable = false, Word.WdColor tableBackColor = Word.WdColor.wdColorWhite)//Добавление текста с форматирование в указанный Range
        {
            range.Text = text;
            range.Font.Name = fontFamily;
            range.Font.Size = fontSize;
            range.Font.Bold = bold;
            range.Paragraphs.Alignment = alignment;
            range.Font.Color = color;
            if (isTable)
                range.Shading.ForegroundPatternColor = tableBackColor;

        }

        protected Word.Table CreateStandartTable(Word.Range range, int countRow, int countColumn, Word.WdLineStyle outsideLineStyle, Word.WdLineStyle insideLineStyle, Word.Document doc)//Создание обычной таблицы
        {
            Word.Table table = doc.Tables.Add(range, countRow, countColumn, Type.Missing, Type.Missing);
            table.Borders.OutsideLineStyle = outsideLineStyle;
            table.Borders.InsideLineStyle = insideLineStyle;
            return table;
        }

        protected class HtmlTableElement
        {
            public string Text { get; set; }
            public Word.WdColor wdBackColor { get; set; }
            public Word.WdColor wdForeColor { get; set; }
            private Color BackColor { get; set; }
            private Color ForeColor { get; set; }
            public byte Bold { get; set; }

            public HtmlTableElement(string t, string bColor = "FFFFFF", string fColor = "000000", byte bold = 0)
            {
                Text = t;
                BackColor = GetColor(bColor);
                ForeColor = GetColor(fColor);
                Bold = bold;
                wdBackColor = (Word.WdColor)(BackColor.R + 0x100 * BackColor.G + 0x10000 * BackColor.B);
                wdForeColor = (Word.WdColor)(ForeColor.R + 0x100 * ForeColor.G + 0x10000 * ForeColor.B);
            }

            Color GetColor(string color)
            {
                string red = color.Substring(0, 2);
                string green = color.Substring(2, 2);
                string blue = color.Substring(4, 2);

                int rCode = int.Parse(red, System.Globalization.NumberStyles.HexNumber);
                int gCode = int.Parse(green, System.Globalization.NumberStyles.HexNumber);
                int bCode = int.Parse(blue, System.Globalization.NumberStyles.HexNumber);

                return Color.FromArgb(rCode, gCode, bCode);
            }

        }

        /*string color = "FFEDEC";..

        string red = color.Substring(0, 2);
        string green = color.Substring(2, 2);
        string blue = color.Substring(4, 2);

        int rCode = int.Parse(red, System.Globalization.NumberStyles.HexNumber);
        int gCode = int.Parse(green, System.Globalization.NumberStyles.HexNumber);
        int bCode = int.Parse(blue, System.Globalization.NumberStyles.HexNumber);

        BackColor = Color.FromArgb(rCode, gCode, bCode);*/
    }
}
