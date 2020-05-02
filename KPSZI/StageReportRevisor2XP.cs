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

namespace KPSZI
{
    class StageReportRevisor2XP : StageReport
    {        
        protected override ImageList imageListForTabPage { get; set; }
        protected override Button BtnExportToWord { get => mf.btnExportToWord_2XP; }
        protected override Encoding htmlEncoding { get => Encoding.GetEncoding(1251); }

        public StageReportRevisor2XP(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS, string template, WebBrowser wb)
            : base(stageTab, stageNode, mainForm, IS, template, wb)
        {

        }
        
        protected override void initTabPage()
        {
        }

        public override void ReportToWord(string pathHTML, string nameWord, bool groupExport = false, Word.Document doc = null, Word.Application app = null, Word.Paragraph paragraph = null)
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.Load(pathHTML);
            if (!groupExport)
            {
                app = new Word.Application();
                doc = app.Documents.Add(Environment.CurrentDirectory + "/" + nameWord);
                paragraph = doc.Paragraphs.Add();
            }
            app.Visible = true;

            int countColumn = 0;

            #region Парсинг ключевых HTML элементов
            HtmlNodeCollection titleNodes = htmlDoc.DocumentNode.SelectNodes("//p");
            string title = titleNodes[2].InnerText;
            string[] titleElems = title.Replace("\r\n", "$").Split('$');
            HtmlNodeCollection reportNodes = htmlDoc.DocumentNode.SelectNodes("//tr[contains(@bgcolor, '#ffffff')]/td");
            HtmlNodeCollection table_head_Nodes = htmlDoc.DocumentNode.SelectNodes("//tr[@class=\"tdheader\"]/td");

            HtmlTableElement[] header = new HtmlTableElement[table_head_Nodes.Count];
            for (int i = 0; i < header.Length; i++)
            {
                string text = Regex.Replace(table_head_Nodes[i].InnerText, @"<[^>]+>|&nbsp;", "").Trim();
                header[i] = new HtmlTableElement(text, "000080", "FFFFFF", 1);
            }

            HtmlTableElement[,] data = new HtmlTableElement[reportNodes.Count / 11, 11];

            int k = 0;
            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    string text = Regex.Replace(reportNodes[k].InnerText, @"<[^>]+>|&nbsp;", "").Trim();
                    data[i, j] = new HtmlTableElement(text);
                    k++;
                }
            }

            /*string[,] data = new string[reportNodes.Count / 11, 11];
            int k = 0;
            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    data[i, j] = Regex.Replace(reportNodes[k].InnerText, @"<[^>]+>|&nbsp;", "").Trim();
                    k++;
                }
            }*/
            #endregion

            #region Заполнение шапки
            FillRangeInWord(paragraph.Range, "Приложение Г", "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdColor.wdColorBlack);
            paragraph.Range.InsertParagraphAfter();
            FillRangeInWord(paragraph.Range, titleNodes[0].InnerText, "Times New Roman", 16, 1, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            paragraph.Range.InsertParagraphAfter();
            FillRangeInWord(paragraph.Range, titleNodes[1].InnerText, "Times New Roman", 16, 1, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            paragraph.Range.InsertParagraphAfter();
            FillRangeInWord(paragraph.Range, titleElems[0], "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack);
            paragraph.Range.InsertParagraphAfter();
            FillRangeInWord(paragraph.Range, titleElems[1], "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack);
            paragraph.Range.InsertParagraphAfter();
            #endregion

            #region Заполнение таблицы
            Word.Table table = CreateStandartTable(paragraph.Range, 2, 11, Word.WdLineStyle.wdLineStyleSingle, Word.WdLineStyle.wdLineStyleSingle, doc);
            countColumn = table.Rows[1].Cells.Count;

            table.Cell(1, 1).Merge(table.Cell(2, 1));
            table.Cell(1, 2).Merge(table.Cell(1, 6));
            table.Cell(1, 3).Merge(table.Cell(1, 7));
            table.Cell(1, 1).Width = 300;
            table.Cell(1, 2).Width = 100;
            table.Cell(1, 3).Width = 100;
            for (int i = 0; i < countColumn - 1; i++)
            {
                table.Cell(2, i + 2).Width = 20;
            }
            FillRangeInWord(table.Cell(1, 1).Range, table_head_Nodes[0].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            FillRangeInWord(table.Cell(1, 2).Range, table_head_Nodes[1].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            FillRangeInWord(table.Cell(1, 3).Range, table_head_Nodes[2].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            for (int i = 0; i < countColumn - 6; i++)
            {
                FillRangeInWord(table.Cell(2, i + 2).Range, table_head_Nodes[i + 3].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                FillRangeInWord(table.Cell(2, i + 2 + 5).Range, table_head_Nodes[i + 3 + 5].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            }

            for (int i = 2; i < data.GetLength(0) + 2; i++)
            {
                table.Rows.Add();
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    FillRangeInWord(table.Cell(i + 1, j + 1).Range, data[i - 2, j].Text, "Times New Roman", 12, data[i - 2, j].Bold, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                    if (j == 0)
                        table.Cell(i + 1, j + 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
            }
            #endregion

        }
    }
}
