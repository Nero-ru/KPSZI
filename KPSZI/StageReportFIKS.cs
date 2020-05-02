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
    class StageReportFIKS : StageReport
    {
        protected override ImageList imageListForTabPage { get; set; }
        protected override Button BtnExportToWord { get => mf.btnExportToWord_FIKS; }
        protected override Encoding htmlEncoding { get => Encoding.GetEncoding(1251); }

        public StageReportFIKS(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS, string template, WebBrowser wb)
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
            int countColumn;
            int tableIndex = 0;

            #region Парсинг ключевых HTML элементов
            HtmlNodeCollection reportNodes = htmlDoc.DocumentNode.SelectNodes("//table[1]//tr");
            HtmlNodeCollection centerNodes = htmlDoc.DocumentNode.SelectNodes("//center");

            HtmlTableElement[][] elems = new HtmlTableElement[reportNodes.Count][];

            for (int i = 0; i < elems.Length; i++)
            {
                HtmlAgilityPack.HtmlDocument tempHtml = new HtmlAgilityPack.HtmlDocument();
                tempHtml.LoadHtml(reportNodes[i].InnerHtml);
                HtmlNodeCollection tempNodes = tempHtml.DocumentNode.SelectNodes("//td");
                HtmlTableElement[] tempArray = new HtmlTableElement[tempNodes.Count];

                for (int j = 0; j < tempArray.Length; j++)
                {
                    string text = tempNodes[j].InnerText;
                    string bColor = "FFFFFF";
                    string fColor = "000000";
                    byte bold = 0;

                    if (i % 3 == 0)
                    {
                        if (i == 0)
                        {
                            bColor = "DFDFDF";
                            bold = 1;
                        }

                        else
                        {
                            bColor = "FFEDEC";
                            bold = 1;
                        }
                    }

                    else if (i % 3 == 1)
                    {
                        if (i < elems.Length - 2)
                        {
                            bColor = "F4F4F4";
                        }
                        else
                        {
                            bColor = "FFE0E0";
                            bold = 1;
                        }
                    }
                    tempArray[j] = new HtmlTableElement(text, bColor, fColor, bold);
                }
                elems[i] = tempArray;
            }

            /*string[][] elems = new string[reportNodes.Count][];
            for (int i = 0; i < elems.Length; i++)
            {
                HtmlAgilityPack.HtmlDocument tempHtml = new HtmlAgilityPack.HtmlDocument();
                tempHtml.LoadHtml(reportNodes[i].InnerHtml);
                HtmlNodeCollection tempNodes = tempHtml.DocumentNode.SelectNodes("//td");

                string[] tempArray = new string[tempNodes.Count];

                for (int j = 0; j < tempArray.Length; j++)
                {
                    tempArray[j] = tempNodes[j].InnerText;
                }
                elems[i] = tempArray;
            }*/
            #endregion

            #region Заполнение шапки
            FillRangeInWord(paragraph.Range, "Приложение В", "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdColor.wdColorBlack);
            paragraph.Range.InsertParagraphAfter();
            FillRangeInWord(paragraph.Range, "Результат фиксации файлов СЗИ с помощью программы «ФИКС»", "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            paragraph.Range.InsertParagraphAfter();
            FillRangeInWord(paragraph.Range, "Таблица Б1. Имена файлов в дистрибутиве СЗИ «Secret Net Studio 8» и их контрольные значения.", "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack);
            paragraph.Range.InsertParagraphAfter();
            for (int i = 0; i < 4; i++)
            {
                FillRangeInWord(paragraph.Range, centerNodes[i].InnerText, "Times New Roman", 14, 1, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                paragraph.Range.InsertParagraphAfter();
            }

            #endregion

            #region Заполнение таблицы
            Word.Table table = CreateStandartTable(paragraph.Range, 1, 6, Word.WdLineStyle.wdLineStyleSingle, Word.WdLineStyle.wdLineStyleSingle, doc);
            countColumn = table.Rows[1].Cells.Count;
            for (int i = 0; i < countColumn; i++)
            {
                FillRangeInWord(table.Cell(1, i + 1).Range, elems[tableIndex][i].Text, "Times New Roman", 12, elems[tableIndex][i].Bold, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            }
            tableIndex++;
            for (int i = 0; i < countColumn; i++)
            {
                table.Cell(1, i + 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            }
            for (int i = 0; i < elems.GetLength(0) - 1; i++)
            {
                table.Rows.Add();
            }
            for (int i = 0; i < (elems.GetLength(0) - 1) / 3; i++)
            {
                table.Cell(2 + i * 3, 1).Merge(table.Cell(2 + i * 3, 6));
                FillRangeInWord(table.Cell(2 + i * 3, 1).Range, elems[tableIndex][0].Text, "Times New Roman", 12, elems[tableIndex][0].Bold, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack, true, elems[tableIndex][0].wdBackColor);
                tableIndex++;
                for (int j = 0; j < countColumn; j++)
                {
                    FillRangeInWord(table.Cell(3 + i * 3, j + 1).Range, elems[tableIndex][j].Text, "Times New Roman", 12, elems[tableIndex][j].Bold, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack, true, elems[tableIndex][j].wdBackColor);
                }
                tableIndex++;
                table.Cell(4 + i * 3, 1).Merge(table.Cell(4 + i * 3, 3));
                for (int j = 0; j < countColumn - 2; j++)
                {
                    FillRangeInWord(table.Cell(4 + i * 3, j + 1).Range, elems[tableIndex][j].Text, "Times New Roman", 12, elems[tableIndex][j].Bold, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack, true, elems[tableIndex][j].wdBackColor);
                }
                tableIndex++;
                if (i + 1 == (elems.GetLength(0) - 1) / 3)
                {
                    table.Cell(elems.GetLength(0) - 1, 1).Merge(table.Cell(elems.GetLength(0) - 1, 3));
                    for (int j = 0; j < countColumn - 2; j++)
                    {
                        FillRangeInWord(table.Cell(elems.GetLength(0) - 1, j + 1).Range, elems[tableIndex][j].Text, "Times New Roman", 12, elems[tableIndex][j].Bold, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack, true, elems[tableIndex][j].wdBackColor);
                    }
                    tableIndex++;
                    table.Cell(elems.GetLength(0), 1).Merge(table.Cell(elems.GetLength(0), 6));
                    FillRangeInWord(table.Cell(elems.GetLength(0), 1).Range, elems[tableIndex][0].Text, "Times New Roman", 12, elems[tableIndex][0].Bold, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdColor.wdColorBlack, true, elems[tableIndex][0].wdBackColor);
                }
            }
            #endregion
        }
    }
}
