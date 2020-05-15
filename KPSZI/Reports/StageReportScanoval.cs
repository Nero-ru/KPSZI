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
using DocumentFormat.OpenXml.Wordprocessing;

namespace KPSZI
{
    class StageReportScanoval : StageReport
    {
        protected override Encoding htmlEncoding { get => Encoding.UTF8; }
        protected HtmlTableElement[][][] tables = null;
        protected string caption1 = null;
        protected string caption2 = null;

        public StageReportScanoval(MainForm mainForm, string template)
            : base(mainForm, template)
        {

        }

        public override void Parce(string pathHTML)
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.Load(pathHTML, htmlEncoding);
            try
            {
                #region Парсинг ключевых HTML элементов
                HtmlNodeCollection reportNodes = htmlDoc.DocumentNode.SelectNodes("//table");
                caption1 = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div[3]/div")[0].InnerText;
                caption2 = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div[4]/div")[0].InnerText;
                tables = new HtmlTableElement[reportNodes.Count][][];

                for (int i = 0; i < reportNodes.Count; i++)
                {
                    HtmlAgilityPack.HtmlDocument tempHtmlRows = new HtmlAgilityPack.HtmlDocument();
                    tempHtmlRows.LoadHtml(reportNodes[i].InnerHtml);
                    HtmlNodeCollection tempNodesRows = tempHtmlRows.DocumentNode.SelectNodes("//tr");

                    HtmlTableElement[][] tempTable = new HtmlTableElement[tempNodesRows.Count][];

                    for (int j = 0; j < tempNodesRows.Count; j++)
                    {
                        HtmlAgilityPack.HtmlDocument tempHtmlColumns = new HtmlAgilityPack.HtmlDocument();
                        tempHtmlColumns.LoadHtml(tempNodesRows[j].InnerHtml);
                        HtmlNodeCollection tempNodesColumns;

                        if (i == reportNodes.Count - 1 && j % 8 == 7)
                        {
                            tempNodesColumns = tempHtmlColumns.DocumentNode.SelectNodes("//td//div");
                        }
                        else
                        {
                            tempNodesColumns = tempHtmlColumns.DocumentNode.SelectNodes("//td | //th");
                        }
                        HtmlTableElement[] tempRow = new HtmlTableElement[tempNodesColumns.Count];

                        for (int k = 0; k < tempNodesColumns.Count; k++)
                        {
                            string text = tempNodesColumns[k].InnerText.Trim();
                            string bColor = "FFFFFF";
                            string fColor = "000000";
                            byte bold = 0;

                            if (i == 0)
                            {
                                if (k == 0)
                                {
                                    bColor = "D9D9D9";
                                }
                                else
                                {
                                    bColor = "EFF4FB";
                                }
                            }

                            else if (i == 1)
                            {
                                if (j == 0)
                                {
                                    bold = 1;
                                }
                                else if (j == tempNodesRows.Count - 1)
                                {
                                    bColor = "EFF4FB";
                                }
                                else if (k == 0)
                                {
                                    fColor = "FFFFFF";
                                    switch (j)
                                    {
                                        case 1:
                                            bColor = "89171A";
                                            break;
                                        case 2:
                                            bColor = "CC0000";
                                            break;
                                        case 3:
                                            bColor = "F5770F";
                                            break;
                                        case 4:
                                            bColor = "00705C";
                                            break;
                                    }
                                }
                            }

                            else if (i == 2)
                            {
                                if (j == 0)
                                {
                                    bColor = "D9D9D9";
                                    bold = 1;
                                }
                                else
                                {
                                    if (j % 2 == 1)
                                    {
                                        if (k == 1)
                                        {
                                            fColor = "FFFFFF";

                                            if (text == "Критический")
                                            {
                                                bColor = "89171A";
                                            }
                                            else if (text == "Высокий")
                                            {
                                                bColor = "CC0000";
                                            }
                                            else if (text == "Средний")
                                            {
                                                bColor = "F5770F";
                                            }
                                            else
                                            {
                                                bColor = "00705C";
                                            }
                                        }
                                    }
                                    else
                                    {
                                        fColor = "515151";
                                    }
                                }
                            }

                            else if (i == 3)
                            {
                                if (j % 8 == 0)
                                {
                                    if (k == 0)
                                    {
                                        fColor = "FFFFFF";
                                        if (tempNodesColumns[k + 1].InnerText.Trim() == "Уровень опасности: Критический")
                                        {
                                            bColor = "89171A";
                                        }
                                        else if (tempNodesColumns[k + 1].InnerText.Trim() == "Уровень опасности: Высокий")
                                        {
                                            bColor = "CC0000";
                                        }
                                        else if (tempNodesColumns[k + 1].InnerText.Trim() == "Уровень опасности: Средний")
                                        {
                                            bColor = "F5770F";
                                        }
                                        else
                                        {
                                            bColor = "00705C";
                                        }
                                    }
                                }

                                else if (j % 8 == 1)
                                {
                                    if (k == 0)
                                    {
                                        bColor = "D9D9D9";
                                    }
                                    else
                                    {
                                        bColor = "EFF4FB";
                                        bold = 1;
                                    }
                                }

                                else if (j % 8 == 2 || j % 8 == 4 || j % 8 == 6)
                                {
                                    if (k == 0)
                                    {
                                        bColor = "D9D9D9";
                                    }
                                }

                                else if (j % 8 == 7)
                                {
                                    if (k % 2 == 0)
                                    {
                                        bold = 1;
                                    }
                                }
                            }

                            tempRow[k] = new HtmlTableElement(text, bColor, fColor, bold);
                        }
                        tempTable[j] = tempRow;
                    }
                    tables[i] = tempTable;
                }

                /*
                for (int i = 0; i < reportNodes.Count; i++)
                {
                    for (int j = 0; j < tempNodesRows.Count; j++)
                    {                                                                                
                        for (int k = 0; k < tempNodesColumns.Count; k++)
                        {
                            tempRow[k] = tempNodesColumns[k].InnerText;
                        }

                        tempTable[j] = tempRow;
                    }
                    tables[i] = tempTable;
                }*/

                /*string[][][] tables = new string[reportNodes.Count][][];

                for (int i = 0; i < reportNodes.Count; i++)
                {
                    HtmlAgilityPack.HtmlDocument tempHtmlRows = new HtmlAgilityPack.HtmlDocument();
                    tempHtmlRows.LoadHtml(reportNodes[i].InnerHtml);
                    HtmlNodeCollection tempNodesRows = tempHtmlRows.DocumentNode.SelectNodes("//tr");

                    string[][] tempTable = new string[tempNodesRows.Count][];

                    for (int j = 0; j < tempNodesRows.Count; j++)
                    {
                        HtmlAgilityPack.HtmlDocument tempHtmlColumns = new HtmlAgilityPack.HtmlDocument();
                        tempHtmlColumns.LoadHtml(tempNodesRows[j].InnerHtml);

                        HtmlNodeCollection tempNodesColumns;
                        if (i == reportNodes.Count - 1 && j % 8 == 7)
                        {
                            tempNodesColumns = tempHtmlColumns.DocumentNode.SelectNodes("//td//div");
                        }
                        else
                        {
                            tempNodesColumns = tempHtmlColumns.DocumentNode.SelectNodes("//td | //th");
                        }

                        string[] tempRow = new string[tempNodesColumns.Count];

                        for (int k = 0; k < tempNodesColumns.Count; k++)
                        {
                            tempRow[k] = tempNodesColumns[k].InnerText;
                        }

                        tempTable[j] = tempRow;
                    }
                    tables[i] = tempTable;
                }*/
                #endregion
            }
            catch
            {
                MessageBox.Show("Отчет \"Scanoval\" был выбран неправильно!\nПроверьте выбранные отчеты и повторите попытку", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        public override void ReportToWord(string nameWord, bool groupExport = false, Word.Document doc = null, Word.Application app = null, Word.Paragraph paragraph = null)
        {
            if (tables == null || caption1 == null || caption2 == null)
                MessageBox.Show("Отчет СЗИ \"ScanOval\" был обработан некорректно, убедитесь в правильности выбора отчета", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                if (!groupExport)
                {
                    app = new Word.Application();
                    doc = app.Documents.Add(Environment.CurrentDirectory + "/" + nameWord);
                    paragraph = doc.Paragraphs.Add();
                }
                app.Visible = true;



                FillRangeInWord(paragraph.Range, "Приложение А", "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdColor.wdColorBlack);
                paragraph.Range.InsertParagraphAfter();
                FillRangeInWord(paragraph.Range, "Результаты сканирования программного обеспечения на рабочих станциях и серверах ПО «ScanOval»", "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                paragraph.Range.InsertParagraphAfter();

                #region Таблица №1
                Word.Table table = CreateStandartTable(paragraph.Range, 5, 2, Word.WdLineStyle.wdLineStyleNone, Word.WdLineStyle.wdLineStyleNone, doc);
                table.Spacing = 3;
                for (int i = 0; i < tables[0].GetLength(0); i++)
                {
                    for (int j = 0; j < tables[0][i].GetLength(0); j++)
                    {
                        FillRangeInWord(table.Cell(i + 1, j + 1).Range, tables[0][i][j].Text, "Times New Roman", 12, tables[0][i][j].Bold, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack, true, tables[0][i][j].wdBackColor);
                    }
                }
                paragraph.Range.InsertParagraphAfter();
                #endregion

                #region Таблица №2
                table = CreateStandartTable(paragraph.Range, 6, 3, Word.WdLineStyle.wdLineStyleNone, Word.WdLineStyle.wdLineStyleNone, doc);
                table.Spacing = 3;
                for (int i = 0; i < tables[1].GetLength(0); i++)
                {
                    for (int j = 0; j < tables[1][i].GetLength(0); j++)
                    {
                        FillRangeInWord(table.Cell(i + 1, j + 1).Range, tables[1][i][j].Text, "Times New Roman", 12, tables[1][i][j].Bold, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack, true, tables[1][i][j].wdBackColor);
                    }
                }
                paragraph.Range.InsertParagraphAfter();
                #endregion

                #region Таблица №3
                FillRangeInWord(paragraph.Range, caption1, "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack);
                table.Spacing = 6;
                paragraph.Range.InsertParagraphAfter();
                table = CreateStandartTable(paragraph.Range, Int32.Parse(tables[1][5][1].Text) * 2 + 1, 3, Word.WdLineStyle.wdLineStyleNone, Word.WdLineStyle.wdLineStyleNone, doc);
                table.Columns[1].Width = 100;
                table.Columns[2].Width = 80;
                table.Columns[3].Width = 319;
                for (int i = 0; i < tables[2].GetLength(0); i++)
                {
                    for (int j = 0; j < tables[2][i].GetLength(0); j++)
                    {
                        if ((i - 1) % 2 == 1 && j == 1)
                        {
                            table.Cell(i + 1, j + 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            FillRangeInWord(table.Cell(i + 1, j + 2).Range, tables[2][i][j].Text, "Times New Roman", 12, tables[2][i][j].Bold, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack, true, tables[2][i][j].wdBackColor);
                        }
                        else
                        {
                            if (j == 2)
                            {
                                table.Cell(i + 1, j + 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                                FillRangeInWord(table.Cell(i + 1, j + 1).Range, tables[2][i][j].Text, "Times New Roman", 12, tables[2][i][j].Bold, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack, true, tables[2][i][j].wdBackColor);
                            }
                            else
                            {
                                table.Cell(i + 1, j + 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                                FillRangeInWord(table.Cell(i + 1, j + 1).Range, tables[2][i][j].Text, "Times New Roman", 12, tables[2][i][j].Bold, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack, true, tables[2][i][j].wdBackColor);
                            }
                        }

                    }
                }
                paragraph.Range.InsertParagraphAfter();
                #endregion

                #region Таблица №4
                FillRangeInWord(paragraph.Range, caption2, "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack);
                table = CreateStandartTable(paragraph.Range, 1, 2, Word.WdLineStyle.wdLineStyleNone, Word.WdLineStyle.wdLineStyleNone, doc);
                table.Spacing = 3;
                int indexI = 0;
                bool flag = false;
                for (int i = 0; i < tables[3].GetLength(0); i++)
                {
                    for (int j = 0; j < tables[3][i].GetLength(0); j++)
                    {
                        if (tables[3][i].GetLength(0) == 1)
                        {
                            table.Rows.Add();
                            flag = true;
                            table.Cell(i + 1 + indexI, j + 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            table.Cell(i + 1 + indexI, j + 1).Merge(table.Cell(i + 1 + indexI, j + 2));
                            FillRangeInWord(table.Cell(i + 1 + indexI, j + 1).Range, Regex.Replace(tables[3][i][j].Text, "\r\n", ""), "Times New Roman", 10, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack, true, tables[3][i][j].wdBackColor);
                            continue;
                        }
                        if ((i + 1) % 8 == 0)
                        {
                            table.Cell(i + 1 + indexI, (j % 2) + 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            FillRangeInWord(table.Cell(i + 1 + indexI, (j % 2) + 1).Range, tables[3][i][j].Text, "Times New Roman", 10, tables[3][i][j].Bold, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack, true, tables[3][i][j].wdBackColor);
                            if (j % 2 == 1 && j != tables[3][i].GetLength(0) - 1)
                            {
                                indexI++;
                                table.Rows.Add();
                            }
                            if (j == tables[3][i].GetLength(0) - 1)
                            {
                                table.Rows.Add();
                                table.Rows.Add();
                                indexI += 2;
                            }
                            continue;
                        }
                        if (j == 0)
                        {
                            table.Cell(i + 1 + indexI, j + 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            FillRangeInWord(table.Cell(i + 1 + indexI, j + 1).Range, tables[3][i][j].Text, "Times New Roman", 10, tables[3][i][j].Bold, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack, true, tables[3][i][j].wdBackColor);
                        }
                        else
                        {
                            table.Cell(i + 1 + indexI, j + 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            FillRangeInWord(table.Cell(i + 1 + indexI, j + 1).Range, tables[3][i][j].Text, "Times New Roman", 10, tables[3][i][j].Bold, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack, true, tables[3][i][j].wdBackColor);
                        }
                    }
                    if (flag || i == tables[3].GetLength(0) - 1) flag = false;
                    else
                    {
                        table.Rows.Add();
                    }

                }
                #endregion

                paragraph.Range.InsertBreak(Word.WdBreakType.wdPageBreak);
            }
        }
    }
}
