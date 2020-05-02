using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System.IO;
using System.Drawing;
using Word = Microsoft.Office.Interop.Word;
namespace KPSZI
{
    class StageReportScannerVS : StageReport
    {
        protected override ImageList imageListForTabPage { get; set; }
        protected override Button BtnExportToWord { get => mf.btnExportToWord_ScannerVS; }
        protected override Encoding htmlEncoding { get => Encoding.UTF8; }

        public StageReportScannerVS(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS, string template, WebBrowser wb)
            : base(stageTab, stageNode, mainForm, IS, template, wb)
        {

        }

        protected override void initTabPage()
        {

        }

        public override void ReportToWord(string pathHTML, string nameWord, bool groupExport = false, Word.Document doc = null, Word.Application app = null, Word.Paragraph paragraph = null)
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.Load(pathHTML, htmlEncoding);
            /*if (!groupExport)
            {
                app = new Word.Application();
                doc = app.Documents.Add(Environment.CurrentDirectory + "/" + nameWord);
                paragraph = doc.Paragraphs.Add();
            }
            app.Visible = true;*/
            int indexP = 0;
            int countIP = 0;

            #region Парсинг ключевых HTML элементов
            HtmlNodeCollection p_Nodes = htmlDoc.DocumentNode.SelectNodes("//p");
            HtmlNodeCollection h1_Nodes = htmlDoc.DocumentNode.SelectNodes("//h1");
            HtmlNodeCollection h2_Nodes = htmlDoc.DocumentNode.SelectNodes("//h2");
            HtmlNodeCollection[,] tables_Nodes = new HtmlNodeCollection[4, 3];

            tables_Nodes[0, 0] = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/div[1]/table/caption");
            tables_Nodes[0, 1] = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/div[1]/table/thead/tr/td");
            tables_Nodes[0, 2] = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/div[1]/table/tbody/tr/td");

            tables_Nodes[1, 0] = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/div[4]/table/caption");
            tables_Nodes[1, 1] = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/div[4]/table/thead/tr/td");
            tables_Nodes[1, 2] = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/div[4]/table/tbody/tr/td");

            tables_Nodes[2, 0] = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/table[1]/caption");
            tables_Nodes[2, 1] = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/table[1]/thead/tr/th");
            tables_Nodes[2, 2] = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/table[1]/tbody/tr/th");

            tables_Nodes[3, 0] = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/table[2]/caption");
            tables_Nodes[3, 1] = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/table[2]/thead/tr/th");
            tables_Nodes[3, 2] = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/table[2]/tbody/tr/th");

            HtmlNodeCollection imgNodes = htmlDoc.DocumentNode.SelectNodes("//img");
            Image[] imgData = new Image[imgNodes.Count];

            for (int i = 0; i < imgData.Length; i++)
            {
                string base64String = imgNodes[i].GetAttributeValue("src", string.Empty).Split(',')[1];
                byte[] data = Convert.FromBase64String(base64String);

                using (var stream = new MemoryStream(data, 0, data.Length))
                {
                    imgData[i] = Image.FromStream(stream);
                }
            }

            #endregion

            #region Раздел "Резюме для руководителей"
            FillRangeInWord(paragraph.Range, "Приложение Б", "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdColor.wdColorBlack);
            paragraph.Range.InsertParagraphAfter();
            FillRangeInWord(paragraph.Range, "Резюме для руководителя", "Times New Roman", 24, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack);
            paragraph.Range.InsertParagraphAfter();
            for (int i = 0; i < 5; i++)
            {
                FillRangeInWord(paragraph.Range, p_Nodes[indexP++].InnerText, "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack);
                paragraph.Range.InsertParagraphAfter();
            }
            paragraph.Range.InsertParagraphAfter();
            FillRangeInWord(paragraph.Range, tables_Nodes[0, 0][0].InnerText, "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorGray85);
            paragraph.Range.InsertParagraphAfter();
            Word.Table table = CreateStandartTable(paragraph.Range, 2, 5, Word.WdLineStyle.wdLineStyleSingle, Word.WdLineStyle.wdLineStyleSingle, doc);
            int count = table.Rows[1].Cells.Count;
            for (int i = 3; i < count + 1; i++)
            {
                table.Rows[1].Cells[2].Merge(table.Rows[1].Cells[3]);
            }
            table.Rows[1].Cells[1].Merge(table.Rows[2].Cells[1]);
            FillRangeInWord(table.Cell(1, 1).Range, tables_Nodes[0, 1][0].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            FillRangeInWord(table.Cell(1, 2).Range, tables_Nodes[0, 1][1].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            FillRangeInWord(table.Cell(2, 2).Range, tables_Nodes[0, 1][2].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            FillRangeInWord(table.Cell(2, 3).Range, tables_Nodes[0, 1][3].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            FillRangeInWord(table.Cell(2, 4).Range, tables_Nodes[0, 1][4].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            FillRangeInWord(table.Cell(2, 5).Range, tables_Nodes[0, 1][5].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);

            for (int i = 0; i < tables_Nodes[0, 2].Count / 5; i++)
            {
                table.Rows.Add();
                countIP++;
                for (int j = 0; j < count; j++)
                {
                    FillRangeInWord(table.Cell(i + 3, j + 1).Range, tables_Nodes[0, 2][count * i + j].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                }

            }
            //Тут долдны быть картинки
            paragraph.Range.InsertParagraphAfter();
            #endregion

            #region Раздел "Границы проекта"
            FillRangeInWord(paragraph.Range, h1_Nodes[1].InnerText, "Times New Roman", 16, 1, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            paragraph.Range.InsertParagraphAfter();
            table = CreateStandartTable(paragraph.Range, 3, 7, Word.WdLineStyle.wdLineStyleSingle, Word.WdLineStyle.wdLineStyleSingle, doc);
            count = table.Rows[1].Cells.Count;
            for (int i = 0; i < 2; i++)
            {
                table.Cell(1, 1).Merge(table.Cell(i + 2, 1));
                table.Cell(1, 2).Merge(table.Cell(i + 2, 2));
                table.Cell(1, 3).Merge(table.Cell(i + 2, 3));
            }

            for (int i = 3; i < count - 1; i++)
            {
                table.Cell(1, 4).Merge(table.Cell(1, 5));
            }
            table.Cell(2, 4).Merge(table.Cell(3, 4));
            table.Cell(2, 5).Merge(table.Cell(3, 5));
            table.Cell(2, 6).Merge(table.Cell(2, 7));
            FillRangeInWord(table.Cell(1, 1).Range, tables_Nodes[1, 1][0].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            FillRangeInWord(table.Cell(1, 2).Range, tables_Nodes[1, 1][1].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            FillRangeInWord(table.Cell(1, 3).Range, tables_Nodes[1, 1][2].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            FillRangeInWord(table.Cell(1, 4).Range, tables_Nodes[1, 1][3].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            FillRangeInWord(table.Cell(2, 4).Range, tables_Nodes[1, 1][4].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            FillRangeInWord(table.Cell(2, 5).Range, tables_Nodes[1, 1][5].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            FillRangeInWord(table.Cell(2, 6).Range, tables_Nodes[1, 1][6].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            FillRangeInWord(table.Cell(3, 6).Range, tables_Nodes[1, 1][7].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            FillRangeInWord(table.Cell(3, 7).Range, tables_Nodes[1, 1][8].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
            for (int i = 0; i < tables_Nodes[1, 2].Count / 5; i++)
            {
                table.Rows.Add();
                for (int j = 0; j < count; j++)
                {
                    FillRangeInWord(table.Cell(i + 4, j + 1).Range, tables_Nodes[1, 2][count * i + j].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                }
            }
            paragraph.Range.InsertParagraphAfter();
            FillRangeInWord(paragraph.Range, p_Nodes[indexP++].InnerText, "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack);
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.InsertParagraphAfter();
            #endregion

            #region Вывод всех хостов
            int index = 0;
            for (int i = 0; i < countIP; i++)
            {
                FillRangeInWord(paragraph.Range, (i + 1) + ". " + h1_Nodes[2 + i].InnerText, "Times New Roman", 16, 1, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                paragraph.Range.InsertParagraphAfter();
                table = CreateStandartTable(paragraph.Range, 3, 2, Word.WdLineStyle.wdLineStyleNone, Word.WdLineStyle.wdLineStyleNone, doc);
                for (int j = 1; j <= 3; j++)
                {
                    FillRangeInWord(table.Cell(j, 1).Range, Regex.Replace(htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/div[5]/table/tbody/tr[" + j + "]/td[1]/strong")[0].InnerText, @"<[^>]+>|&nbsp;", ""), "Times New Roman", 12, 1, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                    FillRangeInWord(table.Cell(j, 2).Range, htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/div[5]/table/tbody/tr[" + j + "]/td[2]")[0].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                }
                paragraph.Range.InsertParagraphAfter();
                FillRangeInWord(paragraph.Range, (i + 1) + ".1. " + h2_Nodes[2 * i + index++].InnerText, "Times New Roman", 16, 1, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                paragraph.Range.InsertParagraphAfter();
                table = CreateStandartTable(paragraph.Range, 1, 6, Word.WdLineStyle.wdLineStyleSingle, Word.WdLineStyle.wdLineStyleSingle, doc);
                HtmlNodeCollection table_thead_Nodes = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/div[5]/div[1]/table/thead/tr/td");
                HtmlNodeCollection table_tbody_Nodes = htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/div[5]/div[1]/table/tbody/tr/td");
                count = table.Rows[1].Cells.Count;
                for (int j = 0; j < (table_tbody_Nodes.Count / 6) + 1; j++)
                {
                    if (j != 0)
                        table.Rows.Add();
                    for (int p = 0; p < count; p++)
                    {
                        if (j != 0)
                        {
                            FillRangeInWord(table.Cell(j + 1, p + 1).Range, table_tbody_Nodes[count * (j - 1) + p].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                        }
                        else
                        {
                            FillRangeInWord(table.Cell(j + 1, p + 1).Range, table_thead_Nodes[count * j + p].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                        }
                    }
                }
                paragraph.Range.InsertParagraphAfter();
                FillRangeInWord(paragraph.Range, (i + 1) + ".2. " + h2_Nodes[2 * i + index++].InnerText, "Times New Roman", 16, 1, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                paragraph.Range.InsertParagraphAfter();
                FillRangeInWord(paragraph.Range, htmlDoc.DocumentNode.SelectNodes("/html/body/div/div/div/div/main/div[5]/div[2]/div/h3")[0].InnerText, "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack);
                paragraph.Range.InsertParagraphAfter();
                for (int j = 0; j < 6; j++)
                {
                    FillRangeInWord(paragraph.Range, p_Nodes[indexP++].InnerText, "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack);
                    paragraph.Range.InsertParagraphAfter();
                }
            }
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.InsertParagraphAfter();
            #endregion

            #region CVSS 2.0
            FillRangeInWord(paragraph.Range, tables_Nodes[2, 0][0].InnerText, "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorGray85);
            paragraph.Range.InsertParagraphAfter();
            table = CreateStandartTable(paragraph.Range, 6, 4, Word.WdLineStyle.wdLineStyleSingle, Word.WdLineStyle.wdLineStyleSingle, doc);
            count = table.Rows[1].Cells.Count;
            for (int i = 0; i < tables_Nodes[2, 1].Count / 4; i++)
            {
                for (int j = 0; j < count; j++)
                {
                    FillRangeInWord(table.Cell(i + 1, j + 1).Range, tables_Nodes[2, 1][count * i + j].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                }
            }
            #endregion

            //doc.SaveAs(FileName: Environment.CurrentDirectory + "\\ReportScannerVS.docx");
            //doc.Close();
            //app.Quit();
            //app = null;
            //doc = null;
        }
    }
}
