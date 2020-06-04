﻿using System;
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
using DocumentFormat.OpenXml.Spreadsheet;

namespace KPSZI
{
    class StageReportScannerVS : StageReport
    {
        protected override Encoding htmlEncoding { get => Encoding.UTF8; }
        protected HtmlElement[] tablesCaptions = new HtmlElement[]
        {
            new HtmlElement("Количество найденных уязвимостей приведено в таблице"),
            new HtmlElement("CVSS 2.0"),
            new HtmlElement("CVSS 3.0"),
        };

        protected HtmlElement[] captions = new HtmlElement[] // заголовки
        {
            new HtmlElement("1. Резюме для руководителя"),
            new HtmlElement("2. Границы проекта"),
            new HtmlElement("3. Порты и сервисы"),
            new HtmlElement("4. Уязвимости")
        };

        protected HtmlElement[] projectBordersNote = new HtmlElement[] // Примечание после пункта 2. Границы проекта
        {
            new HtmlElement("Примечание:"),
            new HtmlElement("\"+\" - проверка проводилась на указанном хосте;"),
            new HtmlElement("\"-\" - проверка НЕ проводилась на указанном хосте")
        };

        protected HtmlElement[] summary = null; // резюме для руководителя
        protected object[][] vulnSections = null; // секции в пункте 4. Уязвимости
        protected HtmlTableElement[][][] tables = null; // табличные элементы отчета
        protected Image[] imgData = null; // картинки
        

        public StageReportScannerVS(MainForm mainForm, string template)
            : base(mainForm, template)
        {
            mainForm.btnExportTest.Click += new EventHandler(Test);
        }

        public void Test(object sender, EventArgs e)
        {
            Parce(@"reports/ScannerVS/scaner-vs_report_short_13.04.2020.html");
        }

        public override void Parce(string pathHTML)
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.Load(pathHTML, htmlEncoding);

            #region Парсинг ключевых HTML элементов

            HtmlNodeCollection reportNodes = htmlDoc.DocumentNode.SelectNodes("//table");
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

                    HtmlNodeCollection tempNodesColumns = tempHtmlColumns.DocumentNode.SelectNodes("//td | //th");
                    HtmlTableElement[] tempCell = new HtmlTableElement[tempNodesColumns.Count];

                    for (int k = 0; k < tempNodesColumns.Count; k++)
                    {
                        string text = tempNodesColumns[k].InnerText.Trim();
                        string bColor = "FFFFFF";
                        string fColor = "000000";
                        byte bold = 0;
                        int rowspan = 1;
                        int colspan = 1;

                        if (i > 3)
                        {
                            bold = 1;
                        }

                        if (i == 1)
                        {
                            if (j == 0)
                            {
                                if (k < 3)
                                {
                                    rowspan = 3;
                                }
                                else
                                {
                                    rowspan = 4;
                                }
                            }
                            else if (j == 1)
                            {
                                rowspan = 2;
                            }
                        }
                        else if (i == 4)
                        {
                            if (j == 0)
                            {
                                colspan = 2;
                            }
                            else if (j == 1 || j == 3 || j == 4)
                            {
                                colspan = 4;
                            }
                            else
                            {
                                colspan = 3;
                            }
                        }

                        tempCell[k] = new HtmlTableElement(text, bColor, fColor, bold, rowspan, colspan);
                    }
                    tempTable[j] = tempCell;
                }
                tables[i] = tempTable;
            }

            HtmlNodeCollection summaryNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class='chapter'][1]/p");
            summary = new HtmlElement[summaryNodes.Count];

            for (int i = 0; i < summaryNodes.Count; i++)
            {
                summary[i] = new HtmlElement(summaryNodes[i].InnerText, "000000", 0);
            }

            HtmlNodeCollection vulnSectionNodes = htmlDoc.DocumentNode.SelectNodes("//div[@class='chapter'][4]//div[@class='section']");
            vulnSections = new object[vulnSectionNodes.Count][];

            for (int i = 0; i < vulnSectionNodes.Count; i++)
            {
                HtmlAgilityPack.HtmlDocument vulnHtml = new HtmlAgilityPack.HtmlDocument();
                vulnHtml.LoadHtml(vulnSectionNodes[i].InnerHtml);

                HtmlNodeCollection sectionChilds = vulnHtml.DocumentNode.SelectNodes("/*");
                object[] childs = new object[sectionChilds.Count];

                for (int j = 0; j < sectionChilds.Count; j++)
                {
                    object section = null;

                    string name = sectionChilds[j].Name;

                    if (name == "h2")
                    {
                        section = new HtmlElement(sectionChilds[j].InnerText, "000000", 1);
                    }
                    else if (name == "ol")
                    {
                        HtmlAgilityPack.HtmlDocument olSectionHtml = new HtmlAgilityPack.HtmlDocument();
                        olSectionHtml.LoadHtml(sectionChilds[j].InnerHtml);

                        HtmlNodeCollection olSection = olSectionHtml.DocumentNode.SelectNodes("/li");
                        section = new HtmlElement[olSection.Count];

                        for (int k = 0; k < olSection.Count; k++)
                        {
                            ((HtmlElement[])section)[k] = new HtmlElement(olSection[k].InnerText.Trim(), "000000", 0);
                        }
                    }
                    else if (name == "div")
                    {
                        HtmlAgilityPack.HtmlDocument divSectionHtml = new HtmlAgilityPack.HtmlDocument();
                        divSectionHtml.LoadHtml(sectionChilds[j].InnerHtml);

                        HtmlNodeCollection divSection = divSectionHtml.DocumentNode.SelectNodes("/ul/li/a");
                        section = new HtmlElement[divSection.Count + 1];

                        string text = String.Empty;
                        for (int k = 0; k < divSection.Count + 1; k++)
                        {
                            if (k == 0)
                            {
                                text = "Подробнее";
                            }
                            else
                            {
                                text = divSection[k - 1].InnerText;
                            }
                            ((HtmlElement[])section)[k] = new HtmlElement(text, "000000", 0);
                        }
                    }
                    else
                    {
                        section = new HtmlElement(sectionChilds[j].InnerText, "000000", 0);
                    }

                    childs[j] = section;
                }
                vulnSections[i] = childs;
            }
            HtmlNodeCollection imgNodes = htmlDoc.DocumentNode.SelectNodes("//img");
            imgData = new Image[imgNodes.Count];

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
        }

        public override void ReportToWord(string nameWord, bool groupExport = false, Word.Document doc = null, Word.Application app = null, Word.Paragraph paragraph = null)
        {
            if (summary == null || vulnSections == null || tables == null || imgData == null || tablesCaptions == null)
                 MessageBox.Show("Отчет СЗИ \"Сканнер-ВС\" был обработан некорректно, убедитесь в правильности выбора отчета", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                 if (!groupExport)
                 {
                     app = new Word.Application();
                     doc = app.Documents.Add(Environment.CurrentDirectory + "/" + nameWord);
                     paragraph = doc.Paragraphs.Add();
                 }
                 app.Visible = true;
                 int captionIndex = 0;
                 int tableCaprionIndex = 0;
                int vulnSectionIndex = 0; 

                 FillRangeInWord(paragraph.Range, "Приложение Б", "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdColor.wdColorBlack);

                 #region Раздел "Резюме для руководителей"
                 
                 paragraph.Range.InsertParagraphAfter();
                 FillRangeInWord(paragraph.Range, captions[captionIndex++].Text, "Times New Roman", 16, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack);
                 paragraph.Range.InsertParagraphAfter();
                 for (int i = 0; i < summary.Length; i++)
                 {
                     FillRangeInWord(paragraph.Range, summary[i].Text, "Times New Roman", 14, summary[i].Bold, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack);
                     paragraph.Range.InsertParagraphAfter();
                 }
                 paragraph.Range.InsertParagraphAfter();

                 FillRangeInWord(paragraph.Range, tablesCaptions[tableCaprionIndex++].Text, "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorGray85);
                 paragraph.Range.InsertParagraphAfter();

                 Word.Table table = CreateStandartTable(paragraph.Range, tables[0].Length, 3, Word.WdLineStyle.wdLineStyleSingle, Word.WdLineStyle.wdLineStyleSingle, doc);
                //int count = table.Rows[1].Cells.Count;
                //for (int i = 3; i < count + 1; i++)
                //{
                //    table.Rows[1].Cells[2].Merge(table.Rows[1].Cells[3]);
                //}
                //table.Rows[1].Cells[1].Merge(table.Rows[2].Cells[1]);
                for (int i = 0; i < tables[0].Length; i++)
                {
                    for (int j = 0; j < tables[0][i].Length; j++)
                    {
                        FillRangeInWord(table.Cell(i + 1, j + 1).Range, tables[0][i][j].Text, "Times New Roman", 12, tables[0][i][j].Bold, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack, true, tables[0][i][j].wdBackColor);
                    }
                }
                paragraph.Range.InsertParagraphAfter();

                // Картинки тут

                #endregion

                #region Раздел "Границы проекта"
                FillRangeInWord(paragraph.Range, captions[captionIndex++].Text, "Times New Roman", 16, 1, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                paragraph.Range.InsertParagraphAfter();
                table = CreateStandartTable(paragraph.Range, 3, 7, Word.WdLineStyle.wdLineStyleSingle, Word.WdLineStyle.wdLineStyleSingle, doc);
                int count = table.Rows[1].Cells.Count;
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

                for (int i = 0; i < tables[1].Length; i++)
                {
                    for (int j = 0; j < tables[1][i].Length; j++)
                    {
                        if (i > 2)
                            table.Rows.Add();
                        FillRangeInWord(table.Cell(i + 1, j + 1).Range, tables[1][i][j].Text, "Times New Roman", 12, tables[1][i][j].Bold, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack, true, tables[1][i][j].wdBackColor);
                    }
                }
                //FillRangeInWord(table.Cell(1, 1).Range, tables_Nodes[1, 1][0].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                //FillRangeInWord(table.Cell(1, 2).Range, tables_Nodes[1, 1][1].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                //FillRangeInWord(table.Cell(1, 3).Range, tables_Nodes[1, 1][2].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                //FillRangeInWord(table.Cell(1, 4).Range, tables_Nodes[1, 1][3].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                //FillRangeInWord(table.Cell(2, 4).Range, tables_Nodes[1, 1][4].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                //FillRangeInWord(table.Cell(2, 5).Range, tables_Nodes[1, 1][5].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                //FillRangeInWord(table.Cell(2, 6).Range, tables_Nodes[1, 1][6].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                //FillRangeInWord(table.Cell(3, 6).Range, tables_Nodes[1, 1][7].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                //FillRangeInWord(table.Cell(3, 7).Range, tables_Nodes[1, 1][8].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                //for (int i = 0; i < tables_Nodes[1, 2].Count / 5; i++)
                //{
                //    table.Rows.Add();
                //    for (int j = 0; j < count; j++)
                //    {
                //        FillRangeInWord(table.Cell(i + 4, j + 1).Range, tables_Nodes[1, 2][count * i + j].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                //    }
                //}
                paragraph.Range.InsertParagraphAfter();
                for (int i = 0; i < projectBordersNote.Length; i++)
                {
                    FillRangeInWord(paragraph.Range, projectBordersNote[i].Text, "Times New Roman", 12, projectBordersNote[i].Bold, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorBlack);
                    paragraph.Range.InsertParagraphAfter();
                }
                paragraph.Range.InsertParagraphAfter();
                #endregion

                #region Порты и сервисы
                FillRangeInWord(paragraph.Range, captions[captionIndex++].Text, "Times New Roman", 16, 1, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                paragraph.Range.InsertParagraphAfter();
                table = CreateStandartTable(paragraph.Range, tables[2].Length, 5, Word.WdLineStyle.wdLineStyleSingle, Word.WdLineStyle.wdLineStyleSingle, doc);
                for (int i = 0; i < tables[2].Length; i++)
                {
                    for (int j = 0; j < tables[2][i].Length; j++)
                    {
                        FillRangeInWord(table.Cell(i + 1, j + 1).Range, tables[2][i][j].Text, "Times New Roman", 12, tables[2][i][j].Bold, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack, true, tables[2][i][j].wdBackColor);
                    }
                }
                paragraph.Range.InsertParagraphAfter();
                #endregion

                #region Уязвимости

                #region Уязвимость обхода проверки подлинности сеанса в Microsoft Windows SMB / NETBIOS, связанной с пустым значением

                FillRangeInWord(paragraph.Range, captions[vulnSectionIndex++].Text, "Times New Roman", 16, 1, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                paragraph.Range.InsertParagraphAfter();
                FillRangeInWord(paragraph.Range, captions[vulnSectionIndex++].Text, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                paragraph.Range.InsertParagraphAfter();
                FillRangeInWord(paragraph.Range, captions[vulnSectionIndex++].Text, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                paragraph.Range.InsertParagraphAfter();
                FillRangeInWord(paragraph.Range, captions[vulnSectionIndex++].Text, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                paragraph.Range.InsertParagraphAfter();
                FillRangeInWord(paragraph.Range, captions[vulnSectionIndex++].Text, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                paragraph.Range.InsertParagraphAfter();
                count = vulnSectionIndex;
                //for (int i = 0; i < vulnSections[0][)
                #endregion

                #endregion

                //#region CVSS 2.0
                //FillRangeInWord(paragraph.Range, tables_Nodes[2, 0][0].InnerText, "Times New Roman", 14, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdColor.wdColorGray85);
                //paragraph.Range.InsertParagraphAfter();
                //table = CreateStandartTable(paragraph.Range, 6, 4, Word.WdLineStyle.wdLineStyleSingle, Word.WdLineStyle.wdLineStyleSingle, doc);
                //count = table.Rows[1].Cells.Count;
                //for (int i = 0; i < tables_Nodes[2, 1].Count / 4; i++)
                //{
                //    for (int j = 0; j < count; j++)
                //    {
                //        FillRangeInWord(table.Cell(i + 1, j + 1).Range, tables_Nodes[2, 1][count * i + j].InnerText, "Times New Roman", 12, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdColor.wdColorBlack);
                //    }
                //}
                //#endregion
            }

            paragraph.Range.InsertBreak(Word.WdBreakType.wdPageBreak);
        }
    }
}
