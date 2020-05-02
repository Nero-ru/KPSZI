using KPSZI.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Threading;

namespace KPSZI
{
    class StageMeasures:Stage
    {
        protected override ImageList imageListForTabPage { get; set; }
        Button btnExportConsole = new Button();
        Button btnExportOpenXML = new Button();
        Button btnExportInterop = new Button();
        Button btnExprotSFHInterop = new Button();
        Button btnExportSFHConsole = new Button();
        Button btnGetBasicMeasuresList;
        Button btnGetAdaptiveMeasuresList;
        Button btnGetConreteMeasuresList;
        Button btnConsoleClear = new Button();

        public static List<GISMeasure> ListOfBasicMeasures = new List<GISMeasure>();
        public static List<GISMeasure> ListOfAdaptiveMeasures = new List<GISMeasure>();
        public static List<GISMeasure> ListOfExcludedMeasures = new List<GISMeasure>();
        public static List<GISMeasure> ListOfConcreteMeasures = new List<GISMeasure>();
        public static List<GISMeasure> ListOfAddedMeasures = new List<GISMeasure>();

        public StageMeasures(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS)
            : base(stageTab, stageNode, mainForm, IS)
        {
            ////Дебаг вариант с статическими угрозами
            //using (KPSZIContext db = new KPSZIContext())
            //{
            //    var Threats = db.Threats.Include("GISMeasures").ToList();
            //    IS.listOfActualNSDThreats = Threats.Where(t =>
            //        t.ThreatNumber == 4 ||
            //        t.ThreatNumber == 7 ||
            //        t.ThreatNumber == 16 ||
            //        t.ThreatNumber == 18 ||
            //        t.ThreatNumber == 23 ||
            //        t.ThreatNumber == 30 ||
            //        t.ThreatNumber == 31 ||
            //        t.ThreatNumber == 32 ||
            //        t.ThreatNumber == 39 ||
            //        t.ThreatNumber == 42 ||
            //        t.ThreatNumber == 45 ||
            //        t.ThreatNumber == 18 ||
            //        t.ThreatNumber == 53 ||
            //        t.ThreatNumber == 67 ||
            //        t.ThreatNumber == 72 ||
            //        t.ThreatNumber == 88 ||
            //        t.ThreatNumber == 91 ||
            //        t.ThreatNumber == 94 ||
            //        t.ThreatNumber == 95 ||
            //        t.ThreatNumber == 111 ||
            //        t.ThreatNumber == 113 ||
            //        t.ThreatNumber == 117 ||
            //        t.ThreatNumber == 122 ||
            //        t.ThreatNumber == 127 ||
            //        t.ThreatNumber == 131 ||
            //        t.ThreatNumber == 132 ||
            //        t.ThreatNumber == 139 ||
            //        t.ThreatNumber == 156 ||
            //        t.ThreatNumber == 157 ||
            //        t.ThreatNumber == 160 ||
            //        t.ThreatNumber == 179 ||
            //        t.ThreatNumber == 180 ||
            //        t.ThreatNumber == 182 ||
            //        t.ThreatNumber == 185 ||
            //        t.ThreatNumber == 186 ||
            //        t.ThreatNumber == 190 ||
            //        t.ThreatNumber == 191 ||
            //        t.ThreatNumber == 201
            //    ).ToList();
            //}
        }

        protected override void initTabPage()
        {

            mf.tabControlMeasures.TabPages.Clear();
            btnGetBasicMeasuresList = mf.btnGetBasicMeasuresList;
            btnGetAdaptiveMeasuresList = mf.btnGetAdaptiveMeasuresList;
            btnGetConreteMeasuresList = mf.btnGetConreteMeasuresList;

            mf.dgvBasicMeas.Columns.Add("Count", "№");
            mf.dgvBasicMeas.Columns.Add("Name", "Наименование меры");
            mf.dgvBasicMeas.Columns[0].Width = 30;//.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            mf.dgvBasicMeas.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            mf.dgvAdaptiveMeas.Columns.Add("Count", "№");
            mf.dgvAdaptiveMeas.Columns.Add("Name", "Наименование меры");
            mf.dgvAdaptiveMeas.Columns[0].Width = 30;//.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            mf.dgvAdaptiveMeas.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            mf.dgvThrMeas.Columns.Add("Count", "№");
            mf.dgvThrMeas.Columns.Add("ThreatName", "Угроза");
            mf.dgvThrMeas.Columns.Add("Measures", "Меры по нейтрализации УБИ");
            mf.dgvThrMeas.Columns[0].Width = 30;//.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            mf.dgvThrMeas.Columns[1].Width = 150;//.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            mf.dgvThrMeas.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            mf.dgvConcreteMeas.Columns.Add("Count", "№");
            mf.dgvConcreteMeas.Columns.Add("Name", "Наименование меры");
            mf.dgvConcreteMeas.Columns[0].Width = 30;//.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            mf.dgvConcreteMeas.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            btnExportConsole.Location = new System.Drawing.Point(100, 800);
            btnExportConsole.Text = "В консоль";
            btnExportConsole.Click += btnExportConsole_Click;
            //mf.tabControl.TabPages[stageTab.Name].Controls.Add(btnExportConsole);

            btnExportOpenXML.Location = new System.Drawing.Point(300, 800);
            btnExportOpenXML.Text = "OpenXML";
            btnExportOpenXML.Click += BtnExport_Click;
            //mf.tabControl.TabPages[stageTab.Name].Controls.Add(btnExportOpenXML);

            btnExportInterop.Location = new System.Drawing.Point(500, 800);
            btnExportInterop.Text = "Interop";
            btnExportInterop.Click += BtnExportInterop_Click;
            //mf.tabControl.TabPages[stageTab.Name].Controls.Add(btnExportInterop);

            btnExportSFHConsole.Location = new System.Drawing.Point(100, 900);
            btnExportSFHConsole.Text = "В консоль";
            btnExportSFHConsole.Click += BtnExportSFHConsole_Click;
            //mf.tabControl.TabPages[stageTab.Name].Controls.Add(btnExportSFHConsole);

            btnExprotSFHInterop.Location = new System.Drawing.Point(500, 900);
            btnExprotSFHInterop.Text = "Interop";
            btnExprotSFHInterop.Click += BtnExprotSFHInterop_Click;
           // mf.tabControl.TabPages[stageTab.Name].Controls.Add(btnExprotSFHInterop);

            btnGetBasicMeasuresList.Text = "Базовый";
            btnGetBasicMeasuresList.Click += BtnGetBasicMeasuresList_Click;
            mf.tabControl.TabPages[stageTab.Name].Controls.Add(btnGetBasicMeasuresList);

            btnGetAdaptiveMeasuresList.Text = "Адаптированный";
            btnGetAdaptiveMeasuresList.Click += BtnGetAdaptiveMeasuresList_Click;
            mf.tabControl.TabPages[stageTab.Name].Controls.Add(btnGetAdaptiveMeasuresList);

            btnGetConreteMeasuresList.Text = "Уточненный";
            btnGetConreteMeasuresList.Click += BtnGetConreteMeasuresList_Click;
            mf.tabControl.TabPages[stageTab.Name].Controls.Add(btnGetConreteMeasuresList);

            btnConsoleClear.Location = new System.Drawing.Point(300, 900);
            btnConsoleClear.Text = "Clear";
            btnConsoleClear.Click += BtnConsoleClear_Click;
            //mf.tabControl.TabPages[stageTab.Name].Controls.Add(btnConsoleClear);
        }

        private void BtnConsoleClear_Click(object sender, EventArgs e)
        {
            Console.Clear();
        }

        private void BtnGetBasicMeasuresList_Click(object sender, EventArgs e)
        {
            if (IS.GISClass == 0)
            {
                if (MessageBox.Show("Не определен класс защищенности для дальнейшней работы. \nПерейти во вкладку \"Классификация\"?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    mf.treeView.SelectedNode = mf.returnTreeNode("tnClassification");
                    return;
                }
                else
                    return;
            }

            mf.wsm.Visible = true;
            mf.wsm.Update();

            mf.tabControlMeasures.TabPages.Clear();

            mf.tabControlMeasures.TabPages.Add(mf.tpBasicMeas);
            mf.btnGetAdaptiveMeasuresList.Enabled = true;
            mf.btnGetConreteMeasuresList.Enabled = false;
            
            mf.tbtpMeasDescription.Text = "Базовый набор мер, сформированный в соответствии с установленным классом защищенности информационной системы.";


            //Очищаем DGV от старых мер
            mf.dgvBasicMeas.Rows.Clear();

            using (KPSZIContext db = new KPSZIContext())
            {
                int i = 0;
                ListOfBasicMeasures = db.GisMeasures.Where(gm => gm.MinimalRequirementDefenceClass >= IS.GISClass).OrderBy(gm => gm.GISMeasureId).ToList();

                foreach(GISMeasure gm in ListOfBasicMeasures)
                {
                    Console.WriteLine(gm.ToString());
                    mf.dgvBasicMeas.Rows.Add(++i, gm.ToString());
                }
            }


            mf.wsm.Visible = false;
        }

        private void BtnGetAdaptiveMeasuresList_Click(object sender, EventArgs e)
        {
            mf.wsm.Visible = true;
            mf.wsm.Update();


            mf.tabControlMeasures.TabPages.Clear();

            mf.tabControlMeasures.TabPages.Add(mf.tpAdaptiveMeas);
            mf.btnGetConreteMeasuresList.Enabled = true;

            mf.tbtpMeasDescription.Text = "Адаптированный набор мер, сформированный в соответствии с особенностями функционирования ИС, структурно-функциональным характеристикам и информационным технологиями, характерными для данной ИС.";

            List<GISMeasure> listOfGMForSFHs = new List<GISMeasure>();
            using (KPSZIContext db = new KPSZIContext())
            {
                foreach (SFH sfh in IS.listOfSFHs)
                {
                    var aaa = db.SFHs.Where(s => s.SFHNumber == sfh.SFHNumber).FirstOrDefault();

                    listOfGMForSFHs.AddRange(aaa.GISMeasures);
                }

                listOfGMForSFHs = listOfGMForSFHs.Distinct().ToList();
                listOfGMForSFHs = listOfGMForSFHs.OrderBy(gm => gm.GISMeasureId).ToList();

                int i = 0;
                //очистить DGV от старых мер
                mf.dgvAdaptiveMeas.Rows.Clear();
                ListOfAdaptiveMeasures = ListOfBasicMeasures.Intersect(listOfGMForSFHs).ToList();
                foreach (GISMeasure gm in ListOfAdaptiveMeasures)
                {
                    Console.WriteLine(gm.ToString());
                    mf.dgvAdaptiveMeas.Rows.Add(++i, gm.ToString());
                }
                ListOfExcludedMeasures = ListOfBasicMeasures.Except(ListOfAdaptiveMeasures).ToList();

                Console.WriteLine("Базовый: {0}\nАдаптированный: {1}\nПересечение: {2}", ListOfBasicMeasures.Count, listOfGMForSFHs.Count, ListOfAdaptiveMeasures.Count);
            }

            mf.wsm.Visible = false;
        }

        private void BtnGetConreteMeasuresList_Click(object sender, EventArgs e)
        {
            mf.wsm.Visible = true;
            mf.wsm.Update();

            mf.tabControlMeasures.TabPages.Clear();

            mf.tabControlMeasures.TabPages.Add(mf.tpConcreteMeas);

            mf.tbtpMeasDescription.Text = "Уточненный набор мер, сформированный с учетом не выбранных ранее мер защиты информации, обеспечивает нейтрализацию актуальных угроз безопасности информации, включенных в Модель угроз безопасности информации. \nВ первой таблице представлены меры защиты информации, нейтрализующие все угрозы из Модели.\n Во второй таблице - итоговый перечень мер";

            using (KPSZIContext db = new KPSZIContext())
            {
                var listOfThreatsDB = db.Threats.ToList();

                var LisfOfThreats = listOfThreatsDB.Intersect(IS.listOfActualNSDThreats).ToList();

                List<GISMeasure> listOfMeasuresTemporary = new List<GISMeasure>();

                int i = 0; 
                mf.dgvThrMeas.Rows.Clear();
                mf.dgvConcreteMeas.Rows.Clear();

                foreach (Threat thr in LisfOfThreats)
                {
                    listOfMeasuresTemporary.AddRange(thr.GISMeasures);

                    string _setOfMeasures = "";
                    foreach (GISMeasure gm in thr.GISMeasures)
                    {
                        _setOfMeasures += gm.ToString()+";\n";
                    }
                    mf.dgvThrMeas.Rows.Add(++i, thr.ToString(), _setOfMeasures);
                }

                listOfMeasuresTemporary = listOfMeasuresTemporary.Distinct().ToList();
                listOfMeasuresTemporary = listOfMeasuresTemporary.OrderBy(m => m.GISMeasureId).ToList();

                ListOfConcreteMeasures = listOfMeasuresTemporary;
                ListOfConcreteMeasures.AddRange(ListOfAdaptiveMeasures);
                ListOfConcreteMeasures = ListOfConcreteMeasures.Distinct().OrderBy(m=>m.GISMeasureId).ToList();

                i = 0;
                foreach (GISMeasure gm in ListOfConcreteMeasures)
                {
                    Console.WriteLine("{0}.{1}. {2}", gm.MeasureGroup.ShortName, gm.Number, gm.Description);
                    mf.dgvConcreteMeas.Rows.Add(++i, gm.ToString());
                }
                Console.WriteLine("Базовый: {0}\nАдаптированный: {1}\nУточненный: {2}", ListOfBasicMeasures.Count, ListOfAdaptiveMeasures.Count, ListOfConcreteMeasures.Count);
                IS.listOfAllNSDMeasures = ListOfConcreteMeasures;
                ListOfAddedMeasures = ListOfConcreteMeasures.Except(ListOfAdaptiveMeasures).ToList();
            }

            mf.wsm.Visible = false;
        }

        private void BtnExportSFHConsole_Click(object sender, EventArgs e)
        {
            int i = 0;
            using (KPSZIContext db = new KPSZIContext())
            {
                var sfhs = db.SFHs.OrderBy(s => s.SFHNumber).ToList();

                foreach (SFH sfh in sfhs)
                {
                    Console.WriteLine("===={0}====", sfh.Name);
                    foreach (GISMeasure gm in sfh.GISMeasures.OrderBy(m => m.GISMeasureId))
                    {
                        Console.WriteLine("\t {0}.{1} {2}", gm.MeasureGroup.ShortName, gm.Number, gm.Description);
                    }
                }
            }
        }

        private void BtnExprotSFHInterop_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            Microsoft.Office.Interop.Word.Paragraph wordParag;
            Microsoft.Office.Interop.Word.Table wordTable;

            //создаём новый документ Word и задаём параметры листа
            wordDoc = wordApp.Documents.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing); //создаём документ Word

            // первый параграф
            wordParag = wordDoc.Paragraphs.Add(Type.Missing);
            //wordParag.Range.Font.Name = "Times New Roman";
            //wordParag.Range.Font.Size = 16;
            //wordParag.Range.Font.Bold = 1;
            //wordParag.Range.Text = "Заголовок";
            //wordParag.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;


            // второй параграф, таблица из n строк и 2 колонок
            //wordDoc.Paragraphs.Add(Type.Missing);
            wordParag.Range.Tables.Add(wordParag.Range, 45, 2, Type.Missing, Type.Missing);
            wordTable = wordDoc.Tables[1];
            wordTable.Range.Font.Bold = 0;
            wordParag.Range.Font.Name = "Times New Roman";
            wordTable.Range.Font.Size = 14;


            //задаём ширину колонок и высоту строк
            wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints;
            wordTable.Columns[1].SetWidth(200f, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
            wordTable.Rows.SetHeight(20f, Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAuto);
            wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
            wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable.Range.Select();

            //заполнение шапки таблицы
            wordTable.Cell(1, 1).Range.Text = "Структурно–функциональная характеристика информационной системы";
            wordTable.Cell(1, 2).Range.Text = "Мера защиты информации, которая может применяться в информационной системе с данной характеристикой";

            //заполняем таблицу 

            int row = 2;
            using (KPSZIContext db = new KPSZIContext())
            {
                var sfhs = db.SFHs.OrderBy(t => t.SFHNumber).ToList();

                foreach (SFH sfh in sfhs)
                {
                    string measuresSFH = "";
                    foreach (GISMeasure gm in sfh.GISMeasures.OrderBy(m => m.GISMeasureId))
                    {
                        measuresSFH += gm.MeasureGroup.ShortName + "." + gm.Number + ". " + gm.Description + ";\n";
                    }
                    measuresSFH.TrimEnd('\n');
                    measuresSFH.TrimEnd(';');
                    wordTable.Cell(row, 1).Range.Text = sfh.Name;
                    wordTable.Cell(row++, 2).Range.Text = measuresSFH;
                }
            }

            //сохраняем документ, закрываем документ, выходим из Word
            wordDoc.SaveAs(Directory.GetCurrentDirectory() + @"\СФХ-меры.docx");
            wordApp.ActiveDocument.Close();
            wordApp.Quit();
        }

        private void BtnExportInterop_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            Microsoft.Office.Interop.Word.Paragraph wordParag;
            Microsoft.Office.Interop.Word.Table wordTable;

            //создаём новый документ Word и задаём параметры листа
            wordDoc = wordApp.Documents.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing); //создаём документ Word

            // первый параграф
            wordParag = wordDoc.Paragraphs.Add(Type.Missing);
            //wordParag.Range.Font.Name = "Times New Roman";
            //wordParag.Range.Font.Size = 16;
            //wordParag.Range.Font.Bold = 1;
            //wordParag.Range.Text = "Заголовок";
            //wordParag.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;


            // второй параграф, таблица из n строк и 2 колонок
            //wordDoc.Paragraphs.Add(Type.Missing);
            wordParag.Range.Tables.Add(wordParag.Range, 209, 3, Type.Missing, Type.Missing);
            wordTable = wordDoc.Tables[1];
            wordTable.Range.Font.Bold = 0;
            wordParag.Range.Font.Name = "Times New Roman";
            wordTable.Range.Font.Size = 14;

            //задаём ширину колонок и высоту строк
            wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints;
            wordTable.Columns[1].SetWidth(200f, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
            wordTable.Rows.SetHeight(20f, Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAuto);
            wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
            wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable.Range.Select();

            //заполнение шапки таблицы
            wordTable.Cell(1, 1).Range.Text = "Номер УБИ";
            wordTable.Cell(1, 2).Range.Text = "Наименование УБИ";
            wordTable.Cell(1, 3).Range.Text = "Мера";

            //заполняем таблицу 

            int row = 2;
            using (KPSZIContext db = new KPSZIContext())
            {
                var threats = db.Threats.OrderBy(t => t.ThreatNumber).ToList();

                foreach (Threat thr in threats)
                {
                    string threatName = thr.ThreatNumber + "." + thr.Name;
                    string measuresThreat = "";
                    foreach (GISMeasure gm in thr.GISMeasures.OrderBy(m => m.GISMeasureId))
                    {
                        measuresThreat += gm.MeasureGroup.ShortName + "." + gm.Number + ". " + gm.Description+";\n";
                    }
                    measuresThreat.TrimEnd('\n');
                    measuresThreat.TrimEnd(';');
                    wordTable.Cell(row, 1).Range.Text = "УБИ." + thr.ThreatNumber.ToString("000");
                    wordTable.Cell(row, 2).Range.Text = thr.Name;
                    wordTable.Cell(row++, 3).Range.Text = measuresThreat;
                }
            }

            //сохраняем документ, закрываем документ, выходим из Word
            wordDoc.SaveAs(Directory.GetCurrentDirectory() + @"\меры.docx");
            wordApp.ActiveDocument.Close();
            wordApp.Quit();
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            //new GeneratedClass().CreatePackage("Output.docx");
        }

        private void btnExportConsole_Click(object sender, EventArgs e)
        {
            mf.wsm.Visible = true;
            mf.wsm.Update();

            using (KPSZIContext db = new KPSZIContext())
            {
                var threats = db.Threats.OrderBy(t => t.ThreatNumber).ToList();

                foreach (Threat thr in threats)
                {
                    Console.WriteLine("{0}. {1}", thr.ThreatNumber, thr.Name);
                    foreach (GISMeasure gm in thr.GISMeasures.OrderBy(m => m.GISMeasureId))
                    {
                        Console.WriteLine("\t {0}.{1} {2}", gm.MeasureGroup.ShortName, gm.Number, gm.Description);
                    }
                }
            }
            mf.wsm.Visible = false;
        }

        public override void saveChanges()
        {

        }

        public override void enterTabPage()
        {

        }
    }
}
