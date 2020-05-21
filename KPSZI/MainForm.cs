using KPSZI.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace KPSZI
{
    public partial class MainForm : Form
    {
        // Dictionary - аналог ассоциативных массивов в C#
        // Обращаемся к элементу массива (коллекции) по названию (string)
        // 
        internal Dictionary<string, Stage> stages = new Dictionary<string, Stage>();
        TreeNode previousSelectedNode;
        InformationSystem IS = new InformationSystem();
        internal WaitingSplashMessage wsm;
        Thread t;

        public void startSplash()
        {
            System.Windows.Forms.Application.Run(new splashForm());
        }

        public MainForm()
        {
            t = new Thread(startSplash);
            t.Start();
            initForm();
            // check database connectio before starting application
            //using (KPSZIContext db = new KPSZIContext())
            //{
            //    if (!db.Database.Exists())
            //    {
            //        t.Abort();
            //        this.Close();
            //        MessageBox.Show("Ошибка подключения к базе данных КПСЗИ");                        
            //    }
            //    else
            //        initForm();                    
            //}

            KeyPreview = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);

            treeView.SelectedNode = returnTreeNode("tnOptions");
        }

        private void initForm()
        {
            InitializeComponent();

            Icon = KPSZI.Properties.Resources.mf;
            StageReportRevisor2XP reportRevisor2XP = new StageReportRevisor2XP(this, @"reports/Revisor2XP/template2XP.docx");
            StageReportFIKS reportFIKS = new StageReportFIKS(this, @"reports/FIKS/templateFIKS.docx");
            StageReportScannerVS reportScannerVS = new StageReportScannerVS(this, @"reports/ScannerVS/tamplateScannerVS.docx");
            StageReportScanoval reportScanoval = new StageReportScanoval(this, @"reports/Scanoval/tamplateScanoval.docx");
            // Заполняем коллекцию этапами (название, ссылка на вкладку, ссылка на пункт в дереве) 
            stages.Add("tnOptions", new StageOptions(returnTabPage("tpOptions"), returnTreeNode("tnOptions"), this, IS));
            stages.Add("tnClassification", new StageClassification(returnTabPage("tpClassification"), returnTreeNode("tnClassification"), this, IS));
            //stages.Add("tnAccessMatrix", new StageAccessMatrix(returnTabPage("tpAccessMatrix"), returnTreeNode("tnAccessMatrix"), this, IS));
            stages.Add("tnTopology", new StageTopology(returnTabPage("tpTopology"), returnTreeNode("tnTopology"), this, IS));
            stages.Add("tnIntruder", new StageIntruder(returnTabPage("tpIntruder"), returnTreeNode("tnIntruder"), this, IS));
            stages.Add("tnActualThreats", new StageActualThreats(returnTabPage("tpActualThreats"), returnTreeNode("tnActualThreats"), this, IS));
            stages.Add("tnHardware", new StageHardware(returnTabPage("tpHardware"), returnTreeNode("tnHardware"), this, IS));
            stages.Add("tnVulnerabilities", new StageVulnerabilities(returnTabPage("tpVulnerabilities"), returnTreeNode("tnVulnerabilities"), this, IS));
            stages.Add("tnTCUI", new StageTCUI(returnTabPage("tpTCUI"), returnTreeNode("tnTCUI"), this, IS));
            stages.Add("tnTechnoGenThreats", new StageTechnoGenThreats(returnTabPage("tpTechnoGenThreats"), returnTreeNode("tnTechnoGenThreats"), this, IS));
            stages.Add("tnSKZI", new StageSKZI(returnTabPage("tpSKZI"), returnTreeNode("tnSKZI"), this, IS));
            stages.Add("tnMeasures", new StageMeasures(returnTabPage("tpMeasures"), returnTreeNode("tnMeasures"), this, IS));
            //stages.Add("tnMeasuresTCUI", new StageMeasuresTCUI(returnTabPage("tpMeasuresTCUI"), returnTreeNode("tnMeasuresTCUI"), this, IS));
            //stages.Add("tnMeasuresTechno", new StageMeasuresTechno(returnTabPage("tpMeasuresTechno"), returnTreeNode("tnMeasuresTechno"), this, IS));
            stages.Add("tnSZI", new StageSZI(returnTabPage("tpSZI"), returnTreeNode("tnSZI"), this, IS));
            //stages.Add("tnTPExport", new StageTPExport(returnTabPage("tpTPExport"), returnTreeNode("tnTPExport"), this, IS));
            stages.Add("tnConfigOptions", new StageConfigOptions(returnTabPage("tpConfigOptions"), returnTreeNode("tnconfigOptions"), this, IS));
            //stages.Add("tnRevisor2XP", reportRevisor2XP);
            //stages.Add("tnFIKS", reportFIKS);
            //stages.Add("tnScannerVS", reportScannerVS);
            //stages.Add("tnScanoval", reportScanoval);
            stages.Add("tnAllReport", new StageAllReport(returnTabPage("tpAllReport"), returnTreeNode("tnAllReport"), this, IS, @"reports/All/Report_v2.docx", reportRevisor2XP, reportFIKS, reportScannerVS, reportScanoval));
            stages.Add("tnSZIConfig", new StageSZIConfig(returnTabPage("tpSZIConfig"), returnTreeNode("tnSZIConfig"), this, IS));
            stages.Add("tnCertificateSZI", new StageCertificateSZI(returnTabPage("tpCertificateSZI"), returnTreeNode("tnCertificateSZI"), this, IS));
            //returnTreeNode("tnActualThreats").ForeColor = Color.Gray;
            //returnTreeNode("tnActualThreats").BackColor = Color.White;

            DateOfReportTextBox.Text = Convert.ToString(DateTime.Now.ToString("D"));

            // закрываем все вкладки в TabControl
            tabControl.TabPages.Clear();

            // связываем дерево с набором иконок
            iconList.Images.Add(KPSZI.Properties.Resources.folder_icon);
            iconList.Images.Add(KPSZI.Properties.Resources.document_settings_icon);
            iconList.Images.Add(KPSZI.Properties.Resources.left_arrow_icon);
            iconList.Images.Add(KPSZI.Properties.Resources.right_arrow_icon);

            treeView.ImageList = iconList;

            // развернуть дерево
            treeView.ExpandAll();

            foreach (TabPage tab in tabControl.TabPages)
                tab.AutoScroll = true;

            menuStrip.BackColor = Color.FromArgb(234, 240, 255);
            //this.BackColor = Color.FromArgb(234, 240, 255);


            tabControlInfoTypes.TabPages.AddRange(((StageClassification)stages["tnClassification"]).tabPagesInfoTypes.ToArray());
            t.Abort();

            //Создание окна "подождика пока я работаю.."
            wsm = new WaitingSplashMessage();
            this.Controls.Add(wsm);
            wsm.Location = new System.Drawing.Point(this.Width / 2 - wsm.Width/2, this.Height / 2 - wsm.Height/2);
            wsm.BringToFront();
            wsm.Visible = false;
        }

        // возвращает ссылку на TabPage по имени вкладки
        public TabPage returnTabPage(string tpName)
        {
            return tabControl.TabPages[tabControl.TabPages.IndexOfKey(tpName)];
        }
        // возвращает ссылку на TreeNode по имени пункта дерева
        public TreeNode returnTreeNode(string tnName)
        {
            return treeView.Nodes.Find(tnName, true)[0];
        }

        // Событие: После переключения вкладки
        private void treeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (treeView.SelectedNode.ForeColor == Color.Gray)
            {
                treeView.SelectedNode = previousSelectedNode;
                return;
            }

            tabControl.TabPages.Clear();
            
            // Получаем имя Node в дереве
            string nodeName = treeView.SelectedNode.Name;
            // Если дочерних Нодов нет ...
            if (treeView.SelectedNode.Nodes.Count == 0)
            {
                // ... Открываем вкладку этапа и выполняем enterTabPage
                tabControl.TabPages.Add(stages[nodeName].stageTab);
                tabControl.SelectedTab.Text = treeView.SelectedNode.Text;
                stages[nodeName].enterTabPage();
            }

            treeView.Focus();
        }

        // Событие: До переключения вкладки
        private void treeView_BeforeSelect(object sender, TreeViewCancelEventArgs e)
        {
            previousSelectedNode = treeView.SelectedNode;
            if (treeView.SelectedNode != null && treeView.SelectedNode.Nodes.Count == 0)
                stages[treeView.SelectedNode.Name].saveChanges();
        }

        private void rewriteThreatsDBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Перезаписать базу угроз
            if (MessageBox.Show("Будут удалены все угрозы из таблицы в БД и записаны заново.\nПродолжить?", "Ахтунг!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                using (KPSZIContext db = new KPSZIContext())
                {
                    FileInfo fi = new FileInfo("thrlist.xlsx");

                    //try
                    //{
                        //// Каскадное удаление данных вместе с внешними ключами
                        //db.Database.ExecuteSqlCommand("SET SCHEMA '" + KPSZIContext.schema_name + "'; TRUNCATE \"Threats\" CASCADE;");

                        db.Threats.AddRange(Threat.GetThreatsFromXlsx(fi, db));
                        db.SaveChanges();
                        db.SeedForThreat();
                        db.SaveChanges();
                    //}
                    //catch (Exception ex)
                    //{
                    //    MessageBox.Show("В rewriteThreatsDBToolStripMenuItem_Click Exception!\n" + ex.Message, "Ахтунг!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    return;
                    //}

                }
                MessageBox.Show("Таблица угроз успешно перезаписана!", "Это успех, парень!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void clearDBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Снести всю базу
            if (MessageBox.Show("Будут очищены все таблицы (не удалены) кроме данных о миграциях", "Ахтунг!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
            {
                using (KPSZIContext db = new KPSZIContext())
                {
                    try
                    {
                        db.Database.ExecuteSqlCommand("SET SCHEMA '" + KPSZIContext.schema_name + "'; TRUNCATE \"ConfigOptions\",\"GISMeasures\", \"ISPDNMeasures\", \"ImplementWayThreats\", \"ImplementWays\", \"InfoTypes\", \"IntruderTypes\", \"MeasureGroups\", \"SFHGISMeasures\", \"SFHTypes\", \"SFHs\", \"SZISortGISMeasures\", \"SZISortSZIs\", \"SZISorts\",  \"SZIs\", \"TCUIThreats\", \"TCUITypes\", \"TCUIs\", \"TechnogenicMeasures\", \"TechnogenicThreats\", \"ThreatGISMeasures\",  \"ThreatSFHs\", \"ThreatSourceThreats\", \"ThreatSources\", \"Threats\", \"Vulnerabilities\", \"VulnerabilityThreats\" CASCADE");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("В clearDBToolStripMenuItem_Click Exception!\n" + ex.Message, "Ахтунг!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                MessageBox.Show("Таблицы базы данных успешно очищены", "Это успех, парень!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void refreshThreatDBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Обновить список угроз
            if (MessageBox.Show("Будет проведено сравнение локальной базы угроз с данными из файла \"thrlist.xlsx\" и автоматическое обновление данных.\nПродолжить?", "Ахтунг!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                using (KPSZIContext db = new KPSZIContext())
                {
                    try
                    {
                        FileInfo fi = new FileInfo("thrlist.xlsx");
                        List<Threat> listThreatsFromFile = Threat.GetThreatsFromXlsx(fi, db);
                        List<Threat> listThreatsFromDB = db.Threats.OrderBy(t => t.ThreatNumber).ToList();

                        DateTime lastUpdateOfLocalDB;
                        if (listThreatsFromDB != null)
                        {
                            // Получение даты последнего обновления угроз из БД

                            lastUpdateOfLocalDB = listThreatsFromDB.Select(t => t.DateOfChange).Max();
                        }
                        else
                        {
                            lastUpdateOfLocalDB = DateTime.MinValue;
                        }

                        // Получение даты последнего обновления угроз из актуального файла с угрозами
                        DateTime lastUpdateOfFile = listThreatsFromFile.Select(t => t.DateOfChange).Max();

                        // Если нет изменений, то прекратить обновление
                        if (lastUpdateOfLocalDB == lastUpdateOfFile && listThreatsFromDB.Count == listThreatsFromFile.Count)
                        {
                            MessageBox.Show("База угроз не требует обновления!", "КПСЗИ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        // Отбор угроз, претерпевших изменения
                        List<Threat> listChangedOrAddedThreats = listThreatsFromFile.Except(listThreatsFromDB).ToList();

                        // Внесение изменений
                        foreach(Threat thr in listChangedOrAddedThreats)
                        {
                            Threat ThrFromDB = listThreatsFromDB.Where(t => t.ThreatNumber == thr.ThreatNumber).FirstOrDefault();

                            if (ThrFromDB == null)
                            {
                                //listThreatsFromDB.Add(thr);
                                db.Threats.Add(thr);

                                continue;
                            }

                            ThrFromDB.AvailabilityViolation = thr.AvailabilityViolation;
                            ThrFromDB.ConfidenceViolation = thr.ConfidenceViolation;
                            ThrFromDB.DateOfAdd = thr.DateOfAdd;
                            ThrFromDB.DateOfChange = thr.DateOfChange;
                            ThrFromDB.Description = thr.Description;
                            ThrFromDB.IntegrityViolation = thr.IntegrityViolation;
                            ThrFromDB.Name = thr.Name;
                            ThrFromDB.ObjectOfInfluence = thr.Name;
                            ThrFromDB.ThreatNumber = thr.ThreatNumber;
                            ThrFromDB.ThreatSources = thr.ThreatSources;
                        }

                        db.SaveChanges();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("В refreshThreatDBToolStripMenuItem_Click Exception!\n" + ex.Message, "Ахтунг!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                MessageBox.Show("Список угроз обновлен", "Это успех, парень!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void downloadThreatListthrlistxlsxToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Загрузить список угроз
            try
            {
                using (var client = new WebClient())
                {
                    client.DownloadFile("https://bdu.fstec.ru/files/documents/thrlist.xlsx", "thrlist.xlsx");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Проблемы при загрузке файла thrlist.xlsx.\n"+ ex.Message, "Ошибка загрузки файла", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            MessageBox.Show("Файл успешно загружен", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void initDBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // инициализация БД начальным значениями из метода Model.KPSZIContext.Seed()

            if (MessageBox.Show("БД будет проинициализирована начальными значениями. Перед выполнением процедуры необходимо очистить (не удалить!) все таблицы.\nПродолжить?", "Ахтунг!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
            {
                using (KPSZIContext db = new KPSZIContext())
                {
                    //try
                    //{
                    //    db.Database.ExecuteSqlCommand("SET SCHEMA '" + KPSZIContext.schema_name + "'; TRUNCATE \"ConfigOptions\",\"GISMeasures\", \"ISPDNMeasures\", \"ImplementWayThreats\", \"ImplementWays\", \"InfoTypes\", \"IntruderTypes\", \"MeasureGroups\", \"SFHGISMeasures\", \"SFHTypes\", \"SFHs\", \"SZISortGISMeasures\", \"SZISortSZIs\", \"SZISorts\",  \"SZIs\", \"TCUIThreats\", \"TCUITypes\", \"TCUIs\", \"TechnogenicMeasures\", \"TechnogenicThreats\", \"ThreatGISMeasures\",  \"ThreatSFHs\", \"ThreatSourceThreats\", \"ThreatSources\", \"Threats\", \"Vulnerabilities\", \"VulnerabilityThreats\" CASCADE");
                    //}
                    //catch (Exception ex)
                    //{
                    //    MessageBox.Show("Ошибка удаления таблиц в БД\n" + ex.Message, "Ахтунг!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    return;
                    //}
                    db.Seed();

                    FileInfo fi = null;
                    try
                    {
                        fi = new FileInfo("_reestr_sszi.xlsx");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Сначала надо загрузить файл (База данных -> Реестр ФСТЭК СЗИ -> Скачать файл \"_reestr_sszi\"\n" + ex.Message, "Ахтунг!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    db.CertificatesSZI.AddRange(CertificateSZI.GetThreatsFromXlsx(fi, db));
                    db.SaveChanges();

                    try
                    {
                        fi = new FileInfo("thrlist.xlsx");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Сначала надо загрузить файл (База данных -> Реестр ФСТЭК СЗИ -> Скачать файл \"thrlist.xlsx\"\n" + ex.Message, "Ахтунг!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    db.Threats.AddRange(Threat.GetThreatsFromXlsx(fi, db));
                    db.SaveChanges();
                    db.SeedForThreat();
                    db.SaveChanges();

                    try
                    {
                        db.SeedForMeasures();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    db.SeedForConfigOptions();

                }
                MessageBox.Show("Заполнение прошло успешно!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public void PressTheKey(KeyPressEventArgs e)
        {
            OnKeyPress(e);
        }

        protected override void OnKeyPress(KeyPressEventArgs e)
        {
            base.OnKeyPress(e);
        }

        private void FillThreatsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FillThreatsForm form = new FillThreatsForm(this);
            form.Show();
        }

        private void setConfigOptionsToolStripMenu_Click(object sender, EventArgs e)
        {
            using (KPSZIContext db = new KPSZIContext())
            {
                db.SeedForConfigOptions();
            }
            MessageBox.Show("Заполнение прошло успешно!", "Это успех, парень!", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void addSZItoMeasuresToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (KPSZIContext db = new KPSZIContext())
            {
                try
                {
                    db.SeedForMeasures();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                MessageBox.Show("Заполнение прошло успешно!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        internal void FindAndReplace(Microsoft.Office.Interop.Word._Application doc, object findText, object replaceWithText)
        {
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }

        private void downloadRegistryListthrlistxlsxToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                using (var client = new WebClient())
                {
                    client.DownloadFile("https://fstec.ru/component/attachments/download/489", "_reestr_sszi.ods");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Проблемы при загрузке файла _reestr_sszi.ods\n" + ex.Message, "Ошибка загрузки файла", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Excel.Application excel = new Excel.Application();
                Excel.Workbook wb = excel.Workbooks.Open(Environment.CurrentDirectory + "/_reestr_sszi.ods");
                excel.Visible = true;
                wb.SaveAs(Environment.CurrentDirectory + "/_reestr_sszi.xlsx", Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wb.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Проблемы при конвертации файла _reestr_sszi.ods в формат .xlsx\n" + ex.Message, "Ошибка конфертации файла", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            File.Delete(Environment.CurrentDirectory + "/_reestr_sszi.ods");
            MessageBox.Show("Файл успешно загружен и сконвертирован в формат .xlsx", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void rewriteRegistryDBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Перезаписать Реестра ФСТЭК СЗИ
            if (MessageBox.Show("Будут удалены все данные о сертификатах СЗИ в БД и записаны заново.\nПродолжить?", "Ахтунг!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                using (KPSZIContext db = new KPSZIContext())
                {
                    FileInfo fi = new FileInfo("_reestr_sszi.xlsx");

                    db.CertificatesSZI.AddRange(CertificateSZI.GetThreatsFromXlsx(fi, db));
                    db.SaveChanges();
                }
                MessageBox.Show("Таблица сертификатах СЗИ успешно перезаписана!", "Это успех, парень!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}