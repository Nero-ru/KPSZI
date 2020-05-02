using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace KPSZI
{
    class StageAllReport : Stage
    {
        protected override ImageList imageListForTabPage { get; set; }
        protected string sourcePathFIKS;
        protected string sourcePathScanoval;
        protected string sourcePathRevisor2XP;
        protected string sourcePathScannerVS;
        protected string templatePath;
        protected StageReportFIKS FIKS;
        protected StageReportRevisor2XP Revisor2XP;
        protected StageReportScannerVS ScannerVS;
        protected StageReportScanoval Scanoval;

        public StageAllReport(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS, string templatePath, StageReportRevisor2XP revisor2XP, StageReportFIKS fiks, StageReportScannerVS scannerVS, StageReportScanoval scanoval)
            : base(stageTab, stageNode, mainForm, IS)
        {
            this.templatePath = templatePath;
            Revisor2XP = revisor2XP;
            FIKS = fiks;
            ScannerVS = scannerVS;
            Scanoval = scanoval;
        }

        public override void enterTabPage()
        {
            mf.btnGetPath_FIKS.Click += new EventHandler(GetPathReport);
            mf.btnGetPath_Revisor2XP.Click += new EventHandler(GetPathReport);
            mf.btnGetPath_ScannerVS.Click += new EventHandler(GetPathReport);
            mf.btnGetPath_Scanoval.Click += new EventHandler(GetPathReport);
            mf.btnExportAllToWord.Click += new EventHandler(ExportAllReportsToWord);
        }

        public override void saveChanges()
        {
            
        }

        protected override void initTabPage()
        {
            
        }

        protected void GetPathReport(object sender, EventArgs e)
        {
            switch ((sender as Button).Name.Split('_')[1])
            {
                case "FIKS":
                    if (mf.ofdReport.ShowDialog() != DialogResult.Cancel)
                    {
                        sourcePathFIKS = mf.ofdReport.FileName;
                        mf.cbStatus_FIKS.Checked = true;
                    }
                    break;
                case "Revisor2XP":
                    if (mf.ofdReport.ShowDialog() != DialogResult.Cancel)
                    {
                        sourcePathRevisor2XP = mf.ofdReport.FileName;
                        mf.cbStatus_Revisor2XP.Checked = true;
                    }
                    break;
                case "ScannerVS":
                    if (mf.ofdReport.ShowDialog() != DialogResult.Cancel)
                    {
                        sourcePathScannerVS = mf.ofdReport.FileName;
                        mf.cbStatus_ScannerVS.Checked = true;
                    }
                    break;
                case "Scanoval":
                    if (mf.ofdReport.ShowDialog() != DialogResult.Cancel)
                    {
                        sourcePathScanoval = mf.ofdReport.FileName;
                        mf.cbStatus_Scanoval.Checked = true;
                    }
                    break;
            }
        }

        protected void ExportAllReportsToWord(object sender, EventArgs e)
        {
            if (mf.cbStatus_FIKS.Checked != true || mf.cbStatus_Revisor2XP.Checked != true || mf.cbStatus_ScannerVS.Checked != true || mf.cbStatus_Scanoval.Checked != true)
                MessageBox.Show("Выберете все отчеты", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
            {
                Word.Application app = new Word.Application();
                Word.Document doc = app.Documents.Add(Environment.CurrentDirectory + "/" + templatePath);
                Word.Paragraph paragraph = doc.Paragraphs.Add(doc.Bookmarks["AllReport"].Range);

                //Task taskRevisor2XP = Task.Factory.StartNew(() => );
                //Task taskFIKS = Task.Factory.StartNew(() => );
                //Task taskScannerVS = Task.Factory.StartNew(() => );
                //Task taskScanoval = Task.Factory.StartNew(() =>);

                // Привет
                Scanoval.ReportToWord(sourcePathScanoval, templatePath, true, doc, app, paragraph);
                ScannerVS.ReportToWord(sourcePathScannerVS, templatePath, true, doc, app, paragraph);
                FIKS.ReportToWord(sourcePathFIKS, templatePath, true, doc, app, paragraph);
                
                Revisor2XP.ReportToWord(sourcePathRevisor2XP, templatePath, true, doc, app, paragraph);
                
                
                

            }
        }
    }
}
