using DocumentFormat.OpenXml.Wordprocessing;
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
                MessageBox.Show("Выберите все отчеты", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
            {
                Word.Application app = new Word.Application();
                Word.Document doc = app.Documents.Add(Environment.CurrentDirectory + "/" + templatePath);
                Word.Paragraph paragraph = doc.Paragraphs.Add(doc.Bookmarks["AllReport"].Range);

                Word.Bookmark bkNameofIS = doc.Bookmarks.Add("ISName", doc.Bookmarks["ISName"].Range);
                Word.Bookmark bkNameofIS1 = doc.Bookmarks.Add("ISName1", doc.Bookmarks["ISName1"].Range);
                Word.Bookmark bkNameofIS2 = doc.Bookmarks.Add("ISName2", doc.Bookmarks["ISName2"].Range);
                Word.Bookmark bkNameofIS3 = doc.Bookmarks.Add("ISName3", doc.Bookmarks["ISName3"].Range);
                Word.Bookmark bkNameofIS4 = doc.Bookmarks.Add("ISName4", doc.Bookmarks["ISName4"].Range);
                Word.Bookmark bkNameofIS5 = doc.Bookmarks.Add("ISName5", doc.Bookmarks["ISName5"].Range);
                Word.Bookmark bkNameofIS6 = doc.Bookmarks.Add("ISName6", doc.Bookmarks["ISName6"].Range);
                Word.Bookmark bkNameofIS7 = doc.Bookmarks.Add("ISName7", doc.Bookmarks["ISName7"].Range);
                Word.Bookmark bkNameofIS8 = doc.Bookmarks.Add("ISName8", doc.Bookmarks["ISName8"].Range);
                Word.Bookmark bkNameofIS9 = doc.Bookmarks.Add("ISName9", doc.Bookmarks["ISName9"].Range);
                Word.Bookmark bkNameofIS10 = doc.Bookmarks.Add("ISName10", doc.Bookmarks["ISName10"].Range);
                Word.Bookmark bkNameofIS11 = doc.Bookmarks.Add("ISName11", doc.Bookmarks["ISName11"].Range);
                Word.Bookmark bkDateReport = doc.Bookmarks.Add("DateReport", doc.Bookmarks["DateReport"].Range);
                Word.Bookmark bkNumOfWP = doc.Bookmarks.Add("NumOfWP", doc.Bookmarks["NumOfWP"].Range);
                Word.Bookmark bkNumsOfCabinets = doc.Bookmarks.Add("NumsOfCabinets", doc.Bookmarks["NumsOfCabinets"].Range);
                Word.Bookmark bkAddressOfComp = doc.Bookmarks.Add("AddressOfComp", doc.Bookmarks["AddressOfComp"].Range);

                bkNameofIS.Range.Text = Convert.ToString(mf.NameOfISTextBox.Text);
                bkNameofIS1.Range.Text = Convert.ToString(mf.NameOfISTextBox.Text);
                bkNameofIS2.Range.Text = Convert.ToString(mf.NameOfISTextBox.Text);
                bkNameofIS3.Range.Text = Convert.ToString(mf.NameOfISTextBox.Text);
                bkNameofIS4.Range.Text = Convert.ToString(mf.NameOfISTextBox.Text);
                bkNameofIS5.Range.Text = Convert.ToString(mf.NameOfISTextBox.Text);
                bkNameofIS6.Range.Text = Convert.ToString(mf.NameOfISTextBox.Text);
                bkNameofIS7.Range.Text = Convert.ToString(mf.NameOfISTextBox.Text);
                bkNameofIS8.Range.Text = Convert.ToString(mf.NameOfISTextBox.Text);
                bkNameofIS9.Range.Text = Convert.ToString(mf.NameOfISTextBox.Text);
                bkNameofIS10.Range.Text = Convert.ToString(mf.NameOfISTextBox.Text);
                bkNameofIS11.Range.Text = Convert.ToString(mf.NameOfISTextBox.Text);
                //bkDateReport.Range.Text = Convert.ToString(DateTime.Now.ToString("D"));
                bkDateReport.Range.Text = Convert.ToString(mf.DateOfReportTextBox.Text + " ");
                bkNumOfWP.Range.Text = Convert.ToString(mf.NumOfWPTextBox.Text);
                bkNumsOfCabinets.Range.Text = Convert.ToString(mf.NumsOfCabsTextBox.Text + " ");
                bkAddressOfComp.Range.Text = Convert.ToString(mf.AddressOfCompTextBox.Text);

                //Task taskRevisor2XP = Task.Factory.StartNew(() => );
                //Task taskFIKS = Task.Factory.StartNew(() => );
                //Task taskScannerVS = Task.Factory.StartNew(() => );
                //Task taskScanoval = Task.Factory.StartNew(() =>);

                Scanoval.ReportToWord(sourcePathScanoval, templatePath, true, doc, app, paragraph);
                ScannerVS.ReportToWord(sourcePathScannerVS, templatePath, true, doc, app, paragraph);
                FIKS.ReportToWord(sourcePathFIKS, templatePath, true, doc, app, paragraph);
                
                Revisor2XP.ReportToWord(sourcePathRevisor2XP, templatePath, true, doc, app, paragraph);
                
                
                

            }
        }
    }
}
