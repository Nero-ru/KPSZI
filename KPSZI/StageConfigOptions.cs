using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KPSZI.Model;

namespace KPSZI
{
    class StageConfigOptions:Stage
    {
        public StageConfigOptions(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS) : base(stageTab, stageNode, mainForm, IS)
        {
        }

        protected override ImageList imageListForTabPage
        {
            get;
            set;
        }
        
        public override void enterTabPage()
        {
            if (IS.GISClass == 0 || IS.listOfAllNSDMeasures.Count == 0)
            {
                MessageBox.Show("Определите класс информационной системы, а также уточненный адаптированный базовый перечень мер защиты информации для определения требований к параметрам настройки.");
                return;
            }
            mf.wsm.Visible = true;
            mf.tabControl.TabPages["tpConfigOptions"].Update();
            mf.wsm.Update();

            mf.dgvConfigNMeasures.Rows.Clear();

            using (KPSZIContext db = new KPSZIContext())
            {
                var tempGM = db.GisMeasures.ToList().Intersect(IS.listOfAllNSDMeasures).ToList();
                foreach (GISMeasure gm in tempGM)
                {
                    DataGridViewRow dgr = new DataGridViewRow();
                    dgr.CreateCells(mf.dgvConfigNMeasures);
                    dgr.Cells[0].Value = gm.MeasureGroup.ShortName + "." + gm.Number + " " + gm.Description;

                    for(int i = gm.ConfigOptions.Count-1; i >=0;i--)
                    {
                        if (gm.ConfigOptions.ElementAt(i).Description != null)
                        {
                            if (gm.ConfigOptions.ElementAt(i).Description.Contains("Усиление"))
                            {
                                if (isDefenceClassConstraint(gm.ConfigOptions.ElementAt(i).DefenceClass, IS.GISClass))
                                    dgr.Cells[1].Value += gm.ConfigOptions.ElementAt(i).Name + " (" + gm.ConfigOptions.ElementAt(i).Description + ")\n";
                            }
                            else
                            {
                                string descr = "";
                                descr = gm.ConfigOptions.ElementAt(i).Description == null ? "" : " (" + gm.ConfigOptions.ElementAt(i).Description + ")";
                                dgr.Cells[1].Value += gm.ConfigOptions.ElementAt(i).Name + descr+" \n";
                            }
                        }
                        else
                            dgr.Cells[1].Value += gm.ConfigOptions.ElementAt(i).Name + "\n";
                    }
                    if (dgr.Cells[1].Value != null)
                    {
                        mf.dgvConfigNMeasures.Rows.Add(dgr);
                    }
                }
            }
            mf.wsm.Visible = false;
        }

        public bool isDefenceClassConstraint(string configOptionClasses, int defenceClassIS)
        {
            if (configOptionClasses == "0")
                return true;
            string[] classes = configOptionClasses.Split(',');
            for (int i = 0; i < classes.Length; i++)
            {
                if (classes[i] == defenceClassIS.ToString())
                    return true;
            }
            return false;
        }


        public override void saveChanges()
        {

        }

        protected override void initTabPage()
        {
            mf.btnExportConfigOptions.Click += new EventHandler(BtnExportWord_CLick);
        }

        public void BtnExportWord_CLick(object sender, EventArgs e)
        {
            mf.wsm.Visible = true;
            mf.wsm.Update();

            Microsoft.Office.Interop.Word._Application wordApp = null;
            try
            {
                wordApp = new Microsoft.Office.Interop.Word.Application();
            }
            catch
            {
                mf.wsm.Visible = false;
                MessageBox.Show("На ПК не установлен пакет Microsoft Office Word 2007 или позднее. Экспорт невозможен.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Microsoft.Office.Interop.Word.Document wordDoc;
            Microsoft.Office.Interop.Word.Paragraph wordParag;


            wordDoc = wordApp.Documents.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            wordParag = wordDoc.Paragraphs.Add();
            wordParag.Range.Text = "Требования к параметрам настройки средств защиты информации в ИС \"" + IS.ISName+"\"";
            wordParag.Range.Font.Name = "Times New Roman";
            wordParag.Range.Font.Size = 18;
            wordParag.Range.Font.Bold = 1;
            wordParag.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wordParag.Range.InsertParagraphAfter(); wordParag.Range.InsertParagraphAfter();
            using (KPSZIContext db = new KPSZIContext())
            {
                foreach (MeasureGroup mg in db.MeasureGroups.ToList())
                {

                    //Группа мер
                    bool isMeasureGroupInTable = false;
                    for (int i = mf.dgvConfigNMeasures.Rows.Count-1; i >= 0; i--)
                    {
                        if (mf.dgvConfigNMeasures.Rows[i].Cells[0].Value.ToString().Contains(mg.ShortName + "."))
                        {
                            isMeasureGroupInTable = true;
                        }
                    }
                    if (!isMeasureGroupInTable)
                        continue;
                    wordParag.Range.Text = '\t'+mg.Name;
                    wordParag.Range.Font.Name = "Times New Roman";
                    wordParag.Range.Font.Size = 16;
                    wordParag.Range.Font.Bold = 1;
                    wordParag.Range.Font.Italic = 0;
                    wordParag.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;

                    wordParag.Range.InsertParagraphAfter();

                    foreach (DataGridViewRow dgvr in mf.dgvConfigNMeasures.Rows)
                    {
                        if (dgvr.Cells[0].Value.ToString().Contains(mg.ShortName + "."))
                        {
                            //Мера                            
                            wordParag.Range.Text = '\t' + dgvr.Cells[0].Value.ToString();
                            wordParag.Range.Font.Name = "Times New Roman";
                            wordParag.Range.Font.Size = 14;
                            wordParag.Range.Font.Bold = 1;
                            wordParag.Range.Font.Italic = 0;
                            wordParag.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;

                            wordParag.Range.InsertParagraphAfter();

                            //тип сзи
                            string szis = "";
                            string measure = dgvr.Cells[0].Value.ToString();
                            List<SZISort> szisorts = db.GisMeasures.ToList().Where(m => (m.MeasureGroup.ShortName + "." + m.Number + " " + m.Description) == measure).First().SZISorts.ToList();

                            var listOfSZIs = db.SZIs.ToList().Intersect(IS.listOfSZIs).ToList();
                            foreach(SZI sz in listOfSZIs)
                            {
                                SZI m = db.SZIs.Where(t => t.SZIId == sz.SZIId).First();
                                sz.SZISorts = m.SZISorts.ToList();
                            }
                            foreach (SZISort s in szisorts)
                            {
                                szis += s.Name  + ": "+listOfSZIs.Where(szi => szi.SZISorts.Contains(s)).First().Name + "; ";
                            }
                            if (szis != "")
                            {
                                szis = szis.Substring(0, szis.Length - 2) + ".";

                                wordParag.Range.Text = '\t' + szis;
                                wordParag.Range.Font.Name = "Times New Roman";
                                wordParag.Range.Font.Size = 14;
                                wordParag.Range.Font.Bold = 0;
                                wordParag.Range.Font.Italic = 1;
                                wordParag.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;

                                wordParag.Range.InsertParagraphAfter();
                            }
                            

                            string[] configOptions = dgvr.Cells[1].Value.ToString().Split('\n');
                            for (int i = 0 ; i < configOptions.Length-1; i++)
                            {
                                //Параметры
                                char divider = i == configOptions.Length -2 ? '.' : ';';
                                wordParag.Range.Text = "\t– "+configOptions[i] + divider;
                                wordParag.Range.Font.Name = "Times New Roman";
                                wordParag.Range.Font.Size = 12;
                                wordParag.Range.Font.Bold = 0;
                                wordParag.Range.Font.Italic = 0;
                                wordParag.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                                wordParag.Range.InsertParagraphAfter();
                            }
                            wordParag.Range.InsertParagraphAfter();
                        }
                        else
                            continue;                        
                    }
                    wordParag.Range.InsertParagraphAfter();
                }
                wordApp.Visible = true;

                mf.wsm.Visible = false;
            }
        }
    }
}
