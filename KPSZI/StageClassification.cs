using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KPSZI.Model;
using System.Drawing;

namespace KPSZI
{
    class StageClassification : Stage
    {
        protected override ImageList imageListForTabPage { get; set; }

        public StageClassification(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS)
            :base(stageTab, stageNode, mainForm, IS)
        {
            
        }

        public List<TabPage> tabPagesInfoTypes = new List<TabPage>();

        protected override void initTabPage()
        {
            using (Model.KPSZIContext db = new KPSZIContext())
            {
                foreach (InfoType it in db.InfoTypes)
                {
                    TabPage tp = new TabPage { Name = it.TypeName, Text = it.TypeName };

                    Label labelConf = new Label { Text = "Конфиденциальность" };
                    labelConf.Location = new Point { X = 10, Y = 15 };
                    labelConf.Width = 200;

                    Label labelInteg = new Label { Text = "Целостность" };
                    labelInteg.Location = new Point { X = 10, Y = 45 };

                    Label labelAvail = new Label { Text = "Доступность" };
                    labelAvail.Location = new Point { X = 10, Y = 75 };


                    ComboBox cbConf = new ComboBox { Name = "cbConf", DropDownStyle = ComboBoxStyle.DropDownList };
                    cbConf.SelectedIndexChanged += GISClassCalculate;
                    cbConf.Location = new Point { X = 135, Y = 7 };
                    cbConf.Items.Add("Высокий");
                    cbConf.Items.Add("Средний");
                    cbConf.Items.Add("Низкий");

                    ComboBox cbInteg = new ComboBox { Name = "cbInteg", DropDownStyle = ComboBoxStyle.DropDownList };
                    cbInteg.SelectedIndexChanged += GISClassCalculate;
                    cbInteg.Location = new Point { X = 135, Y = 37 };
                    cbInteg.Items.Add("Высокий");
                    cbInteg.Items.Add("Средний");
                    cbInteg.Items.Add("Низкий");

                    ComboBox cbAvail = new ComboBox { Name = "cbAvail", DropDownStyle = ComboBoxStyle.DropDownList };
                    cbAvail.SelectedIndexChanged += GISClassCalculate;
                    cbAvail.Location = new Point { X = 135, Y = 67 };
                    cbAvail.Items.Add("Высокий");
                    cbAvail.Items.Add("Средний");
                    cbAvail.Items.Add("Низкий");

                    tp.Controls.AddRange(new Control[] { cbConf, cbAvail, cbInteg, labelAvail, labelConf, labelInteg });
                    tabPagesInfoTypes.Add(tp);
                }
            }

            mf.comboBoxScale.SelectedIndexChanged += new System.EventHandler(GISClassCalculate);
            mf.checkedListBoxCategoryPDN.SelectedIndexChanged += new System.EventHandler(checkedListBoxCategoryPDN_SelectedIndexChanged);
            mf.checkBoxSubjectsStaff.CheckedChanged += new System.EventHandler(checkBoxSubjectsStaff_CheckedChanged);
            mf.comboBoxHundred.SelectedIndexChanged += new System.EventHandler(ISPDNLevelCalculate);
            mf.comboBoxActualThreatsType.SelectedIndexChanged += new System.EventHandler(ISPDNLevelCalculate);
        }
        
        public override void saveChanges()
        {
            
        }

        public override void enterTabPage()
        {
            mf.tabControlInfoTypes.TabPages.Clear();
            if (IS.listOfInfoTypes.Count == 0)
            {
                TabPage startTabPage = new TabPage { Name = "startTabPage", Text = "Виды информации" };
                startTabPage.Controls.Add(new Label { Text = "Выберите виды информации в разделе \"Параметры ИС\" для определения степеней ущерба",
                    Location = new Point { X = 10, Y = 10 }, AutoSize = false, Size = new Size(new Point { X = startTabPage.Width - 30, Y = 100 }) });
                mf.tabControlInfoTypes.TabPages.Add(startTabPage);
            }
            //
            mf.panelPDN.Enabled = false;
            mf.panelPDN.Visible = false;
            foreach (InfoType it in IS.listOfInfoTypes)
            {
                if (it.TypeName == "Персональные данные")
                {
                    mf.panelPDN.Enabled = true;
                    mf.panelPDN.Visible = true;
                }
                mf.tabControlInfoTypes.TabPages.Add(((StageClassification)mf.stages["tnClassification"]).tabPagesInfoTypes.FindLast(t => t.Name == it.TypeName));
            }
            GISClassCalculate(null, null);
        }

        public void GISClassCalculate(object sender, EventArgs e)
        {
            //Определение класса защищенности ГИС
            if (mf.comboBoxScale.SelectedItem == null || (IS.listOfInfoTypes.Count ==1 && IS.listOfInfoTypes.FindLast(t=>t.TypeName=="Персональные данные")!=null) || IS.listOfInfoTypes.Count == 0)
            {
                mf.labelGISClass.Text = "Выберите все поля на форме, чтобы определить класс защищенности ГИС";
                return;
            }

            string scaleGIS = mf.comboBoxScale.SelectedItem.ToString();
            string maxDegreeDamage = "";
            List<string> scales = new List<string>();
            foreach (TabPage tp in mf.tabControlInfoTypes.TabPages)
            {
                foreach (Control con in tp.Controls)
                    if (con is ComboBox)
                    {
                        if (((ComboBox)con).SelectedItem == null)
                        {
                            mf.labelGISClass.Text = "Выберите все поля на форме, чтобы определить класс защищенности ГИС";
                            return;
                        }
                        scales.Add(((ComboBox)con).SelectedItem.ToString());
                    }
            }

            maxDegreeDamage = "Низкий";
            if (scales.Contains("Средний"))
            {
                maxDegreeDamage = "Средний";
            }
            if (scales.Contains("Высокий"))
            {
                maxDegreeDamage = "Высокий";
            }

            switch (maxDegreeDamage)
            {
                case "Высокий":
                    {
                        mf.labelGISClass.Text = "К1 - 1 класс защищенности";
                        IS.GISClass = 1;
                        break;
                    }
                case "Средний":
                    {
                        if (scaleGIS == "Федеральный")
                        {
                            mf.labelGISClass.Text = "К1 - 1 класс защищенности";
                            IS.GISClass = 1;
                        }
                        else
                        {
                            mf.labelGISClass.Text = "К2 - 2 класс защищенности";
                            IS.GISClass = 2;
                        }
                            break;
                    }
                case "Низкий":
                    {
                        if (scaleGIS == "Федеральный")
                        {
                            mf.labelGISClass.Text = "К2 - 2 класс защищенности";
                            IS.GISClass = 2;
                        }
                        else
                        {
                            mf.labelGISClass.Text = "К3 - 3 класс защищенности";
                            IS.GISClass = 3;
                        }
                        break;
                    }
            }

        }

        public void ISPDNLevelCalculate(object sender, EventArgs e)
        {
            if (mf.comboBoxActualThreatsType.SelectedItem == null || mf.comboBoxHundred.SelectedItem == null || IS.listOfCategoriesPDN.Count == 0)
            {
                mf.labelISPDNLevel.Text = "Выберите все поля на форме, чтобы определить уровень защищенности персональных данных";
                
                if (mf.comboBoxActualThreatsType.SelectedItem != null && IS.listOfCategoriesPDN.Count == 1 && IS.listOfCategoriesPDN.Contains("Биометрические"))
                {
                    string actualThreats = mf.comboBoxActualThreatsType.SelectedItem.ToString();
                    switch (actualThreats)
                    {
                        case "1-го типа":
                            {
                                mf.labelISPDNLevel.Text = "Уровень защищенности - 1";
                                break;
                            }
                        case "2-го типа":
                            {
                                mf.labelISPDNLevel.Text = "Уровень защищенности - 2";
                                break;
                            }
                        case "3-го типа":
                            {
                                mf.labelISPDNLevel.Text = "Уровень защищенности - 3";
                                break;
                            }
                    }
                }

                return;
            }

            string ActualThreats = mf.comboBoxActualThreatsType.SelectedItem.ToString();
            int type;
            int.TryParse(ActualThreats.Substring(0, 1),out type);
            IS.typeOfActualThreats = type;
            bool isStaffSubjects = mf.checkBoxSubjectsStaff.Checked;
            string SubjectsPDN = mf.comboBoxHundred.SelectedItem.ToString();

            

            List<int> levels = new List<int>();
            switch (ActualThreats)
            {
                case "1-го типа":
                    {
                        if (IS.listOfCategoriesPDN.Contains("Специальные") || IS.listOfCategoriesPDN.Contains("Биометрические") || IS.listOfCategoriesPDN.Contains("Иные"))
                            levels.Add(1);
                        if (IS.listOfCategoriesPDN.Contains("Общедоступные"))
                            levels.Add(2);
                        break;
                    }

                case "2-го типа":
                    {
                        if (IS.listOfCategoriesPDN.Contains("Специальные") && mf.checkBoxSubjectsStaff.Checked == false && SubjectsPDN == "Более 100,000")
                            levels.Add(1);
                        if (IS.listOfCategoriesPDN.Contains("Специальные"))
                            if ((isStaffSubjects == false && SubjectsPDN == "Менее 100,000") || isStaffSubjects == true )
                                levels.Add(2);
                        if (IS.listOfCategoriesPDN.Contains("Биометрические"))
                            levels.Add(2);

                        if (IS.listOfCategoriesPDN.Contains("Иные") && mf.checkBoxSubjectsStaff.Checked == false && SubjectsPDN == "Более 100,000")
                            levels.Add(2);
                        if (IS.listOfCategoriesPDN.Contains("Иные"))
                            if ((isStaffSubjects == false && SubjectsPDN == "Менее 100,000") || isStaffSubjects == true)
                                levels.Add(3);

                        if (IS.listOfCategoriesPDN.Contains("Общедоступные") && mf.checkBoxSubjectsStaff.Checked == false && SubjectsPDN == "Более 100,000")
                            levels.Add(2);
                        if (IS.listOfCategoriesPDN.Contains("Общедоступные"))
                            if ((isStaffSubjects == false && SubjectsPDN == "Менее 100,000") || isStaffSubjects == true)
                                levels.Add(3);

                        break;
                    }

                case "3-го типа":
                    {
                        if (IS.listOfCategoriesPDN.Contains("Общедоступные"))
                            levels.Add(4);
                        if (IS.listOfCategoriesPDN.Contains("Биометрические"))
                            levels.Add(3);

                        if (IS.listOfCategoriesPDN.Contains("Специальные") && mf.checkBoxSubjectsStaff.Checked == false && SubjectsPDN == "Более 100,000")
                            levels.Add(2);
                        if (IS.listOfCategoriesPDN.Contains("Специальные"))
                            if ((isStaffSubjects == false && SubjectsPDN == "Менее 100,000") || isStaffSubjects == true)
                                levels.Add(3);

                        if (IS.listOfCategoriesPDN.Contains("Иные") && mf.checkBoxSubjectsStaff.Checked == false && SubjectsPDN == "Более 100,000")
                            levels.Add(3);
                        if (IS.listOfCategoriesPDN.Contains("Иные"))
                            if ((isStaffSubjects == false && SubjectsPDN == "Менее 100,000") || isStaffSubjects == true)
                                levels.Add(4);

                        break;
                    }
            }
            IS.levelOfPDN = levels.Min();
            mf.labelISPDNLevel.Text = "Уровень защищенности - " + IS.levelOfPDN;
        }
        
        #region Обработчики
        public void checkBoxSubjectsStaff_CheckedChanged(object sender, EventArgs e)
        {
            if (mf.checkBoxSubjectsStaff.Checked)
            {
                mf.comboBoxHundred.Items.Clear();
                mf.comboBoxHundred.Items.Add("Не требуется");
                mf.comboBoxHundred.SelectedIndex = 0;
                mf.comboBoxHundred.Enabled = false;
            }

            if (mf.checkBoxSubjectsStaff.Checked == false)
            {
                mf.comboBoxHundred.Items.Clear();
                mf.comboBoxHundred.Enabled = true;
                mf.comboBoxHundred.Items.Add("Менее 100,000");
                mf.comboBoxHundred.Items.Add("Более 100,000");
            }
            ISPDNLevelCalculate(null, null);
        }

        private void lbInfoTypes_SelectedIndexChanged(object sender, EventArgs e)
        {
            // При нажатии на галочку все выбранные 
            // виды информации помещаются в экземпляр ИС
            IS.ISName = mf.tbISName.Text;
            object[] buf = new object[mf.lbInfoTypes.CheckedItems.Count];
            IS.listOfInfoTypes.Clear();
            mf.lbInfoTypes.CheckedItems.CopyTo(buf, 0);
            for (int i = 0; i < buf.Length; i++)
            {
                IS.listOfInfoTypes.Add((InfoType)buf.GetValue(i));
            }
        }

        private void checkedListBoxCategoryPDN_SelectedIndexChanged(object sender, EventArgs e)
        {
            object[] buf = new object[mf.checkedListBoxCategoryPDN.CheckedItems.Count];
            IS.listOfCategoriesPDN.Clear();
            mf.checkedListBoxCategoryPDN.CheckedItems.CopyTo(buf, 0);
            for (int i = 0; i < buf.Length; i++)
            {
                IS.listOfCategoriesPDN.Add((string)buf.GetValue(i));
            }
            ISPDNLevelCalculate(null, null);
        }
        #endregion
    }
}
