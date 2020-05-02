using KPSZI.Model;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPSZI
{
    class StageActualThreats : Stage
    {
        public List<Threat> listThreats;
        List<Vulnerability> listVulnerabilities;
        List<SFH> listSFHs;
        List<SFHType> listSFHTypes;
        List<ImplementWay> listImplementWays;
        List<ThreatSource> listSources;
        List<Threat> listFilteredThreats;
        Color[] headerColors = new Color[4] { Color.LemonChiffon, Color.PowderBlue, Color.LightGreen, Color.Plum };
        Dictionary<int, int[,,]> damageDegreeInput;
        DamageDegreeControl DDControl;
        bool firstEnter = true;

        protected override ImageList imageListForTabPage { get; set; }

        public StageActualThreats(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS)
            : base(stageTab, stageNode, mainForm, IS)
        {
            
            
        }

        protected override void initTabPage()
        {
            using (KPSZIContext db = new KPSZIContext())
            {
                //Инициализация списка угроз
                listThreats = db.Threats.OrderBy(t => t.ThreatNumber).ToList();
                foreach (Threat threat in listThreats)
                {
                    threat.Vulnerabilities = db.Threats.Where(t1 => t1.ThreatNumber == threat.ThreatNumber).First().Vulnerabilities;
                    threat.SFHs = db.Threats.Where(t2 => t2.ThreatNumber == threat.ThreatNumber).First().SFHs;
                    threat.ImplementWays = db.Threats.Where(t3 => t3.ThreatNumber == threat.ThreatNumber).First().ImplementWays;
                    threat.ThreatSources = db.Threats.Where(t4 => t4.ThreatNumber == threat.ThreatNumber).First().ThreatSources;
                    threat.GISMeasures = db.Threats.Where(t5 => t5.ThreatNumber == threat.ThreatNumber).First().GISMeasures;
                    threat.setStringVulnerabilities();
                    threat.setStringSFHs();
                    threat.setStringImplementWays();
                    threat.setStringSources();
                }

                listVulnerabilities = db.Vulnerabilities.OrderBy(v => v.VulnerabilityNumber).ToList();
                listSFHs = db.SFHs.OrderBy(s => s.SFHNumber).ToList();
                listSFHTypes = db.SFHTypes.OrderBy(st => st.SFHTypeId).ToList();
                listImplementWays = db.ImplementWays.OrderBy(w => w.WayNumber).ToList();
                listSources = db.ThreatSources.OrderBy(so => so.ThreatSourceId).ToList();
                foreach (Vulnerability vul in listVulnerabilities)
                    vul.Threats = db.Vulnerabilities.Where(v1 => v1.VulnerabilityNumber == vul.VulnerabilityNumber).First().Threats;
                foreach (SFH sfh in listSFHs)
                    sfh.Threats = db.SFHs.Where(s1 => s1.SFHNumber == sfh.SFHNumber).First().Threats;
                foreach (ImplementWay iw in listImplementWays)
                    iw.Threats = db.ImplementWays.Where(w1 => w1.WayNumber == iw.WayNumber).First().Threats;
                foreach (ThreatSource ts in listSources)
                    ts.Threats = db.ThreatSources.Where(so1 => so1.ThreatSourceId == ts.ThreatSourceId).First().Threats;
                listFilteredThreats = new List<Threat>();
            }

            mf.dgvThreats.DefaultCellStyle.SelectionBackColor = Color.AliceBlue;
            mf.dgvThreats.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Black;

            mf.tcThreatsNSD.TabPages.Remove(mf.tpThreatsNSD2);
            mf.tcThreatsNSD.TabPages.Remove(mf.tpThreatsNSD3);

            mf.dgvThreats.DataSource = listThreats;
            mf.dgvThreats.Columns["ThreatID"].Visible = false;
            mf.dgvThreats.Columns["ThreatSources"].Visible = false;
            mf.dgvThreats.Columns["DateOfChange"].Visible = false;
            mf.dgvThreats.Columns["DateOfAdd"].Visible = false;
            mf.dgvThreats.Columns["ImplementWays"].Visible = false;
            mf.dgvThreats.Columns["SFHs"].Visible = false;
            mf.dgvThreats.Columns["Vulnerabilities"].Visible = false;
            mf.dgvThreats.Columns["Description"].Visible = false;
            mf.dgvThreats.Columns["ObjectOfInfluence"].Visible = false;
            mf.dgvThreats.Columns["GISMeasures"].Visible = false;

            mf.dgvThreats.Columns["ThreatNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            mf.dgvThreats.Columns["ConfidenceViolation"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            mf.dgvThreats.Columns["IntegrityViolation"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            mf.dgvThreats.Columns["AvailabilityViolation"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;

            mf.dgvThreats.Columns["ThreatID"].DisplayIndex = 0;
            mf.dgvThreats.Columns["ThreatNumber"].Width = 60;
            mf.dgvThreats.Columns["ThreatNumber"].DisplayIndex = 1;
            mf.dgvThreats.Columns["ThreatNumber"].HeaderText = "№ УБИ";
            mf.dgvThreats.Columns["Name"].HeaderText = "Название УБИ";
            mf.dgvThreats.Columns["Name"].DisplayIndex = 2;
            mf.dgvThreats.Columns["ConfidenceViolation"].Width = 30;
            mf.dgvThreats.Columns["ConfidenceViolation"].HeaderText = "К";
            mf.dgvThreats.Columns["ConfidenceViolation"].DisplayIndex = 3;
            mf.dgvThreats.Columns["IntegrityViolation"].Width = 30;
            mf.dgvThreats.Columns["IntegrityViolation"].HeaderText = "Ц";
            mf.dgvThreats.Columns["IntegrityViolation"].DisplayIndex = 4;
            mf.dgvThreats.Columns["AvailabilityViolation"].Width = 30;
            mf.dgvThreats.Columns["AvailabilityViolation"].HeaderText = "Д";
            mf.dgvThreats.Columns["AvailabilityViolation"].DisplayIndex = 5;
            mf.dgvThreats.Columns["stringVuls"].HeaderText = "Уязвимости";
            mf.dgvThreats.Columns["stringVuls"].DisplayIndex = 6;
            mf.dgvThreats.Columns["stringWays"].HeaderText = "Способы реализации УБИ";
            mf.dgvThreats.Columns["stringWays"].DisplayIndex = 7;
            mf.dgvThreats.Columns["stringSFHS"].HeaderText = "СФХ";
            mf.dgvThreats.Columns["stringSFHS"].DisplayIndex = 8;
            mf.dgvThreats.Columns["stringSources"].HeaderText = "Источник угрозы";
            mf.dgvThreats.Columns["stringSources"].DisplayIndex = 9;

            mf.dgvThreats.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            mf.dgvThreats.SelectionChanged += new System.EventHandler(dgvThreats_SelectionChanged);
            mf.tpActualThreats.Resize += new System.EventHandler(tpActualThreats_Resize);
            mf.clbThreatFilter.SelectedIndexChanged += new System.EventHandler(clbThreatFilter_SelectedIndexChanged);
            mf.tcThreatsNSD.SelectedIndexChanged += new System.EventHandler(tcThreatsNSD_SelectedIndexChanged);
            mf.btnGotoDamage.Click += new System.EventHandler(btnGotoDamage_Click);
            mf.dgvActualThreatsNSD.SelectionChanged += new System.EventHandler(dgvActualThreats_SelectionChanged);
            mf.btnReady.Click += new System.EventHandler(btnReady_Click);
            mf.btnExportModelThreats.Click += new System.EventHandler(btnExportModelThreats_Click);

            mf.clbThreatFilter.SetItemChecked(0, false);
            mf.clbThreatFilter.SetItemChecked(1, false);
            mf.clbThreatFilter.SetItemChecked(2, false);
            mf.clbThreatFilter.SetItemChecked(3, false);
            filterThreatList();
        }

        public void initTabPageThreatsNSD2()
        {
            // дизайн DataGridView для определения актуальных УБИ
            mf.dgvActualThreatsNSD.DefaultCellStyle.SelectionBackColor = Color.AliceBlue;
            mf.dgvThreats.DefaultCellStyle.SelectionForeColor = Color.Black;
            mf.dgvActualThreatsNSD.Columns.Clear();
            mf.dgvActualThreatsNSD.Rows.Clear();
            
            mf.dgvActualThreatsNSD.Columns.Add("ThreatNumber", "№");
            mf.dgvActualThreatsNSD.Columns["ThreatNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            mf.dgvActualThreatsNSD.Columns["ThreatNumber"].Width = 40;

            mf.dgvActualThreatsNSD.Columns.Add("Name", "Название УБИ");

            mf.dgvActualThreatsNSD.Columns.Add("Feasibility", "Возможность реализации УБИ");
            mf.dgvActualThreatsNSD.Columns["Feasibility"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            mf.dgvActualThreatsNSD.Columns["Feasibility"].Width = 115;

            mf.dgvActualThreatsNSD.Columns.Add("DamageDegree", "Степень ущерба от реализации УБИ");
            mf.dgvActualThreatsNSD.Columns["DamageDegree"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            mf.dgvActualThreatsNSD.Columns["DamageDegree"].Width = 115;

            mf.dgvActualThreatsNSD.Columns.Add("IsActual", "Актуальность");
            mf.dgvActualThreatsNSD.Columns["IsActual"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            mf.dgvActualThreatsNSD.Columns["IsActual"].Width = 95;

            int i = 0;
            foreach (Threat t in listFilteredThreats)
            {
                mf.dgvActualThreatsNSD.Rows.Add();
                mf.dgvActualThreatsNSD.Rows[i].Cells[mf.dgvActualThreatsNSD.Columns["ThreatNumber"].Index].Value = t.ThreatNumber;
                mf.dgvActualThreatsNSD.Rows[i].Cells[mf.dgvActualThreatsNSD.Columns["Name"].Index].Value = t.Name;
                i++;
            }
            
            calcFeasibility();

            damageDegreeInput = new Dictionary<int, int[,,]>();
            foreach (DataGridViewRow dgvRow in mf.dgvActualThreatsNSD.Rows)
                damageDegreeInput.Add((int)dgvRow.Cells[mf.dgvActualThreatsNSD.Columns["ThreatNumber"].Index].Value,
                    new int[IS.listOfInfoTypes.Count, 3, 7]);

            DDControl = new DamageDegreeControl(IS.listOfInfoTypes, listFilteredThreats, damageDegreeInput, mf);
            DDControl.Location = new System.Drawing.Point(mf.tpThreatsNSD2.Width - DDControl.Width, 0);
            DDControl.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
            mf.tpThreatsNSD2.Controls.Add(DDControl);   

            mf.tcThreatsNSD.SelectedTab = mf.tpThreatsNSD2;
        }

        public void initTabPageThreatsNSD3()
        {
            // дизайн DataGridView для вывода перечня актуальных УБИ
            mf.dgvFinalNSDThreats.Columns.Clear();
            mf.dgvFinalNSDThreats.Rows.Clear();
            mf.dgvFinalNSDThreats.DefaultCellStyle.SelectionBackColor = Color.AliceBlue;
            mf.dgvFinalNSDThreats.DefaultCellStyle.SelectionForeColor = Color.Black;
                mf.dgvFinalNSDThreats.Columns.Add("ThreatNumber", "№");
            mf.dgvFinalNSDThreats.Columns["ThreatNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            mf.dgvFinalNSDThreats.Columns["ThreatNumber"].Width = 40;

            mf.dgvFinalNSDThreats.Columns.Add("Name", "Название УБИ");

            // Заполнение Dgv
            int i = 0;
            IS.listOfActualNSDThreats.Clear();
            foreach (DataGridViewRow row in mf.dgvActualThreatsNSD.Rows)
            {
                int indexActuality = mf.dgvActualThreatsNSD.Columns["IsActual"].Index;
                if (row.Cells[indexActuality].Value.ToString() != "Актуальная") continue;
                int indexNumber = mf.dgvFinalNSDThreats.Columns["ThreatNumber"].Index;

                mf.dgvFinalNSDThreats.Rows.Add();
                mf.dgvFinalNSDThreats.Rows[i].Cells[mf.dgvFinalNSDThreats.Columns["ThreatNumber"].Index].Value = 
                    row.Cells[mf.dgvActualThreatsNSD.Columns["ThreatNumber"].Index].Value.ToString();
                mf.dgvFinalNSDThreats.Rows[i].Cells[mf.dgvFinalNSDThreats.Columns["Name"].Index].Value =
                    row.Cells[mf.dgvActualThreatsNSD.Columns["Name"].Index].Value.ToString();

                int tNumber = Convert.ToInt32(row.Cells[mf.dgvActualThreatsNSD.Columns["ThreatNumber"].Index].Value.ToString());
                
                IS.listOfActualNSDThreats.Add(listThreats.Where(t => t.ThreatNumber == tNumber).FirstOrDefault());
                i++;
            }

            mf.tcThreatsNSD.SelectedTab = mf.tpThreatsNSD3;
        }

        public void calcFeasibility()
        {
            foreach (Threat threat in listFilteredThreats)
            {
                int maxPotencial = 0;
                int feasibility = 2;
                List<ThreatSource> tsList = threat.ThreatSources.ToList();
                foreach (ThreatSource ts in tsList)
                    if (ts.Potencial != 3 && ts.Potencial > maxPotencial) maxPotencial = ts.Potencial;

                // implementPossibility: 0 - низкий, 1 - средний, 2 - высокий
                if (IS.ProjectSecutiryLvl == 2 && maxPotencial == 0)
                    feasibility = 0;
                if ((IS.ProjectSecutiryLvl == 1 && maxPotencial == 0) || (IS.ProjectSecutiryLvl == 2 && maxPotencial == 1))
                    feasibility = 1;
                
                foreach (DataGridViewRow row in mf.dgvActualThreatsNSD.Rows)
                {
                    int indexNumber = mf.dgvActualThreatsNSD.Columns["ThreatNumber"].Index;
                    int indexDamage = mf.dgvActualThreatsNSD.Columns["DamageDegree"].Index;
                    int indexActuality = mf.dgvActualThreatsNSD.Columns["IsActual"].Index;
                    int indexFeasibility = mf.dgvActualThreatsNSD.Columns["Feasibility"].Index;
                    if (row.Cells[indexNumber].Value.ToString() == threat.ThreatNumber.ToString())
                        switch (feasibility)
                        {
                            case 0:
                                row.Cells[indexFeasibility].Value = "Низкая";
                                break;
                            case 1:
                                row.Cells[indexFeasibility].Value = "Средняя";
                                break;
                            case 2:
                                row.Cells[indexFeasibility].Value = "Высокая";
                                row.Cells[indexDamage].Value = "Не определена";
                                row.Cells[indexActuality].Value = "Актуальная";
                                row.DefaultCellStyle.BackColor = Color.LightGreen;
                                break;                            
                        }
                }
            }
        }

        
        private void btnReady_Click(object sender, EventArgs e)
        {
            int k = 0;
            foreach (DataGridViewRow row in mf.dgvActualThreatsNSD.Rows)
            {
                int indexActuality = mf.dgvActualThreatsNSD.Columns["IsActual"].Index;
                if (row.Cells[indexActuality].Value != null && row.Cells[indexActuality].Value.ToString() != "")
                    k++;
            }

            if (k == mf.dgvActualThreatsNSD.Rows.Count)
            {
                if (!mf.tcThreatsNSD.TabPages.Contains(mf.tpThreatsNSD3))
                {
                    mf.tcThreatsNSD.TabPages.Add(mf.tpThreatsNSD3);
                    initTabPageThreatsNSD3();
                }
                else
                if (MessageBox.Show("Обновить итоговый список актуальных угроз?", "Требуется подтверждение!", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    initTabPageThreatsNSD3();
                }
            }
            else
                MessageBox.Show("Актуальность не определена для всех угроз. Укажите степень ущерба.", "Внимание", MessageBoxButtons.OK);
        }

        public override void enterTabPage()
        {
            mf.lblThreatsCount.Text = "Кол-во УБИ: " + mf.dgvThreats.RowCount;
            
            if (mf.tcThreatsNSD.SelectedTab == mf.tpThreatsNSD1)
            {
                filterThreatList();
                checkInputData();
            }
               
            mf.dgvThreats.ClearSelection();
        }

        public void checkInputData()
        {
            if (IS.listOfSources.Count < 2)
            {
                if (MessageBox.Show("Определите характерные для ИС виды нарушителя для получения итогового перечня УБИ. Перейти к выбору?", "Недостаточно исходных данных", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    mf.treeView.SelectedNode = mf.returnTreeNode("tnIntruder");
                return;
            }
            if (IS.listOfImplementWays.Count == 0)
            {
                if (MessageBox.Show("Определите характерные для ИС способы реализации угроз для получения итогового перечня УБИ. Перейти к выбору?", "Недостаточно исходных данных", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    mf.treeView.SelectedNode = mf.returnTreeNode("tnIntruder");
                return;
            }
            if (IS.listOfVulnerabilities.Count == 0)
            {
                if (MessageBox.Show("Определите характерные для ИС потенциальные уязвимости для получения итогового перечня УБИ. Перейти к выбору?", "Недостаточно исходных данных", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    mf.treeView.SelectedNode = mf.returnTreeNode("tnVulnerabilities");
                return;
            }
            if (IS.listOfSFHs.Count == 0)
            {
                if (MessageBox.Show("Определите структурно-функциональные характеристики ИС для получения итогового перечня УБИ. Перейти к выбору?", "Недостаточно исходных данных", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    mf.treeView.SelectedNode = mf.returnTreeNode("tnOptions");
                return;
            }
        }

        public override void saveChanges()
        {
        }

        private void filterThreatList()
        {
            List<Threat> listThreatsByVul = new List<Threat>();
            List<Threat> listThreatsBySFH = new List<Threat>();
            List<Threat> listThreatsByWay = new List<Threat>();
            List<Threat> listThreatsBySource = new List<Threat>();

            // фильтрация УБИ по источнику угрозы
            if (mf.clbThreatFilter.GetItemCheckState(0) == CheckState.Checked)
                foreach (ThreatSource ts in IS.listOfSources)
                    listThreatsBySource = unionLists(listThreatsBySource, listSources.Where(lts => lts.ThreatSourceId == ts.ThreatSourceId).First().Threats.ToList());
            else
                listThreatsBySource.AddRange(listThreats);

            // фильтрация УБИ по способам реализации
            if (mf.clbThreatFilter.GetItemCheckState(1) == CheckState.Checked)
                foreach (ImplementWay iw in IS.listOfImplementWays)
                    listThreatsByWay = unionLists(listThreatsByWay, listImplementWays.Where(way => way.WayNumber == iw.WayNumber).First().Threats.ToList());
            else
                listThreatsByWay.AddRange(listThreats);

            // фильтрация УБИ по уязвимостям
            if (mf.clbThreatFilter.GetItemCheckState(2) == CheckState.Checked)
                foreach (Vulnerability vul in IS.listOfVulnerabilities)
                    listThreatsByVul = unionLists(listThreatsByVul, listVulnerabilities.Where(w => w.VulnerabilityNumber == vul.VulnerabilityNumber).First().Threats.ToList());
            else
                listThreatsByVul.AddRange(listThreats);

            // фильтрация УБИ по СФХ
            if (mf.clbThreatFilter.GetItemCheckState(3) == CheckState.Checked)
                foreach (SFH sfh in IS.listOfSFHs)
                    listThreatsBySFH = unionLists(listThreatsBySFH, listSFHs.Where(s => s.SFHNumber == sfh.SFHNumber).First().Threats.ToList());
            else
                listThreatsBySFH.AddRange(listThreats);

            listFilteredThreats = intersectLists(listThreatsBySource, intersectLists(listThreatsByWay, intersectLists(listThreatsByVul, listThreatsBySFH)));
            mf.dgvThreats.DataSource = listFilteredThreats;
            mf.lblThreatsCount.Text = "Кол-во УБИ: " + mf.dgvThreats.RowCount;
        }

        private List<Threat> unionLists(List<Threat> first, List<Threat> second)
        {
            List<Threat> result = new List<Threat>();
            result.AddRange(first);

            foreach(Threat item in second)
                if (!result.Exists(r => r.ThreatNumber == item.ThreatNumber))
                    result.Add(item);
            return result;
        }

        private List<Threat> intersectLists(List<Threat> first, List<Threat> second)
        {
            List<Threat> result = new List<Threat>();
            
            foreach (Threat item in first)
                if (second.Exists(r => r.ThreatNumber == item.ThreatNumber))
                    result.Add(item);
            return result.OrderBy(o => o.ThreatNumber).ToList();
        }

        private void dgvThreats_SelectionChanged(object sender, EventArgs e)
        {
            if (mf.dgvThreats.SelectedRows.Count > 0)
                mf.tbThreatDescription.Text =
                    listThreats.Where(t => t.ThreatNumber == (int)mf.dgvThreats.SelectedCells[mf.dgvThreats.Columns["ThreatNumber"].Index].Value).FirstOrDefault().Description;
            else
                mf.tbThreatDescription.Text = "Выберите угрозу для просмотра описания...";
        }

        private void tpActualThreats_Resize(object sender, EventArgs e)
        {
            mf.dgvThreats.Height = mf.tpThreatsNSD1.Height - 145;
            if (DDControl != null)
                mf.dgvActualThreatsNSD.Width = mf.tpThreatsNSD2.Width - DDControl.Width - 30;
            mf.dgvActualThreatsNSD.Height = mf.tpThreatsNSD2.Height - 50;
        }

        private void clbThreatFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            filterThreatList();
        }

        private void tcThreatsNSD_SelectedIndexChanged(object sender, EventArgs e)
        {
            tpActualThreats_Resize(null, null);
        }

        private void btnGotoDamage_Click(object sender, EventArgs e)
        {
            // выбрать все критерии фильтра УБИ
            for (int i = 0; i < mf.clbThreatFilter.Items.Count; i++)
                mf.clbThreatFilter.SetItemChecked(i, true);
            filterThreatList();
            if (mf.dgvThreats.Rows.Count == 0)
            {
                MessageBox.Show("Введите исходные данные для получения списка УБИ", "Список УБИ пуст", MessageBoxButtons.OK);
                checkInputData();
                return;
            }                

            if (IS.listOfInfoTypes.Count == 0)
            {
                if (MessageBox.Show("Для определения степеней ущерба, требуется выбрать виды информации. Перейти к выбору?", "Недостаточно исходных данных", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    mf.treeView.SelectedNode = mf.returnTreeNode("tnOptions");
                }
                return;
            }

            if (!mf.tcThreatsNSD.TabPages.Contains(mf.tpThreatsNSD2))
            {
                mf.tcThreatsNSD.TabPages.Add(mf.tpThreatsNSD2);
                initTabPageThreatsNSD2();
            }
            else
                if (MessageBox.Show("Запустить инициализацию формы определения степеней ущерба заново?", "Требуется подтверждение!", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    mf.tpThreatsNSD2.Controls.Remove(DDControl);
                    DDControl = null;
                    damageDegreeInput = null;
                    initTabPageThreatsNSD2();
                }
        }

        private void dgvActualThreats_SelectionChanged(object sender, EventArgs e)
        {
            if (mf.dgvActualThreatsNSD.SelectedRows.Count == 0)
                return;
            if (DDControl == null)
                return;
            int threatNnew = (int)mf.dgvActualThreatsNSD.SelectedRows[0].Cells[mf.dgvActualThreatsNSD.Columns["ThreatNumber"].Index].Value;
            DDControl.Update(threatNnew);
        }

        public int calcMaxDDInfo(int threatNumber, InfoType info)
        {
            int max = 0;
            for (int iCIA = 0; iCIA < 3; iCIA++)
                for (int iDT = 0; iDT < 7; iDT++)
                {
                    int current = damageDegreeInput[threatNumber][IS.listOfInfoTypes.IndexOf(info), iCIA, iDT];
                    if (current > max)
                        max = current;
                }
            return max;        
        }

        private void btnExportModelThreats_Click(object sender, EventArgs e)
        {
            mf.wsm.Visible = true;
            mf.wsm.Update();

            Microsoft.Office.Interop.Word._Application oWord = null;
            try
            {
                oWord = new Microsoft.Office.Interop.Word.Application();
            }
            catch
            {
                mf.wsm.Visible = false;
                MessageBox.Show("На ПК не установлен пакет Microsoft Office Word 2007 или позднее. Экспорт невозможен.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Microsoft.Office.Interop.Word._Document oDoc = null;
            try
            {
                oDoc = oWord.Documents.Add(Environment.CurrentDirectory + "\\template.docx");
            }
            catch
            {
                mf.wsm.Visible = false;
                MessageBox.Show("Отсутствует файл шаблона тех. проекта \"template.docx\". Экспорт невозможен.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Table wordTable;
            //oWord.Visible = true;

            #region Расчет уровня проектной защищенности
            wordTable = oDoc.Tables[1];
            int rowN = 4;
            Console.WriteLine(IS.listOfSFHs.Count + "");
            foreach (SFHType sfhtype in listSFHTypes)
            {
                foreach (SFH sfh in listSFHs.Where(sfh => sfh.SFHType == sfhtype).ToList())
                {
                    if (IS.listOfSFHs.Where(l => l.SFHNumber == sfh.SFHNumber).ToList().Count > 0)
                        wordTable.Cell(rowN, 4 - sfh.ProjectSecurity).Range.Text = "+";
                    rowN++;
                }
                rowN++;
            }
            rowN++;
            wordTable.Cell(rowN++, 2).Range.Text = mf.dgvProjectSecurityResult.Rows[0].Cells[1].Value.ToString();
            wordTable.Cell(rowN++, 3).Range.Text = mf.dgvProjectSecurityResult.Rows[1].Cells[1].Value.ToString();
            wordTable.Cell(rowN, 4).Range.Text = mf.dgvProjectSecurityResult.Rows[2].Cells[1].Value.ToString();
            #endregion

            #region Виды, типы и потенциал нарушителя
            Range bm = oDoc.Bookmarks["Intruder_Table"].Range;

            int rowsNumber = 1;
            foreach (System.Windows.Forms.CheckBox i in mf.clbIntruderTypes.Controls)
                if (i.CheckState == CheckState.Checked)
                    rowsNumber++;

            bm.Tables.Add(bm, rowsNumber, 3, Type.Missing, Type.Missing);
            wordTable = bm.Tables[1];

            wordTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            wordTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
            wordTable.Columns[1].PreferredWidth = 50f;
            wordTable.Columns[2].PreferredWidth = 25f;
            wordTable.Columns[3].PreferredWidth = 25f;
            wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
            wordTable.Rows[1].Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
            wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable.Range.Select();

            wordTable.Cell(1, 1).Range.Text = "Вид нарушителя";
            wordTable.Cell(1, 2).Range.Text = "Тип нарушителя";
            wordTable.Cell(1, 3).Range.Text = "Потенциал нарушителя";

            int row = 2;
            foreach (System.Windows.Forms.CheckBox i in mf.clbIntruderTypes.Controls)
            {
                if (i.CheckState == CheckState.Checked)
                {
                    wordTable.Cell(row, 1).Range.Text = i.Text;
                    wordTable.Cell(row, 2).Range.Text = ((StageIntruder)mf.stages["tnIntruder"]).getIntruderType(i.Text);
                    wordTable.Cell(row++, 3).Range.Text = ((StageIntruder)mf.stages["tnIntruder"]).getIntruderPotencial(i.Text);
                }
            }
            #endregion

            #region Фильтрация по нарушителю (1)
            bm = oDoc.Bookmarks["NSD_Filter_1_Intruder"].Range;

            mf.clbThreatFilter.SetItemChecked(0, true);
            mf.clbThreatFilter.SetItemChecked(1, false);
            mf.clbThreatFilter.SetItemChecked(2, false);
            mf.clbThreatFilter.SetItemChecked(3, false);
            filterThreatList();

            bm.Tables.Add(bm, listFilteredThreats.Count + 1, 3, Type.Missing, Type.Missing);
            wordTable = bm.Tables[1];

            wordTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            wordTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
            wordTable.Columns[1].PreferredWidth = 7.7f;
            wordTable.Columns[2].PreferredWidth = 55f;
            wordTable.Columns[3].PreferredWidth = 37.3f;
            wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
            wordTable.Rows[1].Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
            wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable.Range.Select();

            wordTable.Cell(1, 1).Range.Text = "№ п/п";
            wordTable.Cell(1, 2).Range.Text = "Идентификатор и название УБИ";
            wordTable.Cell(1, 3).Range.Text = "Источник УБИ (нарушитель)";

            row = 2;
            foreach (Threat t in listFilteredThreats)
            {
                wordTable.Cell(row, 1).Range.Text = (row - 1).ToString();
                wordTable.Cell(row, 2).Range.Text = "УБИ." + t.ThreatNumber + " " + t.Name.ToString();
                wordTable.Cell(row++, 3).Range.Text = t.stringSources;
            }
            #endregion

            #region Фильтрация по уязвимостям (2)
            bm = oDoc.Bookmarks["NSD_Filter_2_Vuls"].Range;

            mf.clbThreatFilter.SetItemChecked(0, true);
            mf.clbThreatFilter.SetItemChecked(1, false);
            mf.clbThreatFilter.SetItemChecked(2, true);
            mf.clbThreatFilter.SetItemChecked(3, false);
            filterThreatList();

            bm.Tables.Add(bm, listFilteredThreats.Count + 1, 3, Type.Missing, Type.Missing);
            wordTable = bm.Tables[1];

            wordTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            wordTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
            wordTable.Columns[1].PreferredWidth = 8f;
            wordTable.Columns[2].PreferredWidth = 42f;
            wordTable.Columns[3].PreferredWidth = 50f;
            wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
            wordTable.Rows[1].Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
            wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable.Range.Select();

            wordTable.Cell(1, 1).Range.Text = "№ п/п";
            wordTable.Cell(1, 2).Range.Text = "Идентификатор и название УБИ";
            wordTable.Cell(1, 3).Range.Text = "Уязвимости, способствующие реализации УБИ";

            row = 2;
            foreach (Threat t in listFilteredThreats)
            {

                wordTable.Cell(row, 1).Range.Text = (row - 1).ToString();
                wordTable.Cell(row, 2).Range.Text = "УБИ." + t.ThreatNumber + " " + t.Name.ToString();
                wordTable.Cell(row++, 3).Range.Text = t.stringVuls;
            }
            #endregion

            #region Фильтрация по способам реализации (3)
            bm = oDoc.Bookmarks["NSD_Filter_3_Implement"].Range;

            mf.clbThreatFilter.SetItemChecked(0, true);
            mf.clbThreatFilter.SetItemChecked(1, true);
            mf.clbThreatFilter.SetItemChecked(2, true);
            mf.clbThreatFilter.SetItemChecked(3, false);
            filterThreatList();

            bm.Tables.Add(bm, listFilteredThreats.Count + 1, 3, Type.Missing, Type.Missing);
            wordTable = bm.Tables[1];

            wordTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            wordTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
            wordTable.Columns[1].PreferredWidth = 8f;
            wordTable.Columns[2].PreferredWidth = 27f;
            wordTable.Columns[3].PreferredWidth = 65f;
            wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
            wordTable.Rows[1].Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
            wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable.Range.Select();

            wordTable.Cell(1, 1).Range.Text = "№ п/п";
            wordTable.Cell(1, 2).Range.Text = "Идентификатор и название УБИ";
            wordTable.Cell(1, 3).Range.Text = "Возможные способы реализации УБИ";

            row = 2;
            foreach (Threat t in listFilteredThreats)
            {
                wordTable.Cell(row, 1).Range.Text = (row - 1).ToString();
                wordTable.Cell(row, 2).Range.Text = "УБИ." + t.ThreatNumber + " " + t.Name.ToString();
                wordTable.Cell(row++, 3).Range.Text = t.stringWays;
            }
            #endregion

            #region Фильтрация по СФХ (4)
            bm = oDoc.Bookmarks["NSD_Filter_4_SFH"].Range;

            mf.clbThreatFilter.SetItemChecked(0, true);
            mf.clbThreatFilter.SetItemChecked(1, true);
            mf.clbThreatFilter.SetItemChecked(2, true);
            mf.clbThreatFilter.SetItemChecked(3, true);
            filterThreatList();

            bm.Tables.Add(bm, listFilteredThreats.Count + 1, 3, Type.Missing, Type.Missing);
            wordTable = bm.Tables[1];

            wordTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            wordTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
            wordTable.Columns[1].PreferredWidth = 8f;
            wordTable.Columns[2].PreferredWidth = 27f;
            wordTable.Columns[3].PreferredWidth = 65f;
            wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
            wordTable.Rows[1].Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
            wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable.Range.Select();

            wordTable.Cell(1, 1).Range.Text = "№ п/п";
            wordTable.Cell(1, 2).Range.Text = "Идентификатор и название УБИ";
            wordTable.Cell(1, 3).Range.Text = "Описание УБИ";

            row = 2;
            foreach (Threat t in listFilteredThreats)
            {
                wordTable.Cell(row, 1).Range.Text = (row - 1).ToString();
                wordTable.Cell(row, 2).Range.Text = "УБИ." + t.ThreatNumber + " " + t.Name.ToString();
                wordTable.Cell(row++, 3).Range.Text = t.Description;
            }
            #endregion

            #region Уровень проектной защищенности (слово)
            bm = oDoc.Bookmarks["Project_Security"].Range;
            string PSLtext = "";
            switch (IS.ProjectSecutiryLvl)
            {
                case 0:
                    PSLtext = "низкий";
                    break;
                case 1:
                    PSLtext = "средний";
                    break;
                case 2:
                    PSLtext = "высокий";
                    break;
            }
            bm.Text = PSLtext;
            #endregion

            #region Возможность реализации УБИ
            bm = oDoc.Bookmarks["Feasibility"].Range;

            bm.Tables.Add(bm, mf.dgvActualThreatsNSD.Rows.Count + 1, 5, Type.Missing, Type.Missing);
            wordTable = bm.Tables[1];

            wordTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            wordTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
            wordTable.Columns[1].PreferredWidth = 8f;
            wordTable.Columns[2].PreferredWidth = 29.8f;
            wordTable.Columns[3].PreferredWidth = 24.6f;
            wordTable.Columns[4].PreferredWidth = 20.2f;
            wordTable.Columns[5].PreferredWidth = 17.4f;
            wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
            wordTable.Rows[1].Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
            wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable.Range.Select();

            wordTable.Cell(1, 1).Range.Text = "№ п/п";
            wordTable.Cell(1, 2).Range.Text = "Идентификатор и название УБИ";
            wordTable.Cell(1, 3).Range.Text = "Источник угрозы (нарушитель)";
            wordTable.Cell(1, 4).Range.Text = "Уровень проектной защищенности ИС";
            wordTable.Cell(1, 5).Range.Text = "Возможность реализации угрозы";

            wordTable.Cell(2, 4).Range.Text = PSLtext;

            row = 2;
            foreach (Threat t in listFilteredThreats)
            {
                wordTable.Cell(row, 1).Range.Text = (row - 1).ToString();
                wordTable.Cell(row, 2).Range.Text = "УБИ." + t.ThreatNumber + " " + t.Name.ToString();
                wordTable.Cell(row++, 3).Range.Text = t.stringSources;
            }
            row = 2;
            foreach (DataGridViewRow dgvrow in mf.dgvActualThreatsNSD.Rows)
                wordTable.Cell(row++, 5).Range.Text = dgvrow.Cells[2].Value.ToString();

            wordTable.Rows[2].Cells[4].Merge(wordTable.Rows[mf.dgvActualThreatsNSD.Rows.Count + 1].Cells[4]);
            #endregion

            #region Степени ущерба
            bm = oDoc.Bookmarks["Damage_Degree"].Range;

            //bm.Tables.Add(bm, mf.dgvActualThreatsNSD.Rows.Count + 2, 3 + IS.listOfInfoTypes.Count, Type.Missing, Type.Missing);
            wordTable = bm.Tables[1];

            //wordTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            //wordTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            //wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
            //wordTable.Columns[1].PreferredWidth = 7f;
            //wordTable.Columns[2].PreferredWidth = 33.1f;

            //wordTable.Columns[3].PreferredWidth = 45.1f;

            //wordTable.Columns[5].PreferredWidth = 13.8f;
            //wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
            //wordTable.Rows[1].Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
            //wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable.Range.Select();

            //wordTable.Cell(1, 1).Range.Text = "№ п/п";
            //wordTable.Cell(1, 2).Range.Text = "Идентификатор и название УБИ";
            //wordTable.Cell(1, 3).Range.Text = "Степень ущерба в результате наруше-ния каждого свойства безопасности относительно всех видов ущерба";
            //wordTable.Cell(1, 3 + IS.listOfInfoTypes.Count).Range.Text = "Итоговая степень ущерба";
            //wordTable.Rows[1].Cells[1].Merge(wordTable.Rows[2].Cells[1]);
            //wordTable.Rows[1].Cells[2].Merge(wordTable.Rows[2].Cells[2]);
            //wordTable.Rows[1].Cells[3 + IS.listOfInfoTypes.Count].Merge(wordTable.Rows[2].Cells[3 + IS.listOfInfoTypes.Count]);
            //if (IS.listOfInfoTypes.Count >= 2)
            //    wordTable.Rows[1].Cells[3].Merge(wordTable.Rows[1].Cells[2 + IS.listOfInfoTypes.Count]);
            int infoN = 0;
            wordTable.Cell(2, 3).Split(1, IS.listOfInfoTypes.Count);
            foreach (InfoType info in IS.listOfInfoTypes)
            {
                wordTable.Cell(2, 3 + infoN).Range.Text = info.TypeName;
                infoN++;
            }

            row = 3;

            foreach (DataGridViewRow dgvrow in mf.dgvActualThreatsNSD.Rows)
            {
                wordTable.Rows.Add();
                wordTable.Cell(row, 1).Range.Text = (row - 1).ToString();
                wordTable.Cell(row, 2).Range.Text = "УБИ." + dgvrow.Cells[0].Value.ToString() + " " + dgvrow.Cells[1].Value.ToString();
                infoN = 0;
                foreach (InfoType info in IS.listOfInfoTypes)
                {
                    switch (calcMaxDDInfo(Convert.ToInt32(dgvrow.Cells[0].Value.ToString()), info))
                    {
                        case 3:
                            wordTable.Cell(row, 3 + infoN).Range.Text = "Высокая";
                            break;
                        case 2:
                            wordTable.Cell(row, 3 + infoN).Range.Text = "Средняя";
                            break;
                        case 1:
                            wordTable.Cell(row, 3 + infoN).Range.Text = "Низкая";
                            break;
                        case 0:
                            wordTable.Cell(row, 3 + infoN).Range.Text = "Не определена";
                            break;
                    }
                    infoN++;
                }
                wordTable.Cell(row++, 3 + IS.listOfInfoTypes.Count).Range.Text = dgvrow.Cells[3].Value.ToString();
            }
            //damageDegreeInput[currentThreatNumber][itype, iCIA, iDT] = 3;
            #endregion

            #region Актуальные УБИ НСД
            bm = oDoc.Bookmarks["ACT_NSD"].Range;

            bm.Tables.Add(bm, mf.dgvActualThreatsNSD.Rows.Count + 1, 5, Type.Missing, Type.Missing);
            wordTable = bm.Tables[1];

            wordTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            wordTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            wordTable.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
            wordTable.Columns[1].PreferredWidth = 7f;
            wordTable.Columns[2].PreferredWidth = 35.3f;
            wordTable.Columns[3].PreferredWidth = 18.2f;
            wordTable.Columns[4].PreferredWidth = 21.2f;
            wordTable.Columns[5].PreferredWidth = 18.1f;
            wordTable.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft;
            wordTable.Rows[1].Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;
            wordTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wordTable.Range.Select();

            wordTable.Cell(1, 1).Range.Text = "№ п/п";
            wordTable.Cell(1, 2).Range.Text = "Идентификатор и название УБИ";
            wordTable.Cell(1, 3).Range.Text = "Возможность реализации УБИ";
            wordTable.Cell(1, 4).Range.Text = "Степень возможного ущерба в результате реализации УБИ";
            wordTable.Cell(1, 5).Range.Text = "Актуальность";

            row = 2;
            foreach (DataGridViewRow dgvrow in mf.dgvActualThreatsNSD.Rows)
            {
                wordTable.Cell(row, 1).Range.Text = (row - 1).ToString();
                wordTable.Cell(row, 2).Range.Text = "УБИ." + dgvrow.Cells[0].Value.ToString() + " " + dgvrow.Cells[1].Value.ToString();
                wordTable.Cell(row, 3).Range.Text = dgvrow.Cells[2].Value.ToString();
                wordTable.Cell(row, 4).Range.Text = dgvrow.Cells[3].Value.ToString();
                wordTable.Cell(row++, 5).Range.Text = dgvrow.Cells[4].Value.ToString();
            }
            #endregion

            //bm.Text = "hello";
            oDoc.Activate();
            mf.FindAndReplace(oWord, "{Название ИС}", "ГИС \"" + IS.ISName + "\"");

            //oDoc.SaveAs(FileName: Environment.CurrentDirectory + "\\2.docx");
            //oDoc.Close();
            mf.wsm.Visible = false;
            oWord.Visible = true;
        }
    }
}
