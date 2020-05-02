using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KPSZI.Model;

namespace KPSZI
{
    public partial class DamageDegreeControl : UserControl
    {
        //int[,,] damageForInfoType;
        public List<string> listOfInfoTypes;
        public DataGridView[] dgvDamage;
        Label[] lblInfoTypeName;
        public int currentThreatNumber;
        internal List<Threat> listFilteredThreats;
        Dictionary<int, int[,,]> damageDegreeInput;
        MainForm mf;


        public DamageDegreeControl(List<InfoType> listOfInfoTypes, List<Threat> listFilteredThreats, Dictionary<int, int[,,]> damageDegreeInput, MainForm mf)
        {
            InitializeComponent();
            this.listFilteredThreats = listFilteredThreats;
            this.listOfInfoTypes = new List<string>();
            for (int i = 0; i < listOfInfoTypes.Count; i++)
                this.listOfInfoTypes.Add(listOfInfoTypes[i].TypeName);
            dgvDamage = new DataGridView[listOfInfoTypes.Count()];
            lblInfoTypeName = new Label[listOfInfoTypes.Count()];
            this.damageDegreeInput = damageDegreeInput;
            this.mf = mf;
            currentThreatNumber = Convert.ToInt32(mf.dgvActualThreatsNSD.Rows[0].Cells[0].Value.ToString());
            initDGVs();
        }

        // Очистка контрола от элементов управления
        public void Clear()
        {
            for (int i = 0; i < dgvDamage.Count(); i++)
            {
                Controls.Remove(dgvDamage[i]);
                Controls.Remove(lblInfoTypeName[i]);
            }
        }

        // Сохранение значений ущерба, введенных в контроле
        public void SaveControl()
        {
            if (!(currentThreatNumber > 0))
                return;
            // сохраняем данные по Уби
            int countSelected = 0;
            for (int itype = 0; itype < listOfInfoTypes.Count; itype++)
            {
                for (int iCIA = 0; iCIA < 3; iCIA++)
                {
                    switch (iCIA)
                    {
                        case 0:
                            if (listFilteredThreats.Where(t => t.ThreatNumber == currentThreatNumber).FirstOrDefault().ConfidenceViolation == false)
                                continue;
                            break;
                        case 1:
                            if (listFilteredThreats.Where(t => t.ThreatNumber == currentThreatNumber).FirstOrDefault().IntegrityViolation == false)
                                continue;
                            break;
                        case 2:
                            if (listFilteredThreats.Where(t => t.ThreatNumber == currentThreatNumber).FirstOrDefault().AvailabilityViolation == false)
                                continue;
                            break;
                    }

                    for (int iDT = 0; iDT < 7; iDT++)
                    {
                        if (dgvDamage[itype].Rows[iDT].Cells[iCIA + 1].Value == null)
                            return;
                        string cellValue = dgvDamage[itype].Rows[iDT].Cells[iCIA + 1].Value.ToString();

                        switch (cellValue)
                        {
                            case "Высокая":
                                damageDegreeInput[currentThreatNumber][itype, iCIA, iDT] = 3;
                                countSelected++;
                                break;
                            case "Средняя":
                                damageDegreeInput[currentThreatNumber][itype, iCIA, iDT] = 2;
                                countSelected++;
                                break;
                            case "Низкая":
                                damageDegreeInput[currentThreatNumber][itype, iCIA, iDT] = 1;
                                countSelected++;
                                break;
                            case "–":
                                damageDegreeInput[currentThreatNumber][itype, iCIA, iDT] = 0;
                                break;
                        }
                    }
                }
            }
        }

        // Выгрузка значений степени ущерба в контрол для выбранной угрозы
        public void LoadControl()
        {
            for (int itype = 0; itype < listOfInfoTypes.Count; itype++)
            {
                for (int iCIA = 0; iCIA < 3; iCIA++)
                {
                    switch (iCIA)
                    {
                        case 0:
                            if (listFilteredThreats.Where(t => t.ThreatNumber == currentThreatNumber).FirstOrDefault().ConfidenceViolation == false)
                                continue;
                            break;
                        case 1:
                            if (listFilteredThreats.Where(t => t.ThreatNumber == currentThreatNumber).FirstOrDefault().IntegrityViolation == false)
                                continue;
                            break;
                        case 2:
                            if (listFilteredThreats.Where(t => t.ThreatNumber == currentThreatNumber).FirstOrDefault().AvailabilityViolation == false)
                                continue;
                            break;
                    }

                    for (int iDT = 0; iDT < 7; iDT++)
                    {
                        int damageValue = damageDegreeInput[currentThreatNumber][itype, iCIA, iDT];

                        switch (damageValue)
                        {
                            case 3:
                                dgvDamage[itype].Rows[iDT].Cells[iCIA + 1].Value = "Высокая";
                                break;
                            case 2:
                                dgvDamage[itype].Rows[iDT].Cells[iCIA + 1].Value = "Средняя";
                                break;
                            case 1:
                                dgvDamage[itype].Rows[iDT].Cells[iCIA + 1].Value = "Низкая";
                                break;
                            case 0:
                                dgvDamage[itype].Rows[iDT].Cells[iCIA + 1].Value = "–";
                                break;
                        }
                    }
                }
                dgvDamage[itype].ClearSelection();
            }
        }

        // Обновление котрола для угрозы threatNew
        public void Update(int threatNew)
        {
            currentThreatNumber = threatNew;
            Clear();
            initDGVs();
            LoadControl();
        }
        
        // Инициализация Контрола
        public void initDGVs()
        {
            int currentH = 0;
            for (int i = 0; i < dgvDamage.Count(); i++)
            {
                // 
                // lblDamageType
                //
                currentH += 9;
                lblInfoTypeName[i] = new Label();
                lblInfoTypeName[i].AutoSize = true;
                lblInfoTypeName[i].Location = new Point(120, currentH);
                lblInfoTypeName[i].Size = new Size(67, 13);
                lblInfoTypeName[i].TabIndex = 4;
                lblInfoTypeName[i].Text = listOfInfoTypes[i];

                bool stubC = !listFilteredThreats.Where(t => t.ThreatNumber == currentThreatNumber).FirstOrDefault().ConfidenceViolation;
                bool stubI = !listFilteredThreats.Where(t => t.ThreatNumber == currentThreatNumber).FirstOrDefault().IntegrityViolation;
                bool stubA = !listFilteredThreats.Where(t => t.ThreatNumber == currentThreatNumber).FirstOrDefault().AvailabilityViolation;
                // 
                // dgvDamage
                // 
                currentH += 26;
                dgvDamage[i] = new DataGridView();
                dgvDamage[i].AllowUserToAddRows = false;
                dgvDamage[i].AllowUserToDeleteRows = false;
                dgvDamage[i].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgvDamage[i].AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                dgvDamage[i].BackgroundColor = SystemColors.Window;
                dgvDamage[i].BorderStyle = BorderStyle.None;
                dgvDamage[i].ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;

                DataGridViewTextBoxColumn damageType = new DataGridViewTextBoxColumn();
                DataGridViewTextBoxColumn columnCstub = new DataGridViewTextBoxColumn();
                columnCstub.DefaultCellStyle.BackColor = dgvDamage[i].GridColor;
                columnCstub.ReadOnly = true;
                columnCstub.HeaderText = "К";
                DataGridViewTextBoxColumn columnIstub = new DataGridViewTextBoxColumn();
                columnIstub.DefaultCellStyle.BackColor = dgvDamage[i].GridColor;
                columnIstub.ReadOnly = true;
                columnIstub.HeaderText = "Ц";
                DataGridViewTextBoxColumn columnAstub = new DataGridViewTextBoxColumn();
                columnAstub.DefaultCellStyle.BackColor = dgvDamage[i].GridColor;
                columnAstub.ReadOnly = true;
                columnAstub.HeaderText = "Д";
                DataGridViewComboBoxColumn columnC = new DataGridViewComboBoxColumn();
                DataGridViewComboBoxColumn columnI = new DataGridViewComboBoxColumn();
                DataGridViewComboBoxColumn columnA = new DataGridViewComboBoxColumn();

                dgvDamage[i].Columns.Add(damageType);
                if (stubC)
                    dgvDamage[i].Columns.Add(columnCstub);
                else
                    dgvDamage[i].Columns.Add(columnC);
                if (stubI)
                    dgvDamage[i].Columns.Add(columnIstub);
                else
                    dgvDamage[i].Columns.Add(columnI);
                if (stubA)
                    dgvDamage[i].Columns.Add(columnAstub);
                else
                    dgvDamage[i].Columns.Add(columnA);

                dgvDamage[i].RowsAdded += new DataGridViewRowsAddedEventHandler(dgvDamage_RowsAdded);

                dgvDamage[i].Rows.Add("Экономический");
                dgvDamage[i].Rows.Add("Социальный");
                dgvDamage[i].Rows.Add("Политический");
                dgvDamage[i].Rows.Add("Репутационный");
                dgvDamage[i].Rows.Add("В области обороны.."); // , безопасности и правопорядка
                dgvDamage[i].Rows.Add("Субъекту ПДн");
                dgvDamage[i].Rows.Add("Технологический");

                dgvDamage[i].EditMode = DataGridViewEditMode.EditOnEnter;
                dgvDamage[i].Margin = new Padding(0);
                dgvDamage[i].Size = new Size(Width, dgvDamage[i].Size.Height);
                dgvDamage[i].Location = new Point(0, currentH);
                dgvDamage[i].RowHeadersVisible = false;
                dgvDamage[i].SelectionMode = DataGridViewSelectionMode.CellSelect;
                
                currentH += dgvDamage[i].Size.Height;
                // 
                // damageType
                // 
                damageType.HeaderText = "Вид ущерба";
                damageType.ReadOnly = true;
                damageType.SortMode = DataGridViewColumnSortMode.NotSortable;
                // 
                // columnC
                // 
                columnC.HeaderText = "К";
                columnC.Items.AddRange(new object[] { "Высокая", "Средняя", "Низкая", "–" });
                columnC.Name = "columnC";
                columnC.Resizable = DataGridViewTriState.True;
                columnC.SortMode = DataGridViewColumnSortMode.Automatic;
                columnC.ToolTipText = "Конфиденциальность";
                columnC.SortMode = DataGridViewColumnSortMode.NotSortable;
                // 
                // columnI
                // 
                columnI.HeaderText = "Ц";
                columnI.Items.AddRange(new object[] { "Высокая", "Средняя", "Низкая", "–"});
                columnI.Name = "columnI";
                columnI.Resizable = DataGridViewTriState.True;
                columnI.SortMode = DataGridViewColumnSortMode.Automatic;
                columnI.ToolTipText = "Целостность";
                columnI.SortMode = DataGridViewColumnSortMode.NotSortable;
                // 
                // columnA
                // 
                columnA.HeaderText = "Д";
                columnA.Items.AddRange(new object[] { "Высокая", "Средняя", "Низкая", "–"});
                columnA.Name = "columnA";
                columnA.Resizable = DataGridViewTriState.True;
                columnA.SortMode = DataGridViewColumnSortMode.Automatic;
                columnA.ToolTipText = "Доступность";
                columnA.SortMode = DataGridViewColumnSortMode.NotSortable;

                Controls.Add(dgvDamage[i]);
                Controls.Add(lblInfoTypeName[i]);
                Height = currentH;

                dgvDamage[i].CellValueChanged += new DataGridViewCellEventHandler(calcDD);
                dgvDamage[i].CurrentCellDirtyStateChanged += dgvDamage_CurrentCellDirtyStateChanged;
                dgvDamage[i].SelectionChanged += dgvDamage_SelectionChanged;
            }
        }
        
        // Вычисление Итоговой степени возможного ущерба для выбранной УБИ
        public void calcDD(object sender, DataGridViewCellEventArgs e)
        {
            if (!(currentThreatNumber > 0))
                return;
            if (mf.dgvActualThreatsNSD.SelectedRows.Count == 0)
                return;

            SaveControl();
            int maxDamage = 0;            
            for (int itype = 0; itype < listOfInfoTypes.Count; itype++)
            {
                for (int iCIA = 0; iCIA < 3; iCIA++)
                {
                    switch (iCIA)
                    {
                        case 0:
                            if (listFilteredThreats.Where(t => t.ThreatNumber == currentThreatNumber).FirstOrDefault().ConfidenceViolation == false)
                                continue;
                            break;
                        case 1:
                            if (listFilteredThreats.Where(t => t.ThreatNumber == currentThreatNumber).FirstOrDefault().IntegrityViolation == false)
                                continue;
                            break;
                        case 2:
                            if (listFilteredThreats.Where(t => t.ThreatNumber == currentThreatNumber).FirstOrDefault().AvailabilityViolation == false)
                                continue;
                            break;
                    }

                    for (int iDT = 0; iDT < 7; iDT++)
                    {
                        if (maxDamage < damageDegreeInput[currentThreatNumber][itype, iCIA, iDT])
                            maxDamage = damageDegreeInput[currentThreatNumber][itype, iCIA, iDT];
                    }
                }
            }
            string DDValue = "";
            if (maxDamage == 3) DDValue = "Высокая";
            if (maxDamage == 2) DDValue = "Средняя";
            if (maxDamage == 1) DDValue = "Низкая";
            
            mf.dgvActualThreatsNSD.SelectedRows[0].Cells[mf.dgvActualThreatsNSD.Columns["DamageDegree"].Index].Value = DDValue;

            if (maxDamage == 0 && mf.dgvActualThreatsNSD.SelectedRows[0].
                Cells[mf.dgvActualThreatsNSD.Columns["Feasibility"].Index].Value.ToString() == "Высокая")
                mf.dgvActualThreatsNSD.SelectedRows[0].Cells[mf.dgvActualThreatsNSD.Columns["DamageDegree"].Index].Value = "Не определена";

            calcActuality();
        }

        public void calcActuality()
        {
            string cellDDValue = mf.dgvActualThreatsNSD.SelectedRows[0].Cells[mf.dgvActualThreatsNSD.Columns["DamageDegree"].Index].Value.ToString();

            if (cellDDValue == "")
                return;

            string cellFeasibilityValue = mf.dgvActualThreatsNSD.SelectedRows[0].Cells[mf.dgvActualThreatsNSD.Columns["Feasibility"].Index].Value.ToString();
            
            DataGridViewCell cellActualityValue = mf.dgvActualThreatsNSD.SelectedRows[0].Cells[mf.dgvActualThreatsNSD.Columns["IsActual"].Index];

            if ((cellFeasibilityValue == "Средняя" && cellDDValue == "Низкая") 
                || (cellFeasibilityValue == "Низкая" && (cellDDValue == "Низкая" || cellDDValue == "Средняя")))
            {
                mf.dgvActualThreatsNSD.SelectedRows[0].DefaultCellStyle.BackColor = Color.White;
                cellActualityValue.Value = "Неактуальная";
            }
            else
            {
                mf.dgvActualThreatsNSD.SelectedRows[0].DefaultCellStyle.BackColor = Color.LightGreen;
                cellActualityValue.Value = "Актуальная";
            }
        }

        // Адаптация высоты DataGridView по мере добавления строк
        private void dgvDamage_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            // автоматическое вычисление высоты DGV
            ((Control)sender).Height = ((DataGridView)sender).Rows.GetRowsHeight(DataGridViewElementStates.Visible) + 
                ((DataGridView)sender).ColumnHeadersHeight;            
        }

        // Обновление значение ComboBox сразу при изменении значения, а не после потери фокуса
        private void dgvDamage_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            ((DataGridView)sender).CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void dgvDamage_SelectionChanged(object sender, EventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            if (dgv.SelectedCells.Count > 0)
                dgv.SelectedCells[0].Selected = false;
        }
    }
}
