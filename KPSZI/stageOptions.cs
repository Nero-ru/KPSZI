using KPSZI.Model;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPSZI
{
    class StageOptions : Stage
    {
        List<InfoType> listInfoTypes;
        List<SFH> listSFHs;
        List<SFHType> listSFHTypes;

        List<CheckBox> checkboxesSFH;
        List<RadioButton> radiobuttonsSFH;

        double[] projectSecuritySumm;
        double[] projectSecurityChecked;

        protected override ImageList imageListForTabPage { get; set; }

        public StageOptions (TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS)
            :base(stageTab, stageNode, mainForm, IS)
        {
            
        }

        protected override void initTabPage()
        {
            using (KPSZIContext db = new KPSZIContext())
            {
                // Инициализация списка видов угроз

                listInfoTypes = db.InfoTypes.ToList();

                ((ListBox)mf.lbInfoTypes).DataSource = listInfoTypes;
                ((ListBox)mf.lbInfoTypes).DisplayMember = "TypeName";
                ((ListBox)mf.lbInfoTypes).ValueMember = "InfoTypeId";

                mf.lbInfoTypes.Size = new Size(400, 16 * mf.lbInfoTypes.Items.Count);

                //Инициализация СФХ групбоксов
                listSFHTypes = db.SFHTypes.ToList();
                listSFHs = db.SFHs.ToList();

                GroupBox gb;
                RadioButton rb;
                CheckBox cb;
                checkboxesSFH = new List<CheckBox>();
                radiobuttonsSFH = new List<RadioButton>();
                int i = 0;
                int j = 0;
                int gbY = 85;

                foreach (SFHType itemSFHType in listSFHTypes)
                {
                    gb = new GroupBox();
                    gb.Location = new Point(7, gbY);
                    gb.Text = itemSFHType.Name;
                    foreach (SFH itemSFH in itemSFHType.SFHs)
                    {
                        if (itemSFHType.MultipleChoice)
                        {
                            cb = new CheckBox();
                            cb.Text = itemSFH.Name;
                            cb.Margin = new Padding(10, 5, 5, 5);
                            cb.Location = new Point(10, 17 + (17 * j));
                            cb.Size = new Size(440, 17);
                            cb.CheckedChanged += new System.EventHandler(rbSFH_CheckedChanged);
                            checkboxesSFH.Add(cb);
                            gb.Controls.Add(cb);
                            j++;
                        }
                        else
                        {
                            rb = new RadioButton();
                            rb.Text = itemSFH.Name;
                            rb.Margin = new Padding(10, 5, 5, 5);
                            rb.Location = new Point(10, 17 + (17 * j));
                            rb.Size = new Size(440, 17);
                            rb.CheckedChanged += new System.EventHandler(rbSFH_CheckedChanged);
                            radiobuttonsSFH.Add(rb);
                            gb.Controls.Add(rb);
                            j++;
                        }
                    }

                    gb.Size = new Size(440, 25 + j * 17);
                    gbY += 30 + j * 17;
                    j = 0;
                    stageTab.Controls.Add(gb);
                    i++;
                }
            }

            mf.lbInfoTypes.SelectedIndexChanged += new System.EventHandler(lbInfoTypes_SelectedIndexChanged);
            mf.dgvProjectSecurityResult.SelectionChanged += new System.EventHandler(dgvProjectSecurityResult_SelectionChanged);

            mf.dgvProjectSecurityResult.Rows.Add(4);
            mf.dgvProjectSecurityResult.Rows[0].Cells[0].Value = "Процент характеристик, соответствующих уровню «высокий»";
            mf.dgvProjectSecurityResult.Rows[1].Cells[0].Value = "Процент характеристик, соответствующих уровню «средний»";
            mf.dgvProjectSecurityResult.Rows[2].Cells[0].Value = "Процент характеристик, соответствующих уровню «низкий»";
            mf.dgvProjectSecurityResult.Rows[3].Cells[0].Value = "Уровень проектной защищенности";
            mf.dgvProjectSecurityResult.Rows[3].DefaultCellStyle.BackColor = Color.FromArgb(234, 240, 255);
            mf.dgvProjectSecurityResult.Rows[3].DefaultCellStyle.ForeColor = Color.FromArgb(28, 32, 57);
            mf.dgvProjectSecurityResult.Columns[0].Width = mf.dgvProjectSecurityResult.Width - 70;
        }
        
        public override void saveChanges()
        {
            projectSecuritySumm = new double[3] { 0, 0, 0 };
            projectSecurityChecked = new double[3] { 0, 0, 0 };

            IS.ISName = mf.tbISName.Text;

            foreach (RadioButton rb in radiobuttonsSFH)
            {
                // Добавление СФХ в IS
                SFH sfh = returnSFH(rb.Text);
                if (rb.Checked && !IS.listOfSFHs.Contains(sfh))
                    IS.listOfSFHs.Add(sfh);
                if (!rb.Checked && IS.listOfSFHs.Contains(sfh))
                    IS.listOfSFHs.Remove(sfh);

                // Подсчет характеристик (проектная защищенность)
                if (rb.Checked) projectSecurityChecked[sfh.ProjectSecurity]++;
                projectSecuritySumm[sfh.ProjectSecurity]++;
            }

            foreach (CheckBox cb in checkboxesSFH)
            {
                SFH sfh = returnSFH(cb.Text);
                if (cb.Checked && !IS.listOfSFHs.Contains(sfh))
                    IS.listOfSFHs.Add(sfh);
                if (!cb.Checked && IS.listOfSFHs.Contains(sfh))
                    IS.listOfSFHs.Remove(sfh);

                if (cb.Checked) projectSecurityChecked[sfh.ProjectSecurity]++;
                projectSecuritySumm[sfh.ProjectSecurity]++;
            }

            // вывод процент характеристик, соответствующих уровню высокий, средний, низкий
            mf.dgvProjectSecurityResult.Rows[0].Cells[1].Value = string.Format("{0:0.##}", projectSecurityChecked[2] / projectSecuritySumm[2] * 100) + "%";
            mf.dgvProjectSecurityResult.Rows[1].Cells[1].Value = string.Format("{0:0.##}", projectSecurityChecked[1] / projectSecuritySumm[1] * 100) + "%";
            mf.dgvProjectSecurityResult.Rows[2].Cells[1].Value = string.Format("{0:0.##}", projectSecurityChecked[0] / projectSecuritySumm[0] * 100) + "%";
            
            // Вычисление уровня проектной защищенности
            if (projectSecurityChecked[2] / projectSecuritySumm[2] >= 0.8 && projectSecurityChecked[0] == 0)
            {
                mf.dgvProjectSecurityResult.Rows[3].Cells[1].Value = "Высокий";
                IS.ProjectSecutiryLvl = 2;
            }
            else
                if (projectSecurityChecked[2] / projectSecuritySumm[2] + projectSecurityChecked[1] / projectSecuritySumm[1] >= 0.9)
                {
                    mf.dgvProjectSecurityResult.Rows[3].Cells[1].Value = "Средний";
                    IS.ProjectSecutiryLvl = 1;
                }
                else
                    {
                        mf.dgvProjectSecurityResult.Rows[3].Cells[1].Value = "Низкий";
                        IS.ProjectSecutiryLvl = 0;
                    }
        }

        public override void enterTabPage()
        {
            // берем из экземпляра выбранные виды информации и возвращаем их в чеклистбокс при переходе по вкладкам
            for (int i = 0; i < IS.listOfInfoTypes.Count; i++)
            {
                int k = mf.lbInfoTypes.Items.IndexOf(IS.listOfInfoTypes[i]);
                mf.lbInfoTypes.SetItemChecked(k, true);
            }
        }

        public SFH returnSFH(string SFHname)
        {
            foreach(SFH itemSFH in listSFHs)
            {
                if (itemSFH.Name == SFHname) return itemSFH;
            }
            return null;
        }
        
        private void rbSFH_CheckedChanged(object sender, EventArgs e)
        {
            saveChanges();
        }
        
        private void lbInfoTypes_SelectedIndexChanged(object sender, EventArgs e)
        {
            // При нажатии на галочку все выбранные 
            // виды информации помещаются в экземпляр ИС

            object[] buf = new object[mf.lbInfoTypes.CheckedItems.Count];
            IS.listOfInfoTypes.Clear();
            mf.lbInfoTypes.CheckedItems.CopyTo(buf, 0);
            for (int i = 0; i < buf.Length; i++)
                IS.listOfInfoTypes.Add((InfoType)buf.GetValue(i));
        }

        private void dgvProjectSecurityResult_SelectionChanged(object sender, EventArgs e)
        {
            if (mf.dgvProjectSecurityResult.SelectedCells.Count > 0)
                mf.dgvProjectSecurityResult.SelectedCells[0].Selected = false;
        }

    }
}
