using KPSZI.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPSZI
{
    public partial class FillThreatsForm : Form
    {
        MainForm mf;
        List<SFHType> listSFHTypes;
        List<SFH> listSFHs;
        List<CheckBox> checkboxesSFH;
        List<RadioButton> radiobuttonsSFH;
        List<Vulnerability> listVuls;
        List<ImplementWay> ListImplementWays;
        List<Threat> listEmptyThreats;

        public FillThreatsForm(MainForm mf)
        {
            InitializeComponent();
            this.mf = mf;
            initForm();
        }

        private void initForm()
        {
            listEmptyThreats = new List<Threat>();
            foreach (Threat t in ((StageActualThreats)mf.stages["tnActualThreats"]).listThreats)
            {
                if (t.ImplementWays.ToList().Count == 0 || t.Vulnerabilities.ToList().Count == 0 || t.SFHs.ToList().Count == 0)
                    listEmptyThreats.Add(t);
            }

            dgvThreats.DataSource = listEmptyThreats;
            dgvThreats.Columns["ThreatID"].Visible = false;
            dgvThreats.Columns["ThreatSources"].Visible = false;
            dgvThreats.Columns["DateOfChange"].Visible = false;
            dgvThreats.Columns["DateOfAdd"].Visible = false;
            dgvThreats.Columns["ImplementWays"].Visible = false;
            dgvThreats.Columns["SFHs"].Visible = false;
            dgvThreats.Columns["Vulnerabilities"].Visible = false;
            dgvThreats.Columns["Description"].Visible = false;
            dgvThreats.Columns["ObjectOfInfluence"].Visible = false;
            dgvThreats.Columns["GISMeasures"].Visible = false;
            dgvThreats.Columns["ConfidenceViolation"].Visible = false;
            dgvThreats.Columns["IntegrityViolation"].Visible = false;
            dgvThreats.Columns["AvailabilityViolation"].Visible = false;
            dgvThreats.Columns["stringSources"].Visible = false;
            dgvThreats.Columns["ThreatNumber"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvThreats.Columns["ThreatID"].DisplayIndex = 0;
            dgvThreats.Columns["ThreatNumber"].Width = 60;
            dgvThreats.Columns["ThreatNumber"].DisplayIndex = 1;
            dgvThreats.Columns["ThreatNumber"].HeaderText = "№ УБИ";
            dgvThreats.Columns["Name"].HeaderText = "Название УБИ";
            dgvThreats.Columns["Name"].DisplayIndex = 2;            
            dgvThreats.Columns["stringVuls"].HeaderText = "Уязвимости";
            dgvThreats.Columns["stringVuls"].DisplayIndex = 3;
            dgvThreats.Columns["stringWays"].HeaderText = "Способы реализации УБИ";
            dgvThreats.Columns["stringWays"].DisplayIndex = 4;
            dgvThreats.Columns["stringSFHS"].HeaderText = "СФХ";
            dgvThreats.Columns["stringSFHS"].DisplayIndex = 5;
            

            using (KPSZIContext db = new KPSZIContext())
            {
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
                int gbY = 5;

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
                    tpSFHs.Controls.Add(gb);
                    i++;
                }

                // инициализация таблицы уязвимостей
                listVuls = db.Vulnerabilities.ToList();
                dgvVulnerabilities.DataSource = listVuls;

                dgvVulnerabilities.Columns["Threats"].Visible = false;
                dgvVulnerabilities.Columns["VulnerabilityId"].Visible = false;
                dgvVulnerabilities.Columns["VulnerabilityNumber"].Visible = false;
                dgvVulnerabilities.Columns["cbVulnerability"].Width = 30;
                dgvVulnerabilities.Columns["cbVulnerability"].MinimumWidth = 30;
                dgvVulnerabilities.Columns["VulnerabilityName"].Width = 300;
                dgvVulnerabilities.Columns["VulnerabilityName"].HeaderText = "Уязвимость";
                dgvVulnerabilities.Columns["VulnerabilityDescription"].HeaderText = "Описание";

                
                foreach (DataGridViewColumn column in dgvVulnerabilities.Columns)
                    column.ReadOnly = true;
                dgvVulnerabilities.Columns["cbVulnerability"].ReadOnly = false;

                ((DataGridViewCheckBoxColumn)dgvVulnerabilities.Columns["cbVulnerability"]).TrueValue = true;
                ((DataGridViewCheckBoxColumn)dgvVulnerabilities.Columns["cbVulnerability"]).FalseValue = false;

                dgvVulnerabilities.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(234, 240, 255);
                dgvVulnerabilities.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(28, 32, 57);
                dgvVulnerabilities.EnableHeadersVisualStyles = false;

                dgvVulnerabilities.SelectionChanged += new System.EventHandler(dgvVulnerabilities_SelectionChanged);


                // инициализация способов реализации                
                ListImplementWays = db.ImplementWays.OrderBy(iw => iw.WayNumber).ToList();

                int k = 0;
                int yLoc = 5;
                foreach (ImplementWay iw in ListImplementWays)
                {
                    CheckBox cbox = new CheckBox();
                    cbox.Text = iw.WayName;
                    cbox.AutoSize = false;
                    int y = 13 * ((TextRenderer.MeasureText(iw.WayName, clbImplementWays.Font).Width + 50) / (clbImplementWays.Width - 25) + 1) + 10;
                    cbox.Size = new Size(clbImplementWays.Width - 25, y);
                    cbox.Location = new Point(8, yLoc);
                    cbox.CheckedChanged += new System.EventHandler(cbWay_CheckedChanged);
                    yLoc += y;
                    k++;
                    clbImplementWays.Controls.Add(cbox);
                }
                yLoc += 14;
                clbImplementWays.Height = yLoc;
            }
        }

        private void rbSFH_CheckedChanged(object sender, EventArgs e)
        {
            saveChanges();
        }

        private void saveChanges()
        {

        }

        private void dgvVulnerabilities_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvThreats.SelectedCells.Count > 0)
                dgvThreats.SelectedCells[0].Selected = false;
        }

        private void cbWay_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        //private void dgvThreats_SelectionChanged(object sender, EventArgs e)
        //{
        //    if (dgvThreats.SelectedRows.Count == 0)
        //        return;
        //    int threatNumber = Convert.ToInt32(dgvThreats.SelectedRows[0].
        //        Cells[dgvThreats.Columns["ThreatNumber"].Index].Value.ToString());
        //    Console.WriteLine(threatNumber);

        //    using (KPSZIContext db = new KPSZIContext())
        //    {
        //        List<ImplementWay> listImplement = db.Threats.Where(t => t.ThreatNumber == threatNumber).FirstOrDefault().ImplementWays.ToList();
        //        foreach (ImplementWay iw in listImplement)
        //        {
        //            foreach(Control control in clbImplementWays.Controls)
        //            {
        //                if (!(control is CheckBox))
        //                    return;
        //                CheckBox cb = control as CheckBox;
        //                foreach()
        //            }
        //        }
        //    }
        //}
    }
}
