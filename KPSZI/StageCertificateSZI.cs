using DocumentFormat.OpenXml.Spreadsheet;
using KPSZI.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Forms;

namespace KPSZI
{
    class StageCertificateSZI : Stage
    {
        protected override ImageList imageListForTabPage { get; set; }

        public StageCertificateSZI(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS)
            : base(stageTab, stageNode, mainForm, IS)
        {

        }

        public override void enterTabPage()
        {

        }

        public override void saveChanges()
        {
            
        }

        protected override void initTabPage()
        {
            ViewAllCertificateSZI(new Button(), new EventArgs());
            mf.btnViewAllCertificateSZI.Click += new EventHandler(ViewAllCertificateSZI);
            mf.btnSearchCertificateSZI.Click += new EventHandler(SearchCertificateSZI);
            mf.btnAddSZIListToReport.Click += new EventHandler(AddSZIToReport);
        }

        private void SearchCertificateSZI(object sender, EventArgs e)// Поиск сертификатов по заданным параметрам и их вывод
        {
            using (KPSZIContext db = new KPSZIContext())
            {
                List<CertificateSZI> cSZIlist = db.CertificatesSZI.ToList();
                if(!String.IsNullOrEmpty(mf.tbNumCertificateSZI.Text))
                    cSZIlist = cSZIlist.Where(el => el.CertificateNumber == mf.tbNumCertificateSZI.Text).ToList();
                if (!String.IsNullOrEmpty(mf.tbNameSZI.Text))
                {
                    if (mf.cbSZIRegisterConsider.Checked)
                    {
                        cSZIlist = cSZIlist.Where(el => el.NameSZI.Contains(mf.tbNameSZI.Text)).ToList();                       
                    }
                    else
                    {
                        cSZIlist = cSZIlist.Where(el => el.NameSZI.ToLower().Contains(mf.tbNameSZI.Text.ToLower())).ToList();
                    }                   
                }

                FillSZIData(cSZIlist.ToList());
            }
        }

        private void ViewAllCertificateSZI(object sender, EventArgs e)// Отображение всех сертификатов в БД
        {
            using (KPSZIContext db = new KPSZIContext())
            {
                List<CertificateSZI> cSZIlist = db.CertificatesSZI.ToList();
                FillSZIData(cSZIlist);
            }
        }

        private void AddSZIToReport(object sender, EventArgs e)
        {
            
            
            DataGridViewCellCollection cells = mf.dgvCertificateSZI.CurrentRow.Cells;
            string number = cells[0].Value.ToString();
            string name = cells[2].Value.ToString();
            string abilityToUse = cells[4].Value.ToString();

            ListViewItem item = new ListViewItem(number);
            item.SubItems.Add(name);
            item.SubItems.Add(abilityToUse);
            mf.lvReportCertificates.Items.Add(item);

            /*foreach (CertificateSZI szi in cSZIlist)
            {
                string number = szi.CertificateNumber;
                string name = szi.NameSZI;
                string abilityToUse = szi.GetAbilityToUse();

                ListViewItem item = new ListViewItem(number);
                item.SubItems.Add(name);
                item.SubItems.Add(abilityToUse);
                mf.lvSZIDifferences.Items.Add(item);
            }*/
        }

        void FillSZIData(List<CertificateSZI> SZIs)
        {
            DateTime today = DateTime.Today;            
            mf.dgvCertificateSZI.Rows.Clear();

            foreach (CertificateSZI szi in SZIs)
            {
                /*string abilityToUse = "Да";
                DateTime validity = Convert.ToDateTime(szi.Validity);
                DateTime validityTech;

                validityTech = string.IsNullOrEmpty(szi.ValidityTechnicalSupport) ? DateTime.MinValue : Convert.ToDateTime(szi.ValidityTechnicalSupport);

                if (validity < today && validityTech < today)
                {
                    abilityToUse = "Нет";
                }*/
                string abilityToUse = szi.GetAbilityToUse();
                mf.dgvCertificateSZI.Rows.Add(szi.CertificateNumber, szi.Validity, szi.NameSZI, szi.ValidityTechnicalSupport, abilityToUse);
            }
        }
    }
}
