using KPSZI.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPSZI
{
    class StageCertificateSZI : Stage
    {
        /// <summary>
        /// Список сертификатов на вывод
        /// </summary
        List<CertificateSZI> cSZI;
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

            //mf.cbNameSZI.KeyPress += new KeyPressEventHandler(KeyPressInComboBox);
            mf.btnSearchCertificateSZI.Click += new EventHandler(SearchCertificateSZI);
            using (KPSZIContext db = new KPSZIContext())
            {
                cSZI = db.CertificatesSZI.ToList();
                foreach (CertificateSZI szi in cSZI)
                {
                    mf.cbNameSZI.Items.Add(szi.NameSZI);
                }
            }
        }

        private void SearchCertificateSZI(object sender, EventArgs e)// Поиск сертификатов по заданным параметрам и их вывод
        {
            using (KPSZIContext db = new KPSZIContext())
            {
                List<CertificateSZI> cSZIlist = db.CertificatesSZI.ToList();
                if(!String.IsNullOrEmpty(mf.tbNumCertificateSZI.Text))
                    cSZIlist = cSZIlist.Where(el => el.CertificateNumber == mf.tbNumCertificateSZI.Text).ToList();
                if (!String.IsNullOrEmpty(mf.cbNameSZI.Text))
                {
                    if (mf.cbSZIRegisterConsider.Checked)
                    {
                        cSZIlist = cSZIlist.Where(el => el.NameSZI.Contains(mf.cbNameSZI.Text)).ToList();                       
                    }
                    else
                    {
                        cSZIlist = cSZIlist.Where(el => el.NameSZI.ToLower().Contains(mf.cbNameSZI.Text.ToLower())).ToList();
                    }                   
                }

                FillSZIData(cSZIlist.ToList());
            }
        }

        private void ViewAllCertificateSZI(object sender, EventArgs e)// Отображение всех сертификатов в БД
        {
            using (KPSZIContext db = new KPSZIContext())
            {
                FillSZIData(db.CertificatesSZI.ToList());
            }
        }

        void FillSZIData(List<CertificateSZI> SZIs)
        {
            DateTime today = DateTime.Today;            
            mf.dgvCertificateSZI.Rows.Clear();

            foreach (CertificateSZI szi in SZIs)
            {
                string abilityToUse = "Да";
                DateTime validity = Convert.ToDateTime(szi.Validity);
                DateTime validityTech;

                validityTech = string.IsNullOrEmpty(szi.ValidityTechnicalSupport) ? DateTime.MinValue : Convert.ToDateTime(szi.ValidityTechnicalSupport);

                if (validity < today && validityTech < today)
                {
                    abilityToUse = "Нет";
                }

                mf.dgvCertificateSZI.Rows.Add(szi.CertificateNumber, szi.Validity, szi.NameSZI, szi.ValidityTechnicalSupport, abilityToUse);
            }
        }

        //private void KeyPressInComboBox(object sender, KeyPressEventArgs e)
        //{
        //    using (KPSZIContext db = new KPSZIContext())
        //    {
        //        List<CertificateSZI> cSZIlist = db.CertificatesSZI.Where(el => el.NameSZI.ToLower().Contains(mf.cbNameSZI.Text.ToLower()) == true).ToList();
        //        if (cSZIlist.Count != 0)
        //        {
        //            mf.cbNameSZI.Items.Clear();
        //        }
        //        else
        //        {
        //            mf.cbNameSZI.Items.Clear();
        //            mf.cbNameSZI.Items.Add("Ничего не найдено");
        //        }
                    

        //        foreach (CertificateSZI szi in cSZIlist)
        //        {
        //            mf.cbNameSZI.Items.Add(szi.NameSZI);
        //        }

        //        //object[] items =  mf.cbNameSZI.Items.OfType<String>().Distinct().ToArray();
        //        //if (cSZIlist.Count != 0)
        //        //{
        //        //    mf.cbNameSZI.Items.Clear();
        //        //}
        //        //else
        //        //{
        //        //    mf.cbNameSZI.Items.Clear();
        //        //    mf.cbNameSZI.Items.Add("Ничего не найдено");
        //        //}
        //        //mf.cbNameSZI.Items.AddRange(items);
        //        mf.cbNameSZI.SelectionStart = mf.cbNameSZI.Text.Length;
        //        if (mf.cbNameSZI.Items.Count > 0)
        //            mf.cbNameSZI.DroppedDown = true;
                
        //    }
        //}
    }
}
